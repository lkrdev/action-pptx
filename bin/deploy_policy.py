"""Pure GCP Org Policy parsing and Cloud Run deploy-mode assessment helpers."""

from __future__ import annotations

from dataclasses import dataclass

VALID_DEPLOY_MODES = ("load-balanced", "public")
VALID_LB_TYPES = ("external", "internal", "none")


@dataclass(frozen=True)
class OrgPolicyReport:
    """Effective GCP Org Policy status relevant to Cloud Run exposure."""

    require_invoker_iam: bool | None
    ingress_all_allowed: bool | None
    ingress_lb_allowed: bool | None
    allowed_member_domains_restricted: bool | None
    notes: tuple[str, ...]


@dataclass(frozen=True)
class ModeAssessment:
    """Compatibility and recommendation status for a deployment mode."""

    mode: str
    status: str
    summary: str
    is_default: bool = False


def _parse_rules_list(payload: dict[str, object]) -> list[dict[str, object]]:
    spec = payload.get("spec")
    if not isinstance(spec, dict):
        return []
    rules = spec.get("rules")
    return [r for r in rules if isinstance(r, dict)] if isinstance(rules, list) else []


def parse_org_policy_constraint(
    constraint: str,
    payload: dict[str, object] | None,
) -> dict[str, bool | None]:
    """Extract normalized boolean flags from an effective org-policy JSON document."""
    if payload is None:
        return {"enforced": None, "allow_all": None, "lb_allowed": None, "restricted": None}

    rules = _parse_rules_list(payload)
    if constraint.endswith("run.managed.requireInvokerIam"):
        return {"enforced": any(rule.get("enforce") is True for rule in rules)}

    if constraint.endswith("run.allowedIngress"):
        if not rules:
            return {"allow_all": True, "lb_allowed": True}
        allow_all = False
        lb_allowed = False
        has_values = False
        for rule in rules:
            if rule.get("denyAll") is True:
                return {"allow_all": False, "lb_allowed": False}
            if rule.get("allowAll") is True:
                allow_all = True
                lb_allowed = True
            values = rule.get("values")
            if isinstance(values, dict):
                has_values = True
                allowed = [str(v).lower() for v in (values.get("allowedValues") or [])]
                denied = [str(v).lower() for v in (values.get("deniedValues") or [])]
                if allowed:
                    allow_all = allow_all or any(
                        v in ("all", "ingress_traffic_all") for v in allowed
                    )
                    lb_allowed = lb_allowed or any(
                        "load-balancing" in v or "load_balancing" in v for v in allowed
                    )
                elif denied:
                    allow_all = not any(v in ("all", "ingress_traffic_all") for v in denied)
                    lb_allowed = not any(
                        "load-balancing" in v or "load_balancing" in v for v in denied
                    )
        if not has_values and not allow_all and not lb_allowed:
            return {"allow_all": True, "lb_allowed": True}
        return {"allow_all": allow_all, "lb_allowed": lb_allowed}

    if constraint.endswith("iam.allowedPolicyMemberDomains"):
        restricted = any(
            not rule.get("allowAll")
            and (
                rule.get("denyAll") is True
                or (
                    isinstance(rule.get("values"), dict)
                    and bool(rule["values"].get("allowedValues"))
                )
            )
            for rule in rules
        )
        return {"restricted": restricted}

    return {}


def assess_deploy_modes(report: OrgPolicyReport) -> list[ModeAssessment]:
    """Rank deployment modes (Actions + Extension) from most secure to simplest."""
    iam_req = report.require_invoker_iam is True
    ingress_all_blocked = report.ingress_all_allowed is False
    lb_blocked = report.ingress_lb_allowed is False

    if iam_req:
        lb_status = "REQUIRES PROXY"
        lb_summary = (
            "run.managed.requireInvokerIam blocks --no-invoker-iam-check; "
            "requires an authenticating proxy/gateway in front of Cloud Run."
        )
    elif lb_blocked:
        lb_status = "BLOCKED"
        lb_summary = "Blocked by run.allowedIngress policy."
    else:
        lb_status = "RECOMMENDED"
        lb_summary = (
            "Most secure for BOTH Actions + Extension: blocks direct .a.run.app access "
            "and fronts Cloud Run with an External or Internal HTTPS Load Balancer."
        )

    if iam_req or ingress_all_blocked:
        reason = (
            "run.managed.requireInvokerIam"
            if iam_req
            else "run.allowedIngress restricts public ingress"
        )
        pub_status = "BLOCKED"
        pub_summary = f"Blocked by Org Policy ({reason})."
    else:
        pub_status = "AVAILABLE"
        pub_summary = (
            "Simplest setup for BOTH Actions + Extension: direct Cloud Run URL "
            "(app still enforces ACTION_HUB_SECRET / Google OAuth)."
        )

    return [
        ModeAssessment(
            mode="load-balanced",
            status=lb_status,
            summary=lb_summary,
            is_default=True,
        ),
        ModeAssessment(
            mode="public",
            status=pub_status,
            summary=pub_summary,
            is_default=False,
        ),
    ]

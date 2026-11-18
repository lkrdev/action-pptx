from __future__ import annotations

from bin.deploy_policy import (
    OrgPolicyReport,
    assess_deploy_modes,
    parse_org_policy_constraint,
)


def test_parse_org_policy_constraint_flags() -> None:
    invoker = parse_org_policy_constraint(
        "constraints/run.managed.requireInvokerIam",
        {"spec": {"rules": [{"enforce": True}]}},
    )
    assert invoker["enforced"] is True

    ingress = parse_org_policy_constraint(
        "constraints/run.allowedIngress",
        {"spec": {"rules": [{"values": {"allowedValues": ["internal-and-cloud-load-balancing"]}}]}},
    )
    assert ingress["allow_all"] is False
    assert ingress["lb_allowed"] is True

    domains = parse_org_policy_constraint(
        "constraints/iam.allowedPolicyMemberDomains",
        {"spec": {"rules": [{"values": {"allowedValues": ["C012345"]}}]}},
    )
    assert domains["restricted"] is True


def test_assess_deploy_modes_defaults_to_load_balanced_when_allowed() -> None:
    report = OrgPolicyReport(
        require_invoker_iam=False,
        ingress_all_allowed=True,
        ingress_lb_allowed=True,
        allowed_member_domains_restricted=True,
        notes=(),
    )
    assessments = assess_deploy_modes(report)
    assert [a.mode for a in assessments] == ["load-balanced", "public"]
    assert assessments[0].status == "RECOMMENDED"
    assert assessments[0].is_default is True
    assert assessments[1].status == "AVAILABLE"


def test_assess_deploy_modes_flags_proxy_needed_when_invoker_iam_enforced() -> None:
    report = OrgPolicyReport(
        require_invoker_iam=True,
        ingress_all_allowed=False,
        ingress_lb_allowed=True,
        allowed_member_domains_restricted=True,
        notes=(),
    )
    assessments = assess_deploy_modes(report)
    by_mode = {a.mode: a for a in assessments}
    assert by_mode["load-balanced"].status == "REQUIRES PROXY"
    assert by_mode["load-balanced"].is_default is True
    assert by_mode["public"].status == "BLOCKED"

from __future__ import annotations

import ast
import re
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
TARGET_DIRS = ("src", "bin", "tests")


def _has_plain_dict(annotation: ast.AST) -> bool:
    subscripted = {id(n.value) for n in ast.walk(annotation) if isinstance(n, ast.Subscript)}
    return any(
        isinstance(n, ast.Name) and n.id in ("dict", "Dict") and id(n) not in subscripted
        for n in ast.walk(annotation)
    )


def _iter_annotations(tree: ast.AST) -> list[tuple[int, ast.AST]]:
    out: list[tuple[int, ast.AST]] = []
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            for arg in (
                *node.args.posonlyargs,
                *node.args.args,
                *node.args.kwonlyargs,
                node.args.vararg,
                node.args.kwarg,
            ):
                if arg and arg.annotation:
                    out.append((arg.lineno, arg.annotation))
            if node.returns:
                out.append((node.lineno, node.returns))
        elif isinstance(node, ast.AnnAssign) and node.annotation:
            out.append((node.lineno, node.annotation))
        elif isinstance(node, ast.TypeAlias):
            out.append((node.lineno, node.value))
    return out


def test_no_plain_dict_annotations() -> None:
    violations = [
        f"{py_file.relative_to(REPO_ROOT)}:{line}: plain dict in '{ast.unparse(ann)}'"
        for folder in TARGET_DIRS
        if (REPO_ROOT / folder).exists()
        for py_file in sorted((REPO_ROOT / folder).rglob("*.py"))
        for line, ann in _iter_annotations(
            ast.parse(py_file.read_text(encoding="utf-8"), filename=str(py_file))
        )
        if _has_plain_dict(ann)
    ]
    assert not violations, (
        "Plain unparameterized 'dict' annotations are disallowed; "
        "use TypedDict, BaseModel, or dict[K, V]:\n" + "\n".join(violations)
    )


def test_no_internal_artifact_foundry_in_uv_lock() -> None:
    """Verify that uv.lock does not contain corporate internal mirror references."""
    lock_file = REPO_ROOT / "uv.lock"
    assert lock_file.exists(), "uv.lock not found"
    content = lock_file.read_text(encoding="utf-8")
    assert "artifact-foundry-prod" not in content and "airlock-proxy.uplink" not in content, (
        "uv.lock contains references to internal corporate mirror! "
        "Run `uv lock --default-index https://pypi.org/simple --no-config` "
        "to re-lock against public PyPI."
    )


def test_readme_manifest_matches_manifest_lkml() -> None:
    """Verify that the ```lookml block in README.md exactly matches manifest.lkml."""
    manifest_file = REPO_ROOT / "manifest.lkml"
    readme_file = REPO_ROOT / "README.md"
    example_manifest = REPO_ROOT / "assets" / "lookml" / "manifest.lkml"
    assert manifest_file.exists(), "manifest.lkml not found at repository root"
    assert readme_file.exists(), "README.md not found at repository root"

    manifest_text = manifest_file.read_text(encoding="utf-8").strip()
    readme_text = readme_file.read_text(encoding="utf-8")
    match = re.search(r"```lookml\n(.*?)\n```", readme_text, re.DOTALL)
    assert match is not None, "No ```lookml ... ``` code block found in README.md"
    readme_manifest = match.group(1).strip()
    assert readme_manifest == manifest_text, (
        "The ```lookml manifest block in README.md does not match manifest.lkml!"
    )
    if example_manifest.exists():
        assert manifest_text in example_manifest.read_text(encoding="utf-8"), (
            "assets/lookml/manifest.lkml does not contain the root manifest.lkml block!"
        )

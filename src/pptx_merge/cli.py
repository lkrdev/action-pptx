from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from pptx_merge.engine import MergeError, _unpack_csv_zip, merge_pptx
from pptx_merge.inspector import TileSchema, inspect_and_validate


def _load_schemas_from_json(path: Path) -> list[TileSchema]:
    raw = json.loads(path.read_text(encoding="utf-8"))
    items = raw.get("dashboard_tiles", raw) if isinstance(raw, dict) else raw
    return [
        TileSchema(
            title=item.get("title", item.get("slug", "")),
            slug=item.get("slug", ""),
            columns=list(item.get("columns", [])),
        )
        for item in items
    ]


def _load_schemas_from_zip(path: Path) -> list[TileSchema]:
    tiles, _, _ = _unpack_csv_zip(path.read_bytes())
    return [TileSchema(title=slug, slug=slug, columns=t.headers) for slug, t in tiles.items()]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="pptx-merge",
        description="Inspect, validate, and populate .pptx templates from Looker csv_zip.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p_inspect = sub.add_parser("inspect", help="Inspect a .pptx deck for bindings.")
    p_inspect.add_argument("deck", type=Path, help="Path to template .pptx file")

    p_val = sub.add_parser("validate", help="Validate a .pptx deck against schema or csv_zip.")
    p_val.add_argument("deck", type=Path, help="Path to template .pptx file")
    group = p_val.add_mutually_exclusive_group(required=True)
    group.add_argument("--schema", type=Path, help="Path to dashboard tiles JSON file")
    group.add_argument("--csv-zip", type=Path, help="Path to Looker csv_zip export")

    p_merge = sub.add_parser("merge", help="Merge a Looker csv_zip export into a .pptx template.")
    p_merge.add_argument("deck", type=Path, help="Path to template .pptx file")
    p_merge.add_argument("--csv-zip", type=Path, required=True, help="Path to Looker csv_zip")
    p_merge.add_argument("-o", "--output", type=Path, required=True, help="Output .pptx file")
    p_merge.add_argument("--title", default="", help="Optional scheduled plan title (_plan.title)")
    p_merge.add_argument("--url", default="", help="Optional scheduled plan URL (_plan.url)")
    p_merge.add_argument("--run-at", default="", help="Optional run timestamp (_plan.run_at)")
    p_merge.add_argument("--theme", default="", help="Optional dynamic theme/brand name")

    args = parser.parse_args(argv)

    try:
        pptx_bytes = args.deck.read_bytes()
    except OSError as exc:
        print(json.dumps({"valid": False, "error": str(exc)}), file=sys.stderr)
        return 1

    if args.command == "inspect":
        report = inspect_and_validate(pptx_bytes)
        print(json.dumps(report.to_dict(), indent=2))
        return 0 if report.valid else 1

    if args.command == "validate":
        try:
            schemas = (
                _load_schemas_from_json(args.schema)
                if args.schema
                else _load_schemas_from_zip(args.csv_zip)
            )
        except Exception as exc:
            print(json.dumps({"valid": False, "error": str(exc)}), file=sys.stderr)
            return 1

        report = inspect_and_validate(pptx_bytes, dashboard_tiles=schemas)
        print(json.dumps(report.to_dict(), indent=2))
        return 0 if report.valid else 1

    if args.command == "merge":
        try:
            result = merge_pptx(
                pptx_bytes,
                args.csv_zip.read_bytes(),
                plan_meta={"title": args.title, "url": args.url, "run_at": args.run_at},
                theme=args.theme or None,
            )
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_bytes(result.pptx_bytes)
        except (OSError, MergeError) as exc:
            print(json.dumps({"success": False, "error": str(exc)}), file=sys.stderr)
            return 1

        print(
            json.dumps(
                {
                    "success": True,
                    "output": str(args.output),
                    "slides_processed": result.slides_processed,
                    "charts_updated": result.charts_updated,
                    "tables_updated": result.tables_updated,
                    "text_tags_replaced": result.text_tags_replaced,
                    "tile_row_counts": result.tile_row_counts,
                },
                indent=2,
            )
        )
        return 0

    return 1


if __name__ == "__main__":
    raise SystemExit(main())

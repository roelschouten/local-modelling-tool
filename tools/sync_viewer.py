#!/usr/bin/env python3
"""Keep the Databricks App's viewer identical to the authored copy in docs/.

docs/index.html is the ONLY file anyone edits. The Databricks App serves its own
copy from model-viewer/model-viewer-app/public/, and the two used to be edited
independently, which is how they drift.

Fork addition: docs/static/ (scripts and example models loaded by index.html)
is mirrored to model-viewer/model-viewer-app/public/static/ the same way.

    python3 tools/sync_viewer.py            copy docs/ -> app/, report the result
    python3 tools/sync_viewer.py --check    exit non-zero if they differ

The --check form runs in CI so a hand-edit of the app copy fails the build
instead of quietly surviving.
"""

from __future__ import annotations

import argparse
import filecmp
import hashlib
import shutil
import sys
from pathlib import Path

SOURCE = Path("docs/index.html")
TARGET = Path("model-viewer/model-viewer-app/public/index.html")
STATIC_SOURCE = Path("docs/static")
STATIC_TARGET = Path("model-viewer/model-viewer-app/public/static")


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()[:12]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--repo",
        type=Path,
        default=Path(__file__).resolve().parent.parent,
        help="repository root (default: parent of tools/)",
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="verify only; do not write. Exits 1 when the copies differ.",
    )
    args = parser.parse_args()

    pairs = [(SOURCE, TARGET)]
    static_src = args.repo / STATIC_SOURCE
    if static_src.is_dir():
        for f in sorted(static_src.rglob("*")):
            if f.is_file():
                rel = f.relative_to(static_src)
                pairs.append((STATIC_SOURCE / rel, STATIC_TARGET / rel))
    extra = []
    static_tgt = args.repo / STATIC_TARGET
    if static_tgt.is_dir():
        wanted = {args.repo / t for _, t in pairs}
        extra = [f for f in sorted(static_tgt.rglob("*")) if f.is_file() and f not in wanted]

    if not (args.repo / SOURCE).is_file():
        print(f"error: {SOURCE} is missing; it is the authored copy", file=sys.stderr)
        return 1

    if args.check:
        bad = 0
        for src, tgt in pairs:
            source, target = args.repo / src, args.repo / tgt
            if not target.is_file():
                print(f"error: {tgt} is missing; run tools/sync_viewer.py", file=sys.stderr)
                bad += 1
            elif not filecmp.cmp(source, target, shallow=False):
                print(
                    f"error: {tgt} ({digest(target)}) differs from the authored copy "
                    f"{src} ({digest(source)}).\n"
                    f"       Edit {src} only, then run: python3 tools/sync_viewer.py",
                    file=sys.stderr,
                )
                bad += 1
        for f in extra:
            print(f"error: {f.relative_to(args.repo)} has no authored copy under {STATIC_SOURCE}", file=sys.stderr)
            bad += 1
        if bad:
            return 1
        print(f"in sync: {len(pairs)} file(s), {SOURCE} {digest(args.repo / SOURCE)}")
        return 0

    for src, tgt in pairs:
        source, target = args.repo / src, args.repo / tgt
        target.parent.mkdir(parents=True, exist_ok=True)
        if target.is_file() and filecmp.cmp(source, target, shallow=False):
            print(f"already in sync: {src} ({digest(source)})")
            continue
        shutil.copyfile(source, target)
        print(f"synced {src} -> {tgt} ({digest(source)})")
    for f in extra:
        f.unlink()
        print(f"removed {f.relative_to(args.repo)} (no authored copy)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

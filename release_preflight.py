from __future__ import annotations

import argparse
import shlex
import subprocess
import sys
from pathlib import Path
from typing import List


DEFAULT_SAMPLE = "浦港养护2026年3月8日巡查问题处置日报.xlsx"


def _quote(parts: List[str]) -> str:
    return " ".join(shlex.quote(p) for p in parts)


def _run(cmd: List[str], cwd: Path) -> int:
    print(f"[RUN] {_quote(cmd)}")
    completed = subprocess.run(cmd, cwd=cwd)
    return completed.returncode


def parse_args(argv: List[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Release preflight checks before packaging customer build."
    )
    parser.add_argument(
        "--sample",
        type=Path,
        default=Path(DEFAULT_SAMPLE),
        help=f"Regression Excel sample path. Default: {DEFAULT_SAMPLE}",
    )
    parser.add_argument(
        "--samples",
        type=int,
        default=3,
        help="How many parsed rows to print in offline smoke check.",
    )
    parser.add_argument(
        "--keep-session",
        action="store_true",
        help="Keep temporary extraction session for manual inspection.",
    )
    return parser.parse_args(argv)


def main(argv: List[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    repo_root = Path(__file__).resolve().parent

    sample_path = args.sample
    if not sample_path.is_absolute():
        sample_path = (repo_root / sample_path).resolve()

    if not sample_path.exists():
        print(f"[FAIL] Regression sample not found: {sample_path}")
        print("       Add the sample file first, then rerun preflight.")
        return 2

    print("[INFO] Release preflight started")
    print(f"[INFO] Using sample: {sample_path}")

    test_cmd = [sys.executable, "-m", "unittest", "discover", "-s", "tests", "-v"]
    if _run(test_cmd, cwd=repo_root) != 0:
        print("[FAIL] Unit tests failed")
        return 1

    smoke_cmd = [
        sys.executable,
        "offline_smoke_check.py",
        str(sample_path),
        "--samples",
        str(args.samples),
    ]
    if args.keep_session:
        smoke_cmd.append("--keep-session")

    if _run(smoke_cmd, cwd=repo_root) != 0:
        print("[FAIL] Offline smoke check failed")
        return 1

    print("[PASS] Release preflight passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

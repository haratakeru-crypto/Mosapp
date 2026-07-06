"""Backward-compatible entry point; delegates to build_variant_artifacts.py."""
import subprocess
import sys
from pathlib import Path

script = Path(__file__).resolve().parent / "build_variant_artifacts.py"
args = [sys.executable, str(script)]
if "--validate-only" in sys.argv:
    args.append("--validate-only")
args.append("--all")
raise SystemExit(subprocess.call(args))

import os
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]
os.chdir(ROOT)
os.environ.setdefault("TORCH_HOME", str(ROOT / ".torch"))


def run(*args):
    subprocess.run([sys.executable, *args], check=True)


if __name__ == "__main__":
    if not Path("results/manifest.csv").exists():
        if not Path("data/source.json").exists():
            run("scripts/download_data.py")
        run("-m", "src.dataset")
    run("scripts/audit_exclusions.py")
    for name in ("resnet50", "densenet121"):
        if not Path(f"logs/run_{name}.json").exists():
            run("-m", "src.train", "--config", f"configs/{name}.yaml")
        run("-m", "src.evaluate", "--checkpoint", f"checkpoints/best_{name}.pth")
    run("-m", "src.report")
    run("scripts/verify_results.py")

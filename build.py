import shutil
import subprocess
import sys
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
MAIN = SRC / "Main.py"
REQUIREMENTS = ROOT / "requirements.txt"
APP_NAME = "BeamOptimizer"


def rm(path: Path):
    if path.is_dir():
        shutil.rmtree(path)
        print(f"Deleted: {path}")
    elif path.exists():
        path.unlink()
        print(f"Deleted: {path}")


def main() -> int:
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "-r", str(REQUIREMENTS)],
            cwd=ROOT,
        )

        rm(ROOT / "build")
        rm(ROOT / "dist")

        subprocess.check_call(
            [
                sys.executable, "-m", "PyInstaller",
                "--onefile",
                "--windowed",
                "--noconfirm",
                "--name", APP_NAME,
                "--paths", str(SRC),
                str(MAIN),
            ],
            cwd=ROOT,
        )

        print(f"\nBuild complete: {ROOT / 'dist' / (APP_NAME + '.exe')}")
        return 0
    except Exception:
        traceback.print_exc()
        return 1
    finally:
        rm(ROOT / "build")
        for spec in ROOT.glob("*.spec"):
            rm(spec)


if __name__ == "__main__":
    code = main()
    input("Press Enter to close ....")
    sys.exit(code)
from __future__ import annotations

import subprocess
import sys
import venv
from pathlib import Path


ROOT = Path(__file__).resolve().parentif ROOT.name == 'src':    ROOT = ROOT.parent
VENV_DIR = ROOT / ".venv"
SCRIPTS_DIR = "Scripts" if sys.platform.startswith("win") else "bin"
PYTHON_EXE = VENV_DIR / SCRIPTS_DIR / ("python.exe" if sys.platform.startswith("win") else "python")
REQUIREMENTS_TXT = ROOT / "requirements.txt"


DEFAULT_PACKAGES = [
    "pyedb>=0.6.0",
    "PySide6>=6.5",
]


def info(message: str) -> None:
    print(f"[info] {message}")


def warn(message: str) -> None:
    print(f"[warn] {message}")


def run(args: list[str]) -> None:
    info("Running: " + " ".join(args))
    subprocess.check_call(args)


def ensure_venv() -> None:
    if PYTHON_EXE.exists():
        info(f"Virtual environment already present at {VENV_DIR}")
        return

    info(f"Creating virtual environment at {VENV_DIR}")
    builder = venv.EnvBuilder(with_pip=True, clear=False, upgrade=False)
    builder.create(str(VENV_DIR))


def install_packages() -> None:
    if not PYTHON_EXE.exists():
        raise RuntimeError("Virtual environment Python interpreter not found after creation.")

    run([str(PYTHON_EXE), "-m", "pip", "install", "--upgrade", "pip", "setuptools", "wheel"])

    if REQUIREMENTS_TXT.exists():
        info("Installing packages from requirements.txt")
        run([str(PYTHON_EXE), "-m", "pip", "install", "-r", str(REQUIREMENTS_TXT)])
    else:
        info("requirements.txt not found; installing default dependencies")
        run([str(PYTHON_EXE), "-m", "pip", "install", *DEFAULT_PACKAGES])


def main() -> int:
    try:
        ensure_venv()
        install_packages()
    except subprocess.CalledProcessError as exc:
        warn(f"Command failed with exit code {exc.returncode}")
        return exc.returncode or 1
    except Exception as exc:  # pragma: no cover - bootstrap diagnostics
        warn(str(exc))
        return 1

    info("Environment ready. Rerun run.bat to launch the GUI.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

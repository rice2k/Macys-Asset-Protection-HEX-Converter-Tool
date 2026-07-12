from __future__ import annotations

import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PATCH_DIR = ROOT / "patches" / "v1.0.4"
PATCH_OUT = ROOT / ".build" / "v1.0.4-combined.patch"


def main() -> None:
    parts = sorted(PATCH_DIR.glob("part*.patch"))
    if len(parts) != 7:
        raise SystemExit(f"Expected 7 v1.0.4 patch parts, found {len(parts)}")

    patch_text = "".join(path.read_text(encoding="utf-8") for path in parts)
    patch_text = patch_text.replace(
        "--- /mnt/data/desktop_app.py\t2026-07-12 21:55:41.358888802 +0000",
        "--- a/desktop_app.py",
        1,
    ).replace(
        "+++ /mnt/data/AP_HEX_Converter_Tool_v1.0.4_Windows_Build/desktop_app.py\t2026-07-12 21:58:36.081880965 +0000",
        "+++ b/desktop_app.py",
        1,
    )

    PATCH_OUT.parent.mkdir(parents=True, exist_ok=True)
    PATCH_OUT.write_text(patch_text, encoding="utf-8")

    check = subprocess.run(
        ["git", "apply", "--check", "--whitespace=nowarn", str(PATCH_OUT)],
        cwd=ROOT,
        text=True,
        capture_output=True,
    )
    if check.returncode != 0:
        raise SystemExit(f"v1.0.4 patch check failed:\n{check.stdout}\n{check.stderr}")

    subprocess.run(
        ["git", "apply", "--whitespace=nowarn", str(PATCH_OUT)],
        cwd=ROOT,
        check=True,
    )
    print("Applied AP HEX Converter Tool v1.0.4 source update.")


if __name__ == "__main__":
    main()

from __future__ import annotations

import json
import subprocess
import tempfile
from pathlib import Path
from typing import Any, Dict, List


ROOT_DIR = Path(__file__).resolve().parents[2]
DEFAULT_XERIFE_ROOT = ROOT_DIR.parent / "almoxerifado-erp"
DEFAULT_BRIDGE_SCRIPT = DEFAULT_XERIFE_ROOT / "scripts" / "import-xerife-stock-batch.mjs"


def run_xerife_stock_batch(*, connection: Dict[str, Any], items: List[Dict[str, Any]], run_id: str) -> Dict[str, Any]:
    if not DEFAULT_BRIDGE_SCRIPT.exists():
        raise FileNotFoundError(f"Bridge script not found: {DEFAULT_BRIDGE_SCRIPT}")

    payload = {
        "run_id": run_id,
        "connection": connection,
        "items": items,
    }

    temp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as handle:
            json.dump(payload, handle, ensure_ascii=True)
            temp_path = Path(handle.name)

        completed = subprocess.run(
            ["node", str(DEFAULT_BRIDGE_SCRIPT), str(temp_path)],
            cwd=DEFAULT_XERIFE_ROOT,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )

        output = (completed.stdout or "").strip()
        if not output:
            stderr = (completed.stderr or "").strip()
            raise RuntimeError(stderr or "The Xerife bridge did not return any output.")

        data = json.loads(output)
        if completed.returncode != 0 or not data.get("success", False):
            message = data.get("error") or (completed.stderr or "").strip() or "Xerife import failed."
            raise RuntimeError(message)
        return data
    finally:
        if temp_path is not None and temp_path.exists():
            temp_path.unlink(missing_ok=True)

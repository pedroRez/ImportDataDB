from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from subprocess import CompletedProcess
from unittest.mock import patch

from src.core import xerife_bridge


class XerifeBridgeTests(unittest.TestCase):
    def test_prefers_result_file_over_noisy_stdout(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            fake_script = root / "import-xerife-stock-batch.mjs"
            fake_script.write_text("// fake bridge", encoding="utf-8")
            expected = {
                "success": True,
                "created": 2,
                "updated": 1,
                "adjusted": 0,
                "items": [],
            }

            def fake_run(args, **_kwargs):
                result_path = Path(args[3])
                result_path.write_text(json.dumps(expected), encoding="utf-8")
                return CompletedProcess(args=args, returncode=0, stdout="2026-05-03 [info] noisy log\n", stderr="")

            with patch.object(xerife_bridge, "DEFAULT_XERIFE_ROOT", root), patch.object(
                xerife_bridge,
                "DEFAULT_BRIDGE_SCRIPT",
                fake_script,
            ), patch("src.core.xerife_bridge.subprocess.run", side_effect=fake_run):
                result = xerife_bridge.run_xerife_stock_batch(
                    connection={"host": "localhost", "port": 5432, "database": "demo", "user": "demo", "password": "pw"},
                    items=[{"codigo_peca": "A1"}],
                    run_id="abc123",
                )

        self.assertTrue(result["success"])
        self.assertEqual(result["created"], 2)
        self.assertEqual(result["updated"], 1)


if __name__ == "__main__":
    unittest.main()

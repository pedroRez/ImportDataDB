from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from src.core import profiles


class ProfileStorageTests(unittest.TestCase):
    def test_ensure_profile_dir_copies_missing_bundled_profiles(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            base_dir = Path(temp_dir)
            bundled_dir = base_dir / "bundled"
            target_dir = base_dir / "user"
            bundled_dir.mkdir(parents=True, exist_ok=True)
            sample_payload = {
                "id": "demo",
                "name": "Demo",
                "target_type": "xerife_stock",
                "sheet_name": "Planilha",
                "header_row": 1,
            }
            (bundled_dir / "demo.json").write_text(json.dumps(sample_payload), encoding="utf-8")

            with patch.object(profiles, "BUNDLED_PROFILE_DIR", bundled_dir), patch.object(profiles, "PROFILE_DIR", target_dir):
                created_dir = profiles.ensure_profile_dir()

            self.assertEqual(created_dir, target_dir)
            self.assertTrue((target_dir / "demo.json").exists())

    def test_ensure_profile_dir_preserves_existing_user_profile(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            base_dir = Path(temp_dir)
            bundled_dir = base_dir / "bundled"
            target_dir = base_dir / "user"
            bundled_dir.mkdir(parents=True, exist_ok=True)
            target_dir.mkdir(parents=True, exist_ok=True)

            (bundled_dir / "demo.json").write_text('{"name": "Bundled"}', encoding="utf-8")
            target_profile = target_dir / "demo.json"
            target_profile.write_text('{"name": "User"}', encoding="utf-8")

            with patch.object(profiles, "BUNDLED_PROFILE_DIR", bundled_dir), patch.object(profiles, "PROFILE_DIR", target_dir):
                profiles.ensure_profile_dir()

            self.assertEqual(target_profile.read_text(encoding="utf-8"), '{"name": "User"}')


if __name__ == "__main__":
    unittest.main()

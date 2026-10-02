import json
import tempfile
import unittest
from pathlib import Path

from agent.agent import collect_installed_software_entries, normalize_inventory_record, resolve_config_path


class AgentInventoryTests(unittest.TestCase):
    def test_normalize_inventory_record_keeps_valid_values(self):
        record = normalize_inventory_record(
            {
                "computer_name": "PC-01",
                "manufacturer": "Dell",
                "model": "OptiPlex",
                "serial_number": "SN-123",
                "windows": "Windows 11",
                "cpu": {"name": "Intel"},
                "ram": {"total_gb": 32},
                "installed_software": [
                    {"name": "App A", "version": "1.0"},
                    {"name": "App A", "version": "1.0"},
                    {"name": "App B", "version": "2.0"},
                ],
            }
        )

        self.assertEqual(record["manufacturer"], "Dell")
        self.assertEqual(record["model"], "OptiPlex")
        self.assertEqual(record["serial_number"], "SN-123")
        self.assertEqual(record["computer_name"], "PC-01")
        self.assertEqual(len(record["installed_software"]), 2)

    def test_collect_installed_software_entries_removes_empty_and_duplicates(self):
        entries = [
            {"name": "App A", "version": "1.0", "publisher": "Vendor"},
            {"name": "App A", "version": "1.0", "publisher": "Vendor"},
            {"name": "", "version": "", "publisher": ""},
            {"name": "App B", "version": None, "publisher": "Vendor"},
        ]

        result = collect_installed_software_entries(entries)
        self.assertEqual(len(result), 2)
        self.assertEqual(result[0]["name"], "App A")
        self.assertEqual(result[1]["name"], "App B")

    def test_resolve_config_path_prefers_explicit_environment(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            config_dir = Path(temp_dir)
            config_file = config_dir / "custom-config.json"
            config_file.write_text(json.dumps({"server_url": "http://example.local:5000"}), encoding="utf-8")

            resolved = resolve_config_path(config_dir, override_path=str(config_file))
            self.assertEqual(resolved, config_file)


if __name__ == "__main__":
    unittest.main()

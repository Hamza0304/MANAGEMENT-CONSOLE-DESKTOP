import unittest
from unittest.mock import patch

from services.device_enrichment_service import (
    _summarize_enrichment_status,
    collect_device_enrichment,
)


class DeviceEnrichmentServiceTests(unittest.TestCase):

    def test_summarize_enrichment_status_access_denied(self):
        self.assertEqual(
            _summarize_enrichment_status({"access_status": "Access Denied"}),
            "ACCESS_DENIED",
        )

    def test_summarize_enrichment_status_partial_when_some_fields_exist(self):
        self.assertEqual(
            _summarize_enrichment_status(
                {
                    "access_status": "Access Denied",
                    "device_type": "Windows PC",
                    "installed_software": [{"application": "Example App"}],
                }
            ),
            "PARTIAL",
        )

    def test_collect_device_enrichment_produces_partial_when_port_135_unavailable(self):
        with patch(
            "services.device_enrichment_service.probe_remote_tcp_port",
            return_value=(False, "Port probe failed"),
        ), patch(
            "services.device_enrichment_service.get_local_installed_software",
            return_value=[],
        ) as local_software:
            result = collect_device_enrichment(
                "172.29.216.82",
                "AA-BB-CC-DD-EE-FF",
                "Host01",
            )

        self.assertEqual(result["device_type"], "Network Device")
        self.assertEqual(result["access_status"], "Port Unavailable")
        self.assertEqual(result["enrichment_status"], "PARTIAL")
        self.assertEqual(result["enrichment_details"]["wmi"]["status"], "SKIPPED")
        local_software.assert_not_called()

    def test_collect_device_enrichment_preserves_discovery_data_when_wmi_is_denied(self):
        with patch(
            "services.device_enrichment_service.probe_remote_tcp_port",
            return_value=(True, "TRUE"),
        ), patch(
            "services.device_enrichment_service.query_remote_windows_enrichment",
            return_value={
                "device_type": "Windows PC",
                "access_status": "Access Denied",
                "enrichment_status": "ACCESS_DENIED",
                "enrichment_details": {
                    "wmi": {
                        "status": "ACCESS_DENIED",
                        "reason": "Remote WMI/CIM access was denied.",
                    }
                },
            },
        ), patch(
            "services.device_enrichment_service.get_remote_installed_software",
            return_value=[],
        ):
            result = collect_device_enrichment(
                "172.29.216.82",
                "AA-BB-CC-DD-EE-FF",
                "Host01",
            )

        self.assertEqual(result["ip"], "172.29.216.82")
        self.assertEqual(result["mac"], "AA-BB-CC-DD-EE-FF")
        self.assertEqual(result["hostname"], "host01")
        self.assertEqual(result["device_type"], "Windows PC")
        self.assertEqual(result["enrichment_status"], "PARTIAL")
        self.assertIn("summary", result["enrichment_details"])

    def test_collect_device_enrichment_reports_timeout(self):
        with patch(
            "services.device_enrichment_service.probe_remote_tcp_port",
            return_value=(True, "TRUE"),
        ), patch(
            "services.device_enrichment_service.query_remote_windows_enrichment",
            return_value={
                "access_status": "Timeout",
                "enrichment_status": "TIMEOUT",
                "enrichment_details": {
                    "wmi": {
                        "status": "TIMEOUT",
                        "reason": "Remote Windows enrichment timed out.",
                    }
                },
            },
        ):
            result = collect_device_enrichment(
                "172.29.216.82",
                "AA-BB-CC-DD-EE-FF",
                "Host01",
            )

        self.assertEqual(result["enrichment_status"], "TIMEOUT")
        self.assertEqual(result["access_status"], "Timeout")


if __name__ == "__main__":
    unittest.main()

import ipaddress
import unittest
from types import SimpleNamespace

from ui.main_window import (
    _merge_discovery_entry,
    classify_discovered_device,
    get_network_targets,
    is_virtual_adapter_name,
    normalize_host_name,
    normalize_ip_address,
    normalize_mac_address,
    parse_ipconfig_dns_cache,
    parse_nbtstat_hostname,
    parse_ping_hostname,
)


class NetworkDiscoveryTests(unittest.TestCase):

    def test_mac_normalization(self):
        self.assertEqual(
            normalize_mac_address("aa-bb-cc-dd-ee-ff"),
            "AA-BB-CC-DD-EE-FF",
        )
        self.assertEqual(
            normalize_mac_address("AA:BB:CC:DD:EE:FF"),
            "AA-BB-CC-DD-EE-FF",
        )
        self.assertIsNone(normalize_mac_address("invalid"))

    def test_hostname_normalization(self):
        self.assertEqual(
            normalize_host_name("PC01.DOMAIN.LOCAL."),
            "pc01",
        )
        self.assertIsNone(normalize_host_name("   "))

    def test_ip_normalization(self):
        self.assertEqual(
            normalize_ip_address("172.29.219.11"),
            "172.29.219.11",
        )
        self.assertEqual(
            normalize_ip_address("not-an-ip"),
            "not-an-ip",
        )

    def test_network_targets_from_range(self):
        targets = get_network_targets("172.29.219.0/24")
        self.assertEqual(targets[0], "172.29.219.1")
        self.assertEqual(targets[-1], "172.29.219.254")

    def test_invalid_network_range_returns_empty_targets(self):
        self.assertEqual(get_network_targets("not-a-network"), [])

    def test_virtual_adapter_detection(self):
        self.assertTrue(is_virtual_adapter_name("vEthernet (Default Switch)"))
        self.assertFalse(is_virtual_adapter_name("Ethernet"))

    def test_device_matching_by_mac(self):
        computers = [
            SimpleNamespace(
                computer_name="Laptop-01",
                mac_address="AA-BB-CC-DD-EE-FF",
                ip_address="172.29.219.20",
            )
        ]

        result = classify_discovered_device(
            "172.29.219.20",
            "Laptop-01",
            "aa:bb:cc:dd:ee:ff",
            computers,
        )

        self.assertEqual(result["status"], "VERIFIED")
        self.assertEqual(result["matched_computer"].computer_name, "Laptop-01")

    def test_device_matching_by_ip(self):
        computers = [
            SimpleNamespace(
                computer_name="Laptop-02",
                mac_address="11-22-33-44-55-66",
                ip_address="172.29.219.40",
            )
        ]

        result = classify_discovered_device(
            "172.29.219.40",
            None,
            None,
            computers,
        )

        self.assertEqual(result["status"], "VERIFIED")
        self.assertEqual(result["matched_computer"].computer_name, "Laptop-02")

    def test_device_matching_by_hostname(self):
        computers = [
            SimpleNamespace(
                computer_name="Laptop-03",
                mac_address="66-55-44-33-22-11",
                ip_address="172.29.219.50",
            )
        ]

        result = classify_discovered_device(
            "172.29.219.99",
            "LAPTOP-03.domain.local.",
            None,
            computers,
        )

        self.assertEqual(result["status"], "MISMATCH")
        self.assertEqual(result["matched_computer"].computer_name, "Laptop-03")

    def test_new_device_when_no_inventory_record_exists(self):
        result = classify_discovered_device(
            "172.29.219.99",
            "Workstation-01",
            "00-11-22-33-44-55",
            [],
        )

        self.assertEqual(result["status"], "NEW_DEVICE")

    def test_merge_preserves_mac_and_hostname_without_duplicates(self):
        discovered_entries = {}

        _merge_discovery_entry(
            discovered_entries,
            "172.29.219.50",
            "Reception-PC",
            None,
            [],
        )

        _merge_discovery_entry(
            discovered_entries,
            "172.29.219.50",
            None,
            "AA-BB-CC-DD-EE-FF",
            [],
        )

        self.assertEqual(len(discovered_entries), 1)
        self.assertEqual(discovered_entries[next(iter(discovered_entries))]["ip_address"], "172.29.219.50")
        self.assertEqual(discovered_entries[next(iter(discovered_entries))]["mac_address"], "AA-BB-CC-DD-EE-FF")
        self.assertEqual(discovered_entries[next(iter(discovered_entries))]["hostname"], "reception-pc")

    def test_ip_address_network_calculation(self):
        network = ipaddress.IPv4Network(
            "172.29.219.11/255.255.255.0",
            strict=False,
        )
        self.assertEqual(str(network), "172.29.219.0/24")

    def test_ip_address_network_calculation_16(self):
        network = ipaddress.IPv4Network(
            "172.29.219.11/255.255.0.0",
            strict=False,
        )
        self.assertEqual(str(network), "172.29.0.0/16")

    def test_ip_address_network_calculation_8(self):
        network = ipaddress.IPv4Network(
            "172.29.219.11/255.0.0.0",
            strict=False,
        )
        self.assertEqual(str(network), "172.0.0.0/8")

    def test_parse_ping_hostname(self):
        output = "Pinging COMPUTER-NAME [192.168.100.1] with 32 bytes of data:"
        self.assertEqual(parse_ping_hostname(output), "computer-name")

    def test_parse_nbtstat_hostname(self):
        output = ""
        output += "Local Area Connection: "
        output += "Node IpAddress: [192.168.100.52] Scope Id: []"
        output += "\n"
        output += "\tHost not found"
        output += "\n"
        output += "\t" 
        output += "   COMPUTER-NAME        <00>  UNIQUE      Registered"
        self.assertEqual(parse_nbtstat_hostname(output), "computer-name")

    def test_parse_ipconfig_dns_cache(self):
        output = ""
        output += "Record Name . . . . . . . : COMPUTER-NAME.home.local\n"
        output += "Record Type . . . . . . . : 1\n"
        output += "Time To Live  . . . . . . : 120\n"
        output += "Data Length . . . . . . . : 4\n"
        output += "Section . . . . . . . . . : Answer\n"
        output += "A (Host) Record . . . . . : 192.168.100.1\n"
        self.assertEqual(parse_ipconfig_dns_cache(output, "192.168.100.1"), "computer-name")


if __name__ == "__main__":
    unittest.main()

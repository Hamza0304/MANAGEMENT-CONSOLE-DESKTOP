from services.device_enrichment_service import collect_device_enrichment
import pprint

TARGET_IP = "172.29.219.11"
TARGET_MAC = "50-84-92-04-E5-48"

print("=" * 70)
print("MANAGEMENT CONSOLE - LOCAL ENRICHMENT TEST")
print("=" * 70)

print(f"Target IP : {TARGET_IP}")
print(f"Target MAC: {TARGET_MAC}")
print()

try:
    result = collect_device_enrichment(
        ip_address=TARGET_IP,
        mac_address=TARGET_MAC,
    )

    print("=" * 70)
    print("RAW ENRICHMENT RESULT")
    print("=" * 70)

    pprint.pp(result, sort_dicts=False)

    print()
    print("=" * 70)
    print("IMPORTANT FIELDS")
    print("=" * 70)

    if isinstance(result, dict):
        for key, value in result.items():
            print(f"{key}: {value}")
    else:
        print("Result is not a dictionary.")
        print(type(result))

except Exception as exc:
    print("=" * 70)
    print("ENRICHMENT FAILED")
    print("=" * 70)
    print(type(exc).__name__)
    print(exc)

    import traceback
    traceback.print_exc()
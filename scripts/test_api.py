r"""Check that an API is reachable with the settings in .env.

    python scripts\test_api.py                       # GET API_BASE_URL
    python scripts\test_api.py --url /posts/1
"""
import argparse
import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from automation.api_client import ApiClient, ApiError  # noqa: E402
from automation.config import ConfigError  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--url", default="", help="full URL or path (default: API_BASE_URL itself)")
    args = ap.parse_args()
    try:
        client = ApiClient()
        target = args.url or client.settings.base_url or ""
        t0 = time.perf_counter()
        resp = client.request("GET", target)
    except (ConfigError, ApiError) as e:
        print(f"[FAIL] {e}")
        return 1
    ms = (time.perf_counter() - t0) * 1000
    print(f"[OK] HTTP {resp.status_code} in {ms:.0f} ms  ({resp.headers.get('Content-Type', '?')})")
    try:
        print(json.dumps(resp.json(), indent=2)[:600])
    except ValueError:
        print(resp.text[:300])
    return 0


if __name__ == "__main__":
    sys.exit(main())

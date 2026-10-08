"""Check connectivity to one or more configured databases.

    python scripts/test_db_connectivity.py             # test ALL configured DBs
    python scripts/test_db_connectivity.py corsair     # test one
    python scripts/test_db_connectivity.py corsair lara
    python scripts/test_db_connectivity.py --list      # show configured DBs (no connection)

Exit code is 0 only if every tested database logged in successfully.
"""
import argparse
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from automation.config import ConfigError, discover_databases, get_database, list_incomplete  # noqa: E402
from automation.connectivity import check_database  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawTextHelpFormatter)
    ap.add_argument("databases", nargs="*", help="database names (default: all configured)")
    ap.add_argument("--list", action="store_true", help="list configured databases and exit")
    ap.add_argument("--timeout", type=int, default=10, help="connect timeout in seconds")
    args = ap.parse_args()

    dbs = discover_databases()
    for name, missing in list_incomplete().items():
        print(f"[WARN] {name}: incomplete config, missing {', '.join(missing)}")

    if args.list:
        for c in dbs.values():
            print(f"{c.name:<12} env={c.env:<5} user={c.user:<15} dsn={c.dsn}")
        return 0 if dbs else 1

    try:
        targets = [get_database(n) for n in args.databases] if args.databases else list(dbs.values())
    except ConfigError as e:
        print(f"[ERROR] {e}")
        return 2
    if not targets:
        print("[ERROR] No databases configured. Copy .env.example to .env and fill it in.")
        return 2

    print(f"Testing {len(targets)} database(s)...\n")
    with ThreadPoolExecutor(max_workers=min(8, len(targets))) as pool:
        results = list(pool.map(lambda c: check_database(c, timeout=args.timeout), targets))

    for r in results:
        status = "OK  " if r.ok else "FAIL"
        print(f"[{status}] {r.name} ({r.env})  {r.dsn}")
        if r.tcp_ok is not None:
            print(f"        network : {'reachable' if r.tcp_ok else 'UNREACHABLE'} ({r.tcp_ms:.0f} ms)")
        if r.ok:
            print(f"        login   : success ({r.login_ms:.0f} ms)")
            print(f"        info    : {r.detail}")
        else:
            print(f"        error   : {r.error}")
            print(f"        hint    : {r.hint}")
        print()

    passed = sum(r.ok for r in results)
    print(f"Summary: {passed}/{len(results)} database(s) connected.")
    return 0 if passed == len(results) else 1


if __name__ == "__main__":
    sys.exit(main())

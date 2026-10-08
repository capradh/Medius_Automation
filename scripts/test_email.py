r"""Send a test email using the SMTP settings in .env.

    python scripts\test_email.py                      # to EMAIL_TO
    python scripts\test_email.py --to me@company.com
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from automation.config import ConfigError, get_smtp_settings, split_addresses  # noqa: E402
from automation.notifications import send_email  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--to", help="recipient(s), comma separated")
    args = ap.parse_args()
    try:
        s = get_smtp_settings()
        print(f"SMTP: host={s.host} port={s.port} security={s.security} user={s.user} from={s.sender}")
        send_email("Test email from Oracle Automation Toolkit",
                   "If you can read this, SMTP is configured correctly.",
                   to=split_addresses(args.to) or None)
    except ConfigError as e:
        print(f"[CONFIG ERROR] {e}")
        return 2
    except Exception as e:  # noqa: BLE001
        print(f"[FAIL] {type(e).__name__}: {e}")
        print("Hints: 535 = wrong user/password (Gmail needs an App Password; Office 365 needs SMTP AUTH enabled); "
              "timeouts = firewall/VPN or wrong port; 'certificate verify failed' = corporate TLS inspection.")
        return 1
    print("[OK] Email sent.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

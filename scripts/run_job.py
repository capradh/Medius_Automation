r"""Run an automation job.

    python scripts\run_job.py <job> [options]
    python scripts\run_job.py <job> --help       (shows that job's options)

Common options: --db NAME  --execute  --notify  --email-to a@x.com,b@x.com  --no-attach
"""
import argparse
import importlib
import logging
import pkgutil
import platform
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))  # works even without pip install -e .

import automation.jobs as jobs_pkg  # noqa: E402
from automation.config import ConfigError, get_database, get_settings, split_addresses  # noqa: E402
from automation.logging_setup import setup_logging  # noqa: E402


def available_jobs() -> list[str]:
    return sorted(m.name for m in pkgutil.iter_modules(jobs_pkg.__path__) if not m.name.startswith("_"))


def build_parser(module=None) -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="run_job.py", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Available jobs: " + ", ".join(available_jobs()),
    )
    p.add_argument("job", help="job name: " + ", ".join(available_jobs()))
    c = p.add_argument_group("common options")
    c.add_argument("--db", help="database name, e.g. corsair (falls back to DEFAULT_DB)")
    c.add_argument("--execute", action="store_true", help="commit changes (default is dry run / rollback)")
    c.add_argument("--notify", action="store_true", help="email the result (and failures) when finished")
    c.add_argument("--email-to", help="recipients (comma/semicolon separated); default EMAIL_TO from .env")
    c.add_argument("--no-attach", action="store_true", help="with --notify: send the summary without attachments")
    if module is not None and hasattr(module, "add_arguments"):
        module.add_arguments(p)
    return p


def notify(subject: str, body: str, args, files=()) -> None:
    from automation.notifications import send_email
    try:
        send_email(subject, body, to=split_addresses(args.email_to) or None,
                   attachments=[] if args.no_attach else list(files))
    except Exception as e:  # noqa: BLE001 - a mail problem must not hide the job result
        logging.getLogger("run_job").error("Could not send notification email: %s", e)


def main(argv=None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    first = argv[0] if argv and not argv[0].startswith("-") else None
    module = None
    if first:
        if first.startswith("_") or first not in available_jobs():
            print(f"Unknown job '{first}'. Available jobs: {', '.join(available_jobs())}")
            return 2
        module = importlib.import_module(f"automation.jobs.{first}")
    args = build_parser(module).parse_args(argv)

    setup_logging(args.job)
    log = logging.getLogger("run_job")

    cfg = None
    if getattr(module, "REQUIRES_DB", False) or args.db:
        name = args.db or get_settings().default_db
        if not name:
            print("This job needs a database. Use --db <name> or set DEFAULT_DB in .env")
            return 2
        try:
            cfg = get_database(name)
        except ConfigError as e:
            log.error(e)
            return 2
        args.db = cfg.name

    if args.execute and cfg and cfg.is_prod and getattr(module, "WRITES", False):
        print(f"\n!! You are about to COMMIT changes on PRODUCTION database '{cfg.name}'.")
        if input(f"Type the database name ({cfg.name}) to continue: ").strip().upper() != cfg.name:
            print("Cancelled.")
            return 1

    context = f"job={args.job}" + (f" db={cfg.name} env={cfg.env}" if cfg else "")
    log.info("Starting %s mode=%s", context, "EXECUTE" if args.execute else "DRY-RUN/READ")
    started = datetime.now()
    header = f"Job: {args.job}\nDatabase: {cfg.name if cfg else '-'}\nMachine: {platform.node()}\nStarted: {started:%Y-%m-%d %H:%M:%S}\n"
    try:
        result = module.run(args)
    except Exception as e:  # noqa: BLE001
        log.exception("Job failed")
        if args.notify:
            notify(f"[FAILED] {args.job}", f"{header}\nERROR: {e}\n\nSee the log file in logs/ for details.", args)
        return 1

    summary = result.summary if result else "Finished."
    files = result.files if result else []
    print(f"\n{summary}")
    for f in files:
        print(f"  -> {f}")
    log.info("Finished %s in %.1fs", context, (datetime.now() - started).total_seconds())
    if args.notify:
        notify(f"[OK] {args.job}", f"{header}\n{summary}\n", args, files)
    return 0


if __name__ == "__main__":
    sys.exit(main())

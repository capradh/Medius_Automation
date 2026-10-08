import logging
from datetime import datetime

from .config import get_settings


def setup_logging(job_name: str = "automation") -> None:
    """Log to console and to logs/<job>_<date>.log (audit trail)."""
    s = get_settings()
    s.log_dir.mkdir(exist_ok=True)
    logfile = s.log_dir / f"{job_name}_{datetime.now():%Y%m%d}.log"
    logging.basicConfig(
        level=s.log_level,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        handlers=[logging.StreamHandler(), logging.FileHandler(logfile, encoding="utf-8")],
        force=True,
    )

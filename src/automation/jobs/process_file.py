"""Upload/process csv, xls, xlsx and json files -> cleaned, calculated, formatted Excel report.

    python scripts/run_job.py process_file                       # everything in the input/ folder
    python scripts/run_job.py process_file --pick                # file-picker window
    python scripts/run_job.py process_file -i data\\sales.csv --group-by region --calc "total=price*qty"
"""
from __future__ import annotations

import logging
from datetime import datetime

import pandas as pd

from ..config import get_settings
from ..io_files import collect_input_files, pick_files, read_table
from ..results import JobResult
from ._common import add_processing_arguments, process_to_excel

log = logging.getLogger(__name__)


def add_arguments(p) -> None:
    g = p.add_argument_group("input options")
    g.add_argument("-i", "--input", nargs="+", metavar="PATH",
                   help="files, folders or wildcards (default: the input/ folder)")
    g.add_argument("--pick", action="store_true", help="choose files in a file-picker window")
    g.add_argument("--sheet", default="0", help="Excel sheet name or number (default: first sheet)")
    g.add_argument("--record-path", help="JSON only: dotted path to the list of records, e.g. data.items")
    g.add_argument("--combine", action="store_true", help="stack all files into ONE report (adds source_file column)")
    add_processing_arguments(p)


def run(args) -> JobResult:
    if args.pick:
        files = pick_files()
    elif args.input:
        files = collect_input_files(args.input)
    else:
        files = collect_input_files([get_settings().input_dir])
    files = [f for f in files if f.name != ".gitkeep"]
    if not files:
        return JobResult("No input files found. Put files in the input/ folder or use --input / --pick.")

    outputs, notes, failed = [], [], 0
    frames = []
    for f in files:
        try:
            df = read_table(f, sheet=args.sheet, record_path=args.record_path)
            if args.combine:
                frames.append(df.assign(source_file=f.name))
                notes.append(f"read {f.name}: {len(df)} rows")
            else:
                out = process_to_excel(df, args, f"{f.stem}_processed")
                outputs.append(out)
                notes.append(f"{f.name}: {len(df)} rows -> {out.name}")
        except Exception as e:  # noqa: BLE001 - one bad file must not stop the batch
            log.error("Skipping %s: %s", f.name, e)
            notes.append(f"FAILED {f.name}: {e}")
            failed += 1

    if args.combine and frames:
        combined = pd.concat(frames, ignore_index=True)
        outputs.append(process_to_excel(combined, args, f"combined_{datetime.now():%Y%m%d_%H%M%S}"))
        notes.append(f"combined {len(frames)} file(s) -> {outputs[-1].name}")

    if failed == len(files):  # nothing succeeded -> report the job as failed (exit code 1, failure email)
        raise RuntimeError("No file could be processed:\n" + "\n".join(notes))

    summary = f"Processed {len(files) - failed}/{len(files)} file(s). Output folder: {get_settings().output_dir}\n" + "\n".join(notes)
    return JobResult(summary, outputs)

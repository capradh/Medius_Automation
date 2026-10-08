"""Call a REST API and turn the JSON response into a processed Excel report.

    python scripts/run_job.py fetch_api --url /posts --group-by user_id
    python scripts/run_job.py fetch_api --url https://api.example.com/v1/orders --param status=open --record-path data.items
    python scripts/run_job.py fetch_api --url /orders --paginate --page-param page --max-pages 20
"""
from __future__ import annotations

from ..api_client import ApiClient
from ..io_files import json_to_frame
from ..results import JobResult
from ._common import add_processing_arguments, process_to_excel


def parse_key_values(items: list[str] | None) -> dict[str, str]:
    out = {}
    for item in items or []:
        if "=" not in item:
            raise ValueError(f"Expected KEY=VALUE, got '{item}'")
        k, v = item.split("=", 1)
        out[k.strip()] = v
    return out


def add_arguments(p) -> None:
    g = p.add_argument_group("API options")
    g.add_argument("--url", required=True, help="full URL, or a path appended to API_BASE_URL")
    g.add_argument("--method", choices=["GET", "POST"], default="GET")
    g.add_argument("--param", action="append", default=[], metavar="KEY=VALUE", help="query parameter (repeatable)")
    g.add_argument("--body", help="POST only: JSON string to send as the request body")
    g.add_argument("--record-path", help="dotted path to the list of records in the response, e.g. data.items")
    g.add_argument("--paginate", action="store_true", help="loop over ?page=1,2,3... until an empty page")
    g.add_argument("--page-param", default="page")
    g.add_argument("--max-pages", type=int, default=50)
    add_processing_arguments(p)


def run(args) -> JobResult:
    import json as _json

    client = ApiClient()
    params = parse_key_values(args.param)
    if args.paginate:
        payload = client.get_all_pages(args.url, params, args.page_param, args.record_path, max_pages=args.max_pages)
        df = json_to_frame(payload)
    elif args.method == "POST":
        body = _json.loads(args.body) if args.body else None
        df = json_to_frame(client.post_json(args.url, body, params), args.record_path)
    else:
        df = json_to_frame(client.get_json(args.url, params), args.record_path)

    out = process_to_excel(df, args, "api_report")
    return JobResult(f"Fetched {len(df)} record(s) from {args.url} -> {out.name}", [out])

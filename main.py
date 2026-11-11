"""CLI entry point: collect a process's summary emails for a week and
produce an Excel report with per-day rows and a week-total row.

Usage:
    python main.py --process <process_key>
    python main.py --process <process_key> --date 2026-11-10
    python main.py --process <process_key> --start 2026-11-09 --end 2026-11-16
"""
import argparse
import json
import sys
from collections import defaultdict
from datetime import datetime, timedelta

sys.path.insert(0, "src")
from outlook_client import connect, get_folder, get_items_in_range  # noqa: E402
from matcher import is_process_email  # noqa: E402
from extractor import extract_fields  # noqa: E402
from report import build_report  # noqa: E402


def week_bounds(any_date: datetime):
    start = any_date - timedelta(days=any_date.weekday())  # Monday
    start = start.replace(hour=0, minute=0, second=0, microsecond=0)
    end = start + timedelta(days=7)
    return start, end


def load_config(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def main():
    parser = argparse.ArgumentParser(description="Weekly summary-email rollup")
    parser.add_argument("--process", required=True, help="Process key from config file")
    parser.add_argument("--config", default="config/processes.json")
    parser.add_argument("--date", help="Any date (YYYY-MM-DD) inside the target week")
    parser.add_argument("--start", help="Explicit range start (YYYY-MM-DD), overrides --date")
    parser.add_argument("--end", help="Explicit range end, exclusive (YYYY-MM-DD)")
    parser.add_argument("--output", help="Output .xlsx path")
    args = parser.parse_args()

    config = load_config(args.config)
    processes = config["processes"]
    if args.process not in processes:
        sys.exit(f"Unknown process '{args.process}'. Configured: {list(processes)}")
    proc_cfg = processes[args.process]

    if args.start and args.end:
        start = datetime.strptime(args.start, "%Y-%m-%d")
        end = datetime.strptime(args.end, "%Y-%m-%d")
    else:
        ref_date = datetime.strptime(args.date, "%Y-%m-%d") if args.date else datetime.now()
        start, end = week_bounds(ref_date)

    namespace = connect()
    folder = get_folder(namespace, proc_cfg.get("folder", "Inbox"))
    items = get_items_in_range(folder, start, end)

    field_names = [f["name"] for f in proc_cfg["fields"]]
    per_day = defaultdict(lambda: {"email_count": 0, **{n: 0 for n in field_names}})

    for item in items:
        if not is_process_email(item, proc_cfg):
            continue
        day = item.ReceivedTime.date().isoformat()
        values = extract_fields(item, proc_cfg["fields"])
        per_day[day]["email_count"] += 1
        for name in field_names:
            if values.get(name) is not None:
                per_day[day][name] += values[name]

    daily_rows = [
        {"date": day, **data} for day, data in sorted(per_day.items())
    ]

    if not daily_rows:
        print(f"No matching emails found for '{args.process}' between "
              f"{start.date()} and {(end - timedelta(days=1)).date()}.")
        return

    output_path = args.output or f"{args.process}_{start.date()}_to_{(end - timedelta(days=1)).date()}.xlsx"
    build_report(proc_cfg.get("display_name", args.process), daily_rows, field_names, output_path)
    print(f"Report written to {output_path}")


if __name__ == "__main__":
    main()

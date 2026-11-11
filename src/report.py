"""Build the weekly Excel rollup for a process."""
from openpyxl import Workbook
from openpyxl.styles import Font


def build_report(process_name, daily_rows, field_names, output_path):
    """daily_rows: list of dicts with keys 'date', 'email_count', and each field name."""
    wb = Workbook()
    ws = wb.active
    ws.title = process_name[:31]

    headers = ["Date", "Email Count"] + field_names
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    totals = {name: 0 for name in field_names}
    for row in daily_rows:
        ws.append(
            [row["date"], row["email_count"]]
            + [row.get(name, 0) or 0 for name in field_names]
        )
        for name in field_names:
            totals[name] += row.get(name) or 0

    total_row = ["Week Total", sum(r["email_count"] for r in daily_rows)] + [
        totals[name] for name in field_names
    ]
    ws.append(total_row)
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)

    for col in ws.columns:
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col) + 2
        ws.column_dimensions[col[0].column_letter].width = width

    wb.save(output_path)

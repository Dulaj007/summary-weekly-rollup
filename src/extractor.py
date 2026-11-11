"""Pull numeric field values out of a summary email body.

Two field modes, chosen via each field's "mode" in config:

- "label" (default): find a cell/line containing the label text and read
  the number from that cell or the next cell in the row. Good for
  "Total: 42" style summaries.

- "table_column": for tables with a header row (e.g. Clinic, Total,
  Success, Bus Exception, Sys Exception) and one data row per item.
  Finds the column by its header text, then either sums every data row
  in that column ("sum_rows") or reads a single grand-total row
  ("total_row", matched by a keyword like "total" in the first cell).

Falls back to a plain-text regex search when no HTML table match is found
(label mode only).
"""
import re
from bs4 import BeautifulSoup

NUMBER_RE = re.compile(r"-?\d[\d,]*\.?\d*")


def _first_number(text):
    m = NUMBER_RE.search(text)
    if not m:
        return None
    return float(m.group().replace(",", ""))


def _tables_from_html(html_body: str):
    """Return a list of tables, each a list of rows, each a list of cell texts."""
    tables = []
    soup = BeautifulSoup(html_body or "", "html.parser")
    for table in soup.find_all("table"):
        rows = []
        for tr in table.find_all("tr"):
            cells = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
            if cells:
                rows.append(cells)
        if rows:
            tables.append(rows)
    return tables


def _matches_from_table(tables, label):
    label_lower = label.lower()
    matches = []
    for rows in tables:
        for cells in rows:
            for i, cell in enumerate(cells):
                if label_lower in cell.lower():
                    num_in_cell = _first_number(cell)
                    if num_in_cell is not None and cell.lower().strip() != label_lower:
                        matches.append(num_in_cell)
                        continue
                    for next_cell in cells[i + 1:]:
                        num = _first_number(next_cell)
                        if num is not None:
                            matches.append(num)
                            break
    return matches


def _matches_from_text(plain_body, label):
    pattern = re.compile(re.escape(label) + r"[^0-9\-]{0,20}(-?\d[\d,]*\.?\d*)", re.IGNORECASE)
    return [float(m.group(1).replace(",", "")) for m in pattern.finditer(plain_body or "")]


def _column_index(header_row, column_header):
    target = column_header.lower().strip()
    for idx, cell in enumerate(header_row):
        cell_norm = cell.lower().strip()
        if cell_norm == target or target in cell_norm:
            return idx
    return None


def _extract_table_column(tables, column_header, aggregate, total_row_keyword):
    for rows in tables:
        if len(rows) < 2:
            continue
        col_idx = _column_index(rows[0], column_header)
        if col_idx is None:
            continue

        data_rows = rows[1:]
        total_rows = [r for r in data_rows if r and total_row_keyword.lower() in r[0].lower()]
        plain_rows = [r for r in data_rows if r not in total_rows]

        if aggregate == "total_row":
            if total_rows:
                row = total_rows[-1]
                if col_idx < len(row):
                    num = _first_number(row[col_idx])
                    if num is not None:
                        return num
            continue  # no total row in this table; try the next one

        # sum_rows: sum the data rows, skipping any grand-total row to avoid double counting
        values = [_first_number(r[col_idx]) for r in plain_rows if col_idx < len(r)]
        values = [v for v in values if v is not None]
        if values:
            return sum(values)
    return None


def extract_fields(mail_item, fields_config):
    """Return {field_name: value_or_None} for a single email."""
    html_body = getattr(mail_item, "HTMLBody", "") or ""
    plain_body = getattr(mail_item, "Body", "") or ""
    tables = _tables_from_html(html_body)

    result = {}
    for field in fields_config:
        name = field["name"]
        mode = field.get("mode", "label")

        if mode == "table_column":
            result[name] = _extract_table_column(
                tables,
                field["column_header"],
                field.get("aggregate", "sum_rows"),
                field.get("total_row_keyword", "total"),
            )
            continue

        label = field["label"]
        occurrence = field.get("occurrence", "first")

        matches = _matches_from_table(tables, label)
        if not matches:
            matches = _matches_from_text(plain_body, label)

        if not matches:
            result[name] = None
        elif occurrence == "last":
            result[name] = matches[-1]
        elif occurrence == "sum":
            result[name] = sum(matches)
        else:
            result[name] = matches[0]
    return result

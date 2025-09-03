import os
import json
import csv
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.cell.write_only import WriteOnlyCell

# ----- Simple script detection for multilingual support -----
def detect_script_simple(text: str) -> str:
    if not text:
        return "DEFAULT"
    for c in text:
        if "\u4e00" <= c <= "\u9fff":  # Chinese
            return "CJK"
        elif "\u0900" <= c <= "\u097F":  # Devanagari
            return "DEVANAGARI"
        elif "\u0370" <= c <= "\u03FF":  # Greek
            return "GREEK"
        elif c.isascii():
            return "LATIN"
    return "DEFAULT"

# ----- Font mapping for Excel XLSX -----
FONT_MAP = {
    "DEFAULT": "Arial",
    "LATIN": "Arial",
    "DEVANAGARI": "Nirmala UI",
    "CJK": "Noto Sans CJK",
    "GREEK": "Arial"
}

# ----- Hybrid PDF → CSV/XLSX conversion (version 2.0) -----
def pdf_to_csv_v2(pdf_path, csv_path=None, xlsx_path=None, csv_delimiter=",",
                  excel_font_map=FONT_MAP, batch_log_every=1000, image_dir=None):
    """
    Convert PDF pages with table/CSV/JSON/XLS-like content into:
    - CSV file (utf-8-sig)
    - XLSX file (streaming write)
    - Optional image folder (if image_dir provided)
    """
    if not csv_path and not xlsx_path:
        raise ValueError("Provide at least one of csv_path or xlsx_path")

    # CSV writer
    csv_file = None
    csv_writer = None
    if csv_path:
        csv_file = open(csv_path, "w", encoding="utf-8-sig", newline="")
        csv_writer = csv.writer(csv_file, delimiter=csv_delimiter)

    # XLSX writer
    wb = None
    ws = None
    if xlsx_path:
        wb = Workbook(write_only=True)
        ws = wb.create_sheet(title="Extracted")

    if image_dir:
        os.makedirs(image_dir, exist_ok=True)

    total_rows = 0
    unsupported_detected = False

    try:
        with pdfplumber.open(pdf_path) as pdf:
            num_pages = len(pdf.pages)
            print(f"Opened PDF: {pdf_path} (pages: {num_pages})")

            for pageno, page in enumerate(pdf.pages, start=1):
                rows = []

                # --- Try table extraction ---
                try:
                    tables = page.extract_tables()
                    if tables:
                        for table in tables:
                            for row in table:
                                rows.append([cell for cell in row])
                except Exception:
                    pass

                # --- If no tables, check for JSON or embedded CSV-like content ---
                if not rows:
                    txt = page.extract_text() or ""
                    if txt.strip().startswith("{") and txt.strip().endswith("}"):
                        # Attempt parse JSON
                        try:
                            data = json.loads(txt)
                            # Flatten JSON into rows
                            def flatten(d, parent_key=""):
                                items = []
                                if isinstance(d, dict):
                                    for k, v in d.items():
                                        items.extend(flatten(v, f"{parent_key}{k}."))
                                elif isinstance(d, list):
                                    for i, v in enumerate(d):
                                        items.extend(flatten(v, f"{parent_key}{i}."))
                                else:
                                    items.append([parent_key.rstrip("."), str(d)])
                                return items
                            rows = flatten(data)
                        except Exception:
                            unsupported_detected = True
                    elif "," in txt or "\t" in txt:
                        # CSV-like text
                        delim = "\t" if "\t" in txt else ","
                        for line in txt.splitlines():
                            if line.strip():
                                rows.append([s.strip() for s in line.split(delim)])
                    else:
                        unsupported_detected = True

                if not rows:
                    continue  # skip page

                # --- Write rows ---
                for row in rows:
                    out_row = [("" if c is None else str(c)) for c in row]

                    # CSV
                    if csv_writer:
                        try:
                            csv_writer.writerow(out_row)
                        except Exception:
                            safe_row = [s.encode("utf-8", errors="ignore").decode("utf-8") for s in out_row]
                            csv_writer.writerow(safe_row)

                    # XLSX
                    if ws:
                        cells = []
                        for cell_val in out_row:
                            script = detect_script_simple(cell_val)
                            font_name = excel_font_map.get(script, excel_font_map.get("DEFAULT", "Arial"))
                            cell = WriteOnlyCell(ws, value=cell_val)
                            cell.font = Font(name=font_name, size=11)
                            cells.append(cell)
                        ws.append(cells)

                    total_rows += 1
                    if total_rows % batch_log_every == 0:
                        print(f"Processed rows: {total_rows} (page {pageno}/{num_pages})")

                # --- Optional images ---
                if image_dir:
                    for img_index, img in enumerate(page.images, start=1):
                        x0, top, x1, bottom = img["x0"], img["top"], img["x1"], img["bottom"]
                        cropped = page.crop((x0, top, x1, bottom))
                        pil_img = cropped.to_image(resolution=150).original
                        img_path = os.path.join(image_dir, f"page{pageno}_img{img_index}.png")
                        pil_img.save(img_path)

        if csv_file:
            csv_file.close()
        if wb:
            wb.save(xlsx_path)

        if unsupported_detected:
            print("⚠️ Some pages contained unsupported content (non-table/JSON/plain text). They were skipped.")

        print(f"✅ Done. Total rows written: {total_rows}")
        out = {"rows": total_rows}
        if csv_path:
            out["csv"] = os.path.abspath(csv_path)
        if xlsx_path:
            out["xlsx"] = os.path.abspath(xlsx_path)
        if image_dir:
            out["images"] = os.path.abspath(image_dir)
        return out

    except Exception as exc:
        if csv_file and not csv_file.closed:
            csv_file.close()
        if wb:
            wb.save(xlsx_path)
        raise RuntimeError(f"Error converting PDF: {exc}") from exc

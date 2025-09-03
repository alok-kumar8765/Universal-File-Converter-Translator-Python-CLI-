# Upload a Unicode Devanagari font, e.g., NotoSansDevanagari-Regular.ttf

# txt_converters.py
# --- add at top of file (or keep existing imports) ---
import os
import html
import shutil
import os, csv, json, io, textwrap
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import simpleSplit
from reportlab.lib.pagesizes import A4, letter
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.utils import simpleSplit
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from docx import Document
import pandas as pd
from openpyxl import Workbook
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from langdetect import detect
import regex  # better than re for Unicode script detection

# For Colab file handling
try:
    from google.colab import files
except ImportError:
    files = None


PAGE_W, PAGE_H = A4

def detect_script(text):
    """Detect script of a line and return font name."""
    # Register multiple fonts
    pdfmetrics.registerFont(TTFont("NotoSans", "/content/NotoSans-Regular.ttf")) #Add Local path of font
    pdfmetrics.registerFont(TTFont("NotoSansDevanagari", "/content/NotoSansDevanagari-Medium.ttf")) #Add Local path of font

    if regex.search(r'\p{Devanagari}', text):
        return "NotoSansDevanagari"
    return "NotoSans"  # default English font

def txt_to_pdf(txt_path, pdf_path, font_size=12, margin=40, line_gap=16):
    # Register multiple fonts
    pdfmetrics.registerFont(TTFont("NotoSans", "/content/NotoSans-Regular.ttf")) #Add Local path of font
    pdfmetrics.registerFont(TTFont("NotoSansDevanagari", "/content/NotoSansDevanagari-Medium.ttf")) #Add Local path of font

    try:
        c = canvas.Canvas(pdf_path, pagesize=A4)

        y = PAGE_H - margin
        max_width = PAGE_W - 2 * margin

        with open(txt_path, "r", encoding="utf-8", errors="ignore") as f:
            for raw_line in f:
                line = raw_line.strip("\n")

                # choose font dynamically
                font_name = detect_script(line)
                c.setFont(font_name, font_size)

                wrapped_lines = simpleSplit(line, font_name, font_size, max_width)
                if not wrapped_lines:
                    wrapped_lines = [" "]

                for wline in wrapped_lines:
                    if y < margin:  # new page
                        c.showPage()
                        y = PAGE_H - margin
                        c.setFont(font_name, font_size)

                    c.drawString(margin, y, wline)
                    y -= line_gap

        c.save()
        print(f"✅ PDF saved: {pdf_path}")
        return pdf_path
    except Exception as e:
        print("❌ txt_to_pdf failed:", str(e))
        return None


def txt_to_doc(input_file, output_file):
    with open(input_file, "r", encoding="utf-8", errors="ignore") as f:
        text = f.read()

    # ✅ Detect language
    try:
        language = detect(text)
    except:
        language = "en"

    # ✅ Font mapping by language
    font_map = {
        #"hi": "Mangal",        # Hindi
        "hi": "Nirmala UI",     # Hindi
        "bn": "Vrinda",           # Bengali
        "ta": "Latha",            # Tamil
        "te": "Gautami",          # Telugu
        "kn": "Tunga",            # Kannada
        "ml": "Kartika",          # Malayalam
        "gu": "Shruti",           # Gujarati
        "pa": "Raavi",            # Punjabi
        "or": "Kalinga",          # Odia

        "zh-cn": "Microsoft YaHei",   # Simplified Chinese
        "zh-tw": "PMingLiU",          # Traditional Chinese
        "ja": "MS Mincho",            # Japanese
        "ko": "Malgun Gothic",        # Korean

        "ar": "Traditional Arabic",   # Arabic
        "fa": "B Nazanin",            # Persian/Farsi
        "ur": "Jameel Noori Nastaleeq", # Urdu
        "he": "David",                # Hebrew

        "en": "Arial",         # English
        "ru": "Times New Roman",      # Russian
        "uk": "Times New Roman",      # Ukrainian
        "el": "Times New Roman",      # Greek
        "th": "Angsana New",          # Thai
        "vi": "Times New Roman",      # Vietnamese

        "fr": "Calibri",              # French
        "de": "Calibri",              # German
        "es": "Calibri",              # Spanish
        "it": "Calibri",              # Italian
        "pt": "Calibri",              # Portuguese
    }
    chosen_font = font_map.get(language, "Noto Sans") # fallback Noto Sans

    doc = Document()
    style = doc.styles['Normal']
    style.font.name = chosen_font
    style.font.size = Pt(12)

    for line in text.splitlines():
        if line.strip():  # skip blank lines
                p = doc.add_paragraph(line.strip())
                r = p.runs[0]
                r.font.name = chosen_font
                r._element.rPr.rFonts.set(qn('w:eastAsia'), chosen_font)

    doc.save(output_file)
    print(f"✅ Saved {output_file} with font '{chosen_font}' (detected lang: {language})")

    # ✅ Suggestion for user
    if language not in font_map:
        print(f"⚠️ Language '{language}' not mapped. "
              f"Please try 'Noto Sans' or manually set an appropriate font in MS Word.")
    else:
        print(f"💡 Tip: If the text doesn’t render well, "
              f"manually set the font in MS Word to '{chosen_font}' for proper display.")

#---------------- Font Detection ------------------
def get_font_for_line(line, font_size):
    # font mapping by language
    FONT_MAP = {
        "hi": "/content/NotoSansDevanagari-Medium.ttf",  # Hindi
        "en": "/content/NotoSans-Regular.ttf",           # English
        "zh-cn": "/content/NotoSansSC-Regular.ttf",      # Simplified Chinese
        "ja": "/content/NotoSansJP-Regular.ttf",         # Japanese
        "ko": "/content/NotoSansKR-Regular.ttf",         # Korean
        "default": "/content/NotoSans-Regular.ttf"       # fallback
    }

    try:
        lang = detect(line) if line.strip() else "en"
    except:
        lang = "en"

    # normalize
    if lang.startswith("zh"):
        lang = "zh-cn"
    elif lang.startswith("ja"):
        lang = "ja"
    elif lang.startswith("ko"):
        lang = "ko"
    elif lang.startswith("hi"):
        lang = "hi"
    else:
        lang = "en"

    font_path = FONT_MAP.get(lang, FONT_MAP["default"])
    if not os.path.exists(font_path):  # fallback if missing
        font_path = FONT_MAP["default"]

    return ImageFont.truetype(font_path, font_size)


#---------------- TXT to Image ------------------
def txt_to_image(txt_path, output_dir, font_size=24, width=1240, height=1754,
                 margin=40, max_lines_per_img=70, split=None):

    try:
        # Read text file
        with open(txt_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()

        os.makedirs(output_dir, exist_ok=True)
        images = []
        line_height = font_size + 10

        # Auto decide if split is None
        if split is None:
            if len(lines) <= 20:
                split = False   # single image
            else:
                # user choice
                choice = input("Text has more than 20 lines. Do you want multiple images? (y/n): ")
                split = True if choice.lower().startswith("y") else False

        if not split:
            # Option 1: All lines in ONE big image
            total_height = margin*2 + len(lines)*line_height
            img = Image.new("RGB", (width, total_height), "white")
            draw = ImageDraw.Draw(img)

            y = margin
            for line in lines:
                font = get_font_for_line(line, font_size)
                draw.text((margin, y), line.strip(), fill="black", font=font)
                y += line_height

            out_path = os.path.join(output_dir, "output_single.png")
            img.save(out_path)
            images.append(out_path)

        else:
            # Option 2: Split into multiple images
            chunks = [lines[i:i+max_lines_per_img] for i in range(0, len(lines), max_lines_per_img)]

            for idx, chunk in enumerate(chunks, start=1):
                img = Image.new("RGB", (width, height), "white")
                draw = ImageDraw.Draw(img)

                y = margin
                for line in chunk:
                    font = get_font_for_line(line, font_size)
                    draw.text((margin, y), line.strip(), fill="black", font=font)
                    y += line_height

                out_path = os.path.join(output_dir, f"page_{idx}.png")
                img.save(out_path)
                images.append(out_path)

        print(f"✅ Images saved: {images}")
        return images

    except Exception as e:
        print("❌ txt_to_image failed:", str(e))
        return None

import os, csv
import sys

# Optional libs
try:
    from langdetect import detect, detect_langs
    _HAS_LANGDETECT = True
except Exception:
    _HAS_LANGDETECT = False

# pandas only if user wants xlsx export or fallback viewing
try:
    import pandas as pd
    _HAS_PANDAS = True
except Exception:
    _HAS_PANDAS = False

# For local GUI popup (falls back to print)
try:
    import tkinter as tk
    from tkinter import messagebox
    _HAS_TK = True
except Exception:
    _HAS_TK = False

def _auto_detect_delimiter(txt_path, sample_bytes=8192):
    """Try to detect delimiter from first KBs of file; fallback to tab."""
    try:
        with open(txt_path, "r", encoding="utf-8", errors="ignore") as f:
            sample = f.read(sample_bytes)
        sniffer = csv.Sniffer()
        dialect = sniffer.sniff(sample)
        return dialect.delimiter
    except Exception:
        # Common fallbacks
        for d in [",", "\t", ";", "|"]:
            if d in sample:
                return d
    return "\t"

def _detect_non_english_and_scripts(txt_path, max_chars=10000):
    """
    Returns tuple (is_non_english_bool, top_languages_list, has_non_ascii_bool, scripts_set)
    Uses langdetect if available; always returns heuristic based on non-ASCII presence.
    """
    text_sample = ""
    with open(txt_path, "r", encoding="utf-8", errors="ignore") as f:
        # read some lines to keep it fast for very large files
        for i, line in enumerate(f):
            text_sample += line
            if len(text_sample) >= max_chars:
                break

    has_non_ascii = any(ord(ch) > 127 for ch in text_sample)
    top_langs = []
    non_english = False

    if _HAS_LANGDETECT:
        try:
            langs = detect_langs(text_sample)
            top_langs = [(str(l).split(":")[0], float(str(l).split(":")[1])) for l in langs]
            # consider non-english if top language not 'en' with decent prob
            if top_langs:
                top_lang = top_langs[0][0]
                top_prob = top_langs[0][1]
                non_english = not (top_lang == "en" and top_prob >= 0.65)
        except Exception:
            top_langs = []
            non_english = has_non_ascii
    else:
        non_english = has_non_ascii

    # simple script heuristics (return set like {'Devanagari','CJK','Arabic',...})
    scripts = set()
    for ch in text_sample[:2000]:
        cp = ord(ch)
        if 0x0900 <= cp <= 0x097F:
            scripts.add("Devanagari")
        elif 0x4E00 <= cp <= 0x9FFF:
            scripts.add("CJK")
        elif 0x3040 <= cp <= 0x30FF:
            scripts.add("Japanese")
        elif 0xAC00 <= cp <= 0xD7AF:
            scripts.add("Hangul")
        elif 0x0600 <= cp <= 0x06FF:
            scripts.add("Arabic")
        elif cp > 127 and not ch.isspace():
            scripts.add("OtherNonASCII")

    return non_english, top_langs, has_non_ascii, scripts

def _notify_user_excel_instructions(csv_path, non_english, top_langs, scripts):
    """
    Shows instructions to user on how to open CSV in Excel correctly.
    If running in Colab -> just prints instructions. If local & tkinter available -> show popup.
    """
    languages_text = ""
    if top_langs:
        languages_text = ", ".join([f"{l}:{p:.2f}" for l, p in top_langs])
    else:
        languages_text = "detected non-ASCII text" if non_english else "English/ASCII"

    instructions = (
        f"CSV saved at: {csv_path}\n\n"
        f"Detected languages/sample: {languages_text}\n\n"
        "For best results open this CSV in Excel like this:\n"
        "1) Open Excel -> Data tab -> Get Data -> From Text/CSV\n"
        "2) Select the file and when the preview appears, choose 'File origin' = UTF-8 (or '65001: Unicode (UTF-8)')\n"
        "3) Click 'Load'.\n\n"
        "Alternative (no import wizard): this script saved the file with a BOM (utf-8-sig) so modern Excel often opens it fine by double-clicking.\n\n"
        "If users still see garbled characters, consider opening in LibreOffice or importing via the Data menu and explicitly choosing UTF-8.\n\n"
        "Tip: To avoid Excel problems at all, open the provided .xlsx version (if available) which preserves texts in all languages."
    )

    # If Colab environment, just print
    try:
        import google.colab  # type: ignore
        in_colab = True
    except Exception:
        in_colab = False

    if in_colab or not _HAS_TK:
        print("\n" + "="*60 + "\nIMPORTANT: How to open this CSV in Excel\n" + "="*60)
        print(instructions)
        return

    # else show a tkinter popup (local desktop)
    try:
        root = tk.Tk()
        root.withdraw()
        # Use a scrolled text popup? Simpler: showinfo with basic instructions
        messagebox.showinfo("CSV saved — How to open in Excel (UTF-8)", instructions)
        root.destroy()
    except Exception:
        print(instructions)


def txt_to_csv(
    txt_path,
    csv_path,
    delimiter="\t",
    auto_detect_delimiter=False,
    excel_friendly=True,   # writes utf-8-sig so Excel auto-detects BOM
    notify_user=True,
    save_xlsx=False        # if True and pandas available, also write .xlsx beside csv
):
    """
    Stream text -> CSV preserving UTF-8 and user friendly Excel behavior.

    Parameters:
      - txt_path: input text file path (UTF-8 expected)
      - csv_path: output path (.csv)
      - delimiter: delimiter to split lines into columns (default tab)
      - auto_detect_delimiter: try to detect delimiter automatically from file sample
      - excel_friendly: if True write using encoding='utf-8-sig' (BOM) -> Excel on Windows reads well
      - notify_user: print or popup instructions on how to open in Excel if non-English content detected
      - save_xlsx: if True and pandas installed, also write an .xlsx copy (recommended)
    """

    if not os.path.exists(txt_path):
        raise FileNotFoundError(f"Input file not found: {txt_path}")

    # auto-detect delimiter if requested
    if auto_detect_delimiter:
        try:
            delimiter = _auto_detect_delimiter(txt_path)
        except Exception:
            pass

    # detect non-english content & scripts
    non_english, top_langs, has_non_ascii, scripts = _detect_non_english_and_scripts(txt_path)

    # choose encoding
    enc = "utf-8-sig" if excel_friendly else "utf-8"

    # Write CSV streaming line-by-line
    with open(txt_path, "r", encoding="utf-8", errors="ignore") as src, \
         open(csv_path, "w", encoding=enc, newline="") as dst:
        writer = csv.writer(dst)
        for raw in src:
            row = raw.rstrip("\n").split(delimiter)
            # ensure row is list of str (no bytes)
            row = [("" if v is None else str(v)) for v in row]
            writer.writerow(row)

    # Optionally save XLSX to avoid Excel encoding issues completely
    xlsx_path = None
    if save_xlsx:
        if not _HAS_PANDAS:
            print("⚠️ pandas not installed; cannot save .xlsx. Install pandas + openpyxl to enable this.")
        else:
            try:
                # read via pandas (fast for reasonably sized files)
                df = pd.read_csv(csv_path, encoding=enc)
                xlsx_path = os.path.splitext(csv_path)[0] + ".xlsx"
                df.to_excel(xlsx_path, index=False, engine="openpyxl")
            except Exception as e:
                print("⚠️ Could not create .xlsx:", e)

    # Notify user (print or popup) with instructions if non-english content or user asked to be notified
    if notify_user:
        _notify_user_excel_instructions(csv_path, non_english, top_langs, scripts)

    return {"csv": csv_path, "xlsx": xlsx_path, "non_english": non_english, "langs": top_langs, "scripts": scripts}

def txt_to_xls(txt_path, xls_path, delimiter="\t"):
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Sheet1")
    with open(txt_path, "r", encoding="utf-8", errors="ignore") as src:
        for line in src:
            ws.append(line.rstrip("\n").split(delimiter))
    wb.save(xls_path)
    return xls_path

def txt_to_json(txt_path, json_path):
    # Converts each non-empty line to list; if delimiter-like structure exists, user can post-process.
    try:
        data = []
        with open(txt_path, "r", encoding="utf-8", errors="ignore") as f:
            for line in f:
                s = line.strip()
                if s:
                    data.append(s)
        with open(json_path, "w", encoding="utf-8") as out:
            json.dump(data, out, ensure_ascii=False, indent=2)
        return json_path

    except Exception as e:
        print(f"❌ TXT to JSON failed: {e}")
        return None

# ---------------- CLI MENU ---------------- #

def main():
    while True:
        print("\n==== TXT File Converter ====")
        print("1. Convert TXT File")
        print("2. Exit")
        choice = input("Enter choice: ").strip()

        if choice == "1":
            # Upload file (Colab) or enter path (offline)
            if files:
                #txt_file = input("Enter path of TXT file: ").strip()
                uploaded = files.upload()
                txt_file = list(uploaded.keys())[0]
            else:
                txt_file = input("Enter path of TXT file: ").strip()

            print("\nSupported conversions: ")
            print("1. TXT → PDF")
            print("2. TXT → DOCX")
            print("3. TXT → PNG (image)")
            print("4. TXT → CSV")
            print("5. TXT → XLSX")
            print("6. TXT → JSON")
            fmt_choice = input("Select target format (1-6): ").strip()

            base, _ = os.path.splitext(txt_file)
            out_file = None

            try:
                if fmt_choice == "1":
                    out_file = base + ".pdf"
                    txt_to_pdf(txt_file, out_file)
                elif fmt_choice == "2":
                    out_file = base + ".docx"
                    txt_to_doc(txt_file, out_file)
                elif fmt_choice == "3":
                    out_file = base + ".png"
                    txt_to_image(txt_file, out_file)
                elif fmt_choice == "4":
                    out_file = base + ".csv"
                    txt_to_csv(txt_file, out_file)
                elif fmt_choice == "5":
                    out_file = base + ".xlsx"
                    txt_to_xls(txt_file, out_file)
                elif fmt_choice == "6":
                    out_file = base + ".json"
                    txt_to_json(txt_file, out_file)
                else:
                    print("❌ Invalid choice!")
                    continue

                print(f"✅ Converted successfully: {out_file}")

                # Colab download option
                if files:
                    files.download(out_file)

            except Exception as e:
                print("❌ Conversion failed:", e)

        elif choice == "2":
            print("👋 Exiting...")
            break
        else:
            print("❌ Invalid choice!")

if __name__ == "__main__":
    main()

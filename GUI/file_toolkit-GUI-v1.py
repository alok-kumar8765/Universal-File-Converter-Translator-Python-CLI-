import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
import pandas as pd
from docx import Document
from PIL import Image, ImageDraw
import pypandoc

# For offline translation (Helsinki-NLP model)
from transformers import MarianMTModel, MarianTokenizer
import torch

# ---------------- FILE FUNCTIONS ----------------
def pdf_to_text(file_path, out_path):
    text = ""
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ""
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(text)
    return out_path

def doc_to_text(file_path, out_path):
    doc = Document(file_path)
    text = "\n".join([para.text for para in doc.paragraphs])
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(text)
    return out_path

def text_to_image(file_path, out_path):
    with open(file_path, "r", encoding="utf-8") as f:
        text = f.read()
    img = Image.new("RGB", (1200, 1600), "white")
    draw = ImageDraw.Draw(img)
    draw.text((40, 40), text, fill="black")
    img.save(out_path)
    return out_path

def csv_to_xls(file_path, out_path):
    df = pd.read_csv(file_path)
    df.to_excel(out_path, index=False)
    return out_path

def generic_convert(in_file, out_file, format_to):
    pypandoc.convert_file(in_file, format_to, outputfile=out_file, extra_args=['--standalone'])
    return out_file

# ---------------- OFFLINE TRANSLATION ----------------
class OfflineTranslator:
    def __init__(self, src="en", tgt="hi"):
        model_name = f"Helsinki-NLP/opus-mt-{src}-{tgt}"
        self.tokenizer = MarianTokenizer.from_pretrained(model_name)
        self.model = MarianMTModel.from_pretrained(model_name)

    def translate(self, text):
        inputs = self.tokenizer(text, return_tensors="pt", padding=True, truncation=True)
        translated = self.model.generate(**inputs)
        return [self.tokenizer.decode(t, skip_special_tokens=True) for t in translated][0]

# ---------------- GUI ----------------
class FileToolkitApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Converter & Translator")
        self.root.geometry("600x400")

        # File selection
        self.file_label = tk.Label(root, text="No file selected")
        self.file_label.pack(pady=10)

        tk.Button(root, text="📂 Choose File", command=self.pick_file).pack()

        # Conversion format
        tk.Label(root, text="Convert To:").pack()
        self.format_combo = ttk.Combobox(root, values=["pdf", "txt", "docx", "csv", "xls", "png"])
        self.format_combo.pack(pady=5)

        tk.Button(root, text="Convert", command=self.convert_file).pack(pady=10)

        # Translation
        tk.Label(root, text="Translate (offline, Helsinki-NLP)").pack()
        self.src_lang = tk.Entry(root)
        self.src_lang.insert(0, "en")
        self.src_lang.pack(pady=2)
        self.dest_lang = tk.Entry(root)
        self.dest_lang.insert(0, "hi")
        self.dest_lang.pack(pady=2)

        tk.Button(root, text="Translate File", command=self.translate_file).pack(pady=10)

        self.file_path = None

    def pick_file(self):
        self.file_path = filedialog.askopenfilename()
        if self.file_path:
            self.file_label.config(text=f"Selected: {os.path.basename(self.file_path)}")

    def convert_file(self):
        if not self.file_path:
            messagebox.showerror("Error", "No file selected")
            return
        fmt = self.format_combo.get().strip().lower()
        if not fmt:
            messagebox.showerror("Error", "Choose a target format")
            return
        out_file = os.path.splitext(self.file_path)[0] + "_converted." + fmt
        try:
            ext = os.path.splitext(self.file_path)[1].lower()
            if ext == ".pdf" and fmt == "txt":
                pdf_to_text(self.file_path, out_file)
            elif ext == ".docx" and fmt == "txt":
                doc_to_text(self.file_path, out_file)
            elif ext == ".txt" and fmt == "png":
                text_to_image(self.file_path, out_file)
            elif ext == ".csv" and fmt == "xls":
                csv_to_xls(self.file_path, out_file)
            else:
                generic_convert(self.file_path, out_file, fmt)
            messagebox.showinfo("Success", f"Converted file saved at:\n{out_file}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def translate_file(self):
        if not self.file_path:
            messagebox.showerror("Error", "No file selected")
            return
        src = self.src_lang.get().strip()
        tgt = self.dest_lang.get().strip()
        try:
            trans = OfflineTranslator(src, tgt)
            text = ""
            if self.file_path.endswith(".pdf"):
                text = open(pdf_to_text(self.file_path, "temp.txt"), encoding="utf-8").read()
            elif self.file_path.endswith(".docx"):
                text = open(doc_to_text(self.file_path, "temp.txt"), encoding="utf-8").read()
            else:
                text = open(self.file_path, encoding="utf-8").read()

            translated = trans.translate(text[:1000])  # Limit for demo
            out_file = os.path.splitext(self.file_path)[0] + f"_{tgt}.txt"
            with open(out_file, "w", encoding="utf-8") as f:
                f.write(translated)
            messagebox.showinfo("Success", f"Translated file saved at:\n{out_file}")
        except Exception as e:
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = FileToolkitApp(root)
    root.mainloop()

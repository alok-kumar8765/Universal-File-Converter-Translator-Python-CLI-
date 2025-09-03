# About This Folder

	 
	This Folder contains seprate converter.
	Supported Files : 
	 - Pdf-to-Doc
	 - Pdf-to-Txt
	 - Pdf-to-Img
	 - Pdf-to-Csv
	 - Pdf-to-Xls
	 
	 - Doc-to-Pdf
	 - Doc-to-Txt
	 - Doc-to-Csv
	 - Doc-to-Xls
	 - Doc-to-Img
	 
	 - Txt-to-Pdf
	 - Txt-to-Doc
	 - Txt-to-Csv
	 - Txt-to-Xls
	 - Txt-to-Img
	 
	 - Csv-to-Pdf
	 - Csv-to-Doc
	 - Csv-to-Txt 
	 - Csv-to-Xls
	 - Csv-to-Json
	 
	 - Xls-to-Pdf
	 - Xls-to-Doc
	 - Xls-to-Txt 
	 - Xls-to-Csv
	 - Xls-to-Json
	 
	 - Json-to-Cvs
	 - Json-to-Xls
	 - Json-to-Txt
	 
	 - Png-to-Jpg
	 - Png-to-Jpeg
	 - Jpg-to-Jpeg
	 - Jpg-to-Png
	 - Jpeg-to-Jpg
	 - Jpeg-to-Png
	 
	 we can handle file size upto 10+ MB, and Page Count
	 upto 4000.
	 
	 This is just a beta/Prototype.
	
	 
# 🔧 Install (Colab & Laptop)

> 1. Colab (run once in a new cell)

 
		!apt-get -y install poppler-utils tesseract-ocr
		!pip install pdfplumber PyPDF2 python-docx pandas openpyxl pillow reportlab pdf2image ijson pytesseract

	
> 2. Windows/Mac/Linux (local)

		pip install pdfplumber PyPDF2 python-docx pandas openpyxl pillow reportlab pdf2image ijson pytesseract

	
> 3. Also install:
		
		- **Poppler (for pdf2image)**
				
			- Windows: download poppler binaries and add bin to PATH (google "poppler for windows")
			- Mac: brew install poppler
			- Linux: sudo apt-get install poppler-utils
				
			
		- **Tesseract OCR (for image → txt)**
		
			- Windows: install from UB Mannheim build (add to PATH)
			- Mac: brew install tesseract
			- Linux: sudo apt-get install tesseract-ocr
				
			
		- **Set Tesseract path on Windows if needed:**
				
			
				import pytesseract
				pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
			
			
## 🧪 Quick Usage Examples of "universal_converter.py"
	
	
		- *** TXT ***
		
			- from txt_converters import txt_to_pdf, txt_to_doc, txt_to_image, txt_to_csv, txt_to_xls, txt_to_json
			- txt_to_pdf("big.txt", "big.pdf")
			- txt_to_doc("big.txt", "big.docx")
			- txt_to_image("big.txt", "big.png")
			- txt_to_csv("big.txt", "big.csv", delimiter="\t")  # or ","
			- txt_to_xls("big.txt", "big.xlsx", delimiter="\t")
			- txt_to_json("big.txt", "big.json")

		- ** PDF **
		
			- from pdf_converters import pdf_to_txt, pdf_to_doc, pdf_to_image, pdf_to_csv, pdf_to_xls, pdf_to_json
			- pdf_to_txt("input.pdf", "out.txt")
			- pdf_to_doc("input.pdf", "out.docx")
			- pdf_to_image("input.pdf", "out_images")  # folder
			- pdf_to_csv("input.pdf", "out.csv")
			- pdf_to_xls("input.pdf", "out.xlsx")
			- pdf_to_json("input.pdf", "out.json")

		- ** DOCX **
		
			- from doc_converters import doc_to_txt, doc_to_pdf, doc_to_csv, doc_to_xls, doc_to_image, doc_to_json
			- doc_to_txt("doc.docx", "doc.txt")
			- doc_to_pdf("doc.docx", "doc.pdf")
			- doc_to_csv("doc.docx", "doc.csv")
			- doc_to_xls("doc.docx", "doc.xlsx")
			- doc_to_image("doc.docx", "doc.png")
			- doc_to_json("doc.docx", "doc.json")

		- ** IMAGES & OCR**
		
			- from image_converters import jpg_to_png, png_to_jpg, jpg_to_jpeg, png_to_jpeg, jpeg_to_png, jpeg_to_jpg, gif_to_png, image_to_txt_ocr
			- jpg_to_png("a.jpg", "a.png")
			- image_to_txt_ocr("scan.png", "scan.txt", lang="eng")  # needs Tesseract

		- ** CSV **
		
			- from csv_converters import csv_to_xls, csv_to_pdf, csv_to_doc, csv_to_txt, csv_to_json
			- csv_to_xls("data.csv", "data.xlsx")
			- csv_to_pdf("data.csv", "data.pdf")
			- csv_to_doc("data.csv", "data.docx")
			- csv_to_txt("data.csv", "data.txt")
			- csv_to_json("data.csv", "data.json")

		- ** XLSX **
		
			- from xls_converters import xls_to_csv, xls_to_doc, xls_to_txt, xls_to_pdf, xls_to_json
			- xls_to_csv("sheet.xlsx", "sheet.csv")
			- xls_to_doc("sheet.xlsx", "sheet.docx")
			- xls_to_txt("sheet.xlsx", "sheet.txt")
			- xls_to_pdf("sheet.xlsx", "sheet.pdf")
			- xls_to_json("sheet.xlsx", "sheet.json")

		- ** JSON **
		
			- from json_converters import json_to_csv, json_to_xls, json_to_txt, txt_to_json, doc_to_json, pdf_to_json
			- json_to_csv("data.json", "data.csv")
			- json_to_xls("data.json", "data.xlsx")
			- json_to_txt("data.json", "data.txt")

	
		
# 📝 TXT Converter Toolkit  

A lightweight but powerful toolkit to convert **.txt files** into multiple formats — with full support for **any language** (Hindi, English, Japanese, Arabic, etc.), and **any size** (from a few lines to hundreds of pages).  

This project ensures that no matter what your text file looks like, you’ll get **proper, usable output** in your chosen format.  

---

## 🚀 Supported Conversions  

- **txt_to_doc** → Convert `.txt` files into `.docx` (Word documents).  
- **txt_to_pdf** → Generate clean `.pdf` files from text (multi-page supported).  
- **txt_to_csv** → Transform structured text into `.csv` (Excel/Sheets compatible).  
- **txt_to_image** → Render text content into images (for sharing/presentations).  
- **txt_to_json** → Parse and convert `.txt` into structured `.json` data.  
- **txt_to_xls** → Export `.txt` content directly into `.xls` (Excel) format.  

---

## 🌍 Key Features  

- ✅ **Language Agnostic**: Works with **any script/language** (e.g., Devanagari, Latin, Chinese, Arabic, etc.).  
- ✅ **Scalable**: Handles **any size** text file — from small notes to large multi-page documents.  
- ✅ **Consistent Output**: Ensures proper formatting in all supported conversions.  
- ✅ **Error Handling**: If a particular language/script creates an issue in the output, the tool will **notify clearly** so it can be fixed.  
- ✅ **Community-Friendly**: Developers and contributors are welcome to improve, add features, or share ideas.  

---

## 🛠️ How to Use  

Each function takes an input `.txt` file and outputs the converted format.  

Example (Python):  


	from converter import txt_to_csv, txt_to_pdf

	# Convert TXT to CSV
	txt_to_csv("input.txt", "output.csv")

	# Convert TXT to PDF
	txt_to_pdf("input.txt", "output.pdf")
	

# 🧑‍🤝‍🧑 Contributions


- Found an issue with a specific language output?

- Have an idea to optimize performance or add new conversions?
	
	
👉 Open an issue or create a pull request — all contributions are welcome!
Together, we can make this toolkit better and more universal.

# 📌 Roadmap

 
- Improve auto language detection & handling edge cases.

- Add support for more structured conversions (e.g., txt → XML, txt → Markdown).

- Expand test cases for multi-lingual + large file inputs.
	

# 📸 Demo (Screenshots / GIFs)

### Example: TXT → PDF Conversion
![TXT to PDF Demo](assets/demo_pdf.gif)

### Example: TXT → CSV Conversion
![TXT to CSV Demo](assets/demo_csv.png)



	
	
	
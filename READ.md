# 🛠️ File Toolkit – Universal File Converter & Translator (Python CLI)

# About Project

	 
	This Project is designed to convert files extension
	without affecting the content of the file.
	A single Python script to **convert and translate files** between multiple formats: 

	 - ✅ Pdf-to-Doc
	 - ✅ Pdf-to-Txt
	 - ✅ Pdf-to-Img
	 - ✅ Pdf-to-Csv
	 - ✅ Pdf-to-Xls
	 
	 - ✅ Doc-to-Pdf
	 - ✅ Doc-to-Txt
	 - ✅ Doc-to-Csv
	 - ✅ Doc-to-Xls
	 - ✅ Doc-to-Img
	 
	 - ✅ Txt-to-Pdf
	 - ✅ Txt-to-Doc
	 - ✅ Txt-to-Csv
	 - ✅ Txt-to-Xls
	 - ✅ Txt-to-Img
	 
	 - ✅ Csv-to-Pdf
	 - ✅ Csv-to-Doc
	 - ✅ Csv-to-Txt 
	 - ✅ Csv-to-Xls
	 - ✅ Csv-to-Json
	 
	 - ✅ Xls-to-Pdf
	 - ✅ Xls-to-Doc
	 - ✅ Xls-to-Txt 
	 - ✅ Xls-to-Csv
	 - ✅ Xls-to-Json
	 
	 - ✅ Json-to-Cvs
	 - ✅ Json-to-Xls
	 - ✅ Json-to-Txt
	 
	 - ✅ Png-to-Jpg
	 - ✅ Png-to-Jpeg
	 - ✅ Jpg-to-Jpeg
	 - ✅ Jpg-to-Png
	 - ✅ Jpeg-to-Jpg
	 - ✅ Jpeg-to-Png
	 - ✅ Image-to-Txt

	 we can handle file size upto 30 MB, and Page Count
	 upto 4000.
	 
	 This is just a beta/Prototype.
	 
	We Also Provide Laguage Translation.
		- you have to uploade your file, Select the language
		of your file & select the desired language you want,
		you will get the proper file in your desired language.
		
		- We Supported file size upto 40 MB.
		- Currently this translation feature worked when you are 
		connected to internet, but we also have offfline mode which 
		is little bulky size.
	 
	
# You Get In This Repo

 - Script Code for Translation & Convert File, which run in you local machine.
 - Executable file which is helpful for non coding bacground person, as they can simply install this software in their  machine and use it.
 - Code that run in google colab.

# Libraries We’ll Use

	- pypandoc → Convert between docx, pdf, txt, md, rtf.

	- pdfplumber / PyMuPDF → Extract text from PDF.

	- python-docx → Handle .docx.

	- pandas → Handle CSV/XLS.

	- PIL → Save text as image.

	- googletrans (or deep-translator) → Translate text between languages. 

# ⚙️ Installation



	# Clone repo
	git clone https://github.com/alok-kumar8765/Universal-File-Converter-Translator-Python-CLI-.git
	cd your-repo

	# Install dependencies
	pip install -r requirements.txt




# Requirement 

	| While installing the following requirements you might face some trouble, here is the solution
	
	 # Install required libraries (run in Colab / once in laptop)
	 !pip install pypandoc pdfplumber python-docx pandas openpyxl pillow googletrans==4.0.0-rc1
	 !pip install deep-translator
	 !apt-get install pandoc
	 !pip install pypandoc
	 !pip install reportlab

# How to Use

	- **Colab**

		- Upload your files → Run cells → Call functions.
		- Example:
		
			pdf_to_text("input.pdf", "output.txt")
			translate_file("input.pdf", "translated.txt", src_lang="en", dest_lang="fr")
		
		
	- **Offline (Laptop)**

		- Install dependencies:
		
			pip install pypandoc pdfplumber python-docx pandas openpyxl pillow googletrans==4.0.0-rc1

		
	- Run the script with python file_toolkit.py

# 🚀 Usage (CLI Tool in Colab / Local)

<small>	This tool runs as a CLI menu.

When you run the program, you’ll see options to convert a TXT file into multiple formats.</small>

```

python file_toolkit_collab.py

```

# Sample CLI Flow:

	
	==== TXT File Converter ====
	1. Convert TXT File
	2. Exit

	Choose: 1
	Upload file: myfile.txt

	Supported conversions:
	1. TXT → PDF
	2. TXT → DOCX
	3. TXT → PNG (image)
	4. TXT → CSV
	5. TXT → XLSX
	6. TXT → JSON

	Select target format (1-6): 5
	✅ Converted successfully: /content/csv.xlsx

	
## 🔹 1. Source & Target Language Codes

	- You type language codes (ISO-639-1):
	
		| Language              | Code    |
		| --------------------- | ------- |
		| English               | `en`    |
		| Hindi                 | `hi`    |
		| French                | `fr`    |
		| German                | `de`    |
		| Greek                 | `el`    |
		| Chinese (Simplified)  | `zh-CN` |
		| Chinese (Traditional) | `zh-TW` |
		| Spanish               | `es`    |
		| Russian               | `ru`    |
		| Japanese              | `ja`    |
		| Korean                | `ko`    |
		| Arabic                | `ar`    |
		| Italian               | `it`    |
	
	- 👉 Example:
	
		-  src_lang = en, dest_lang = hi → English ➝ Hindi
	
		-  src_lang = auto, dest_lang = fr → Auto detect ➝ French
	

## 🔹 2. Where to See the Translated File in Colab

- After running, it saved your file as:

			file_name.txt 

	- 📍 Location: Colab working directory (/content/)

		- To check:

			!ls
	
		- To open & read inside Colab:
			
			with open("B20442415 (5)_hi.txt", "r", encoding="utf-8") as f:
				print(f.read()[:500])   # show first 500 chars
			
		- To download to your laptop:
			
			from google.colab import files
				files.download("B20442415 (5)_hi.txt")
			

# Error while translation developer might face
	
- **🔹 Why It Didn’t Translate**

	1. If the input text is too long, Google rejects it. 
	
	2. If you pass the entire PDF/doc content at once, it fails silently. 
	
	3. Some versions of googletrans are broken in Colab. 
	
	
- **🔹 Solution: Split Text into Chunks**

	- We will split the text into smaller pieces, translate each, then join them.

			from deep_translator import GoogleTranslator

			def translate_text(text, src_lang="auto", dest_lang="en"):
				translated_chunks = []
				# Split into ~4000 character chunks (Google safe limit)
				chunk_size = 4000
				for i in range(0, len(text), chunk_size):
					chunk = text[i:i+chunk_size]
					translated = GoogleTranslator(source=src_lang, target=dest_lang).translate(chunk)
					translated_chunks.append(translated)
				return "\n".join(translated_chunks)
		

# 📌 Badges

![Build Status](https://img.shields.io/github/actions/workflow/status/alok-kumar8765/Universal-File-Converter-Translator-Python-CLI-/ci.yml?branch=main)  
![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)  
![License](https://img.shields.io/github/license/alok-kumar8765/Universal-File-Converter-Translator-Python-CLI-)  
![Code Coverage](https://img.shields.io/codecov/c/github/alok-kumar8765/Universal-File-Converter-Translator-Python-CLI-)  
![Issues](https://img.shields.io/github/issues/alok-kumar8765/Universal-File-Converter-Translator-Python-CLI-)  
![Pull Requests](https://img.shields.io/github/issues-pr/alok-kumar8765/Universal-File-Converter-Translator-Python-CLI-)  
![Last Commit](https://img.shields.io/github/last-commit/alok-kumar8765/Universal-File-Converter-Translator-Python-CLI-)  

# 📦 Docker Support

**Build:	bash**

```
docker build -t file-toolkit .
```

**Run:		bash**

```
docker run -it file-toolkit
```

# 🛠️ Development & Contributing


	- Fork the repo

	- Add new converters in **unified_converter.py**

	- Create pull requests

# 🌍 Roadmap
 
 
	- OCR for Images

	- Web UI (Streamlit/Gradio)

	- HuggingFace Spaces demo
 
 
 # Google Colab 
 
 ```
 https://colab.research.google.com/drive/1mi_qTwWfNjn-Vsk-h0vZAyu57k9c1wIP?usp=sharing
 ```
 
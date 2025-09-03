
# ✅ 1. Create Virtualenv & Django Project

	. Inside
		- /file-convertor_&_translator/
		
	. ✅ Create virtual environment (Windows)
		
		- # In Windows Powershell:
		- cd Universal-File-Converter-Translator-Python-CLI-

		- # Create venv
		- python -m venv venv

		- # Activate venv
		- .\venv\Scripts\activate

	. ✅ Install Requirements

		- pip install requirements.txt

	
	. Run
	
	
		- python -m venv venv
		- source venv/bin/activate         # On Linux/Mac
		- venv\Scripts\activate            # On Windows

---

# Real Desktop App (EXE)

	
	- **1.** Has a GUI (Tkinter/PyQt) with file upload buttons, not typing paths. 
	- **2.** Works as an EXE installer (download from GitHub → install like normal software).
	- **3.** Runs offline translation (no internet needed, unlike Google Translate).
	

	- There is a Tkinter-based file v1(basic script) conversion + translation tool:
	- This GUI v1 support less that 1 mb file for conversion and support offline translation, for offline translation we as using a model "Helsinki-NLP model".
	
	## 🔹 2. Making it an .EXE (for GitHub release)

		| Steps:

			| 1. Save as toolkit.py

			| 2. Install pyinstaller:
				``` pip install pyinstaller ```
	
			| 3. Build EXE:
				``` pyinstaller --onefile --windowed toolkit.py	```
				``` --onefile → single .exe ```
				``` --windowed → no console popup ```
		| It will create dist/toolkit.exe → .
	
	## 🔹 3. Offline Translation
	
		- Using Helsinki-NLP/MarianMT models (small but powerful).
		
		- First run downloads the model (100–300 MB) and caches it.
		
		- After that → works 100% offline.
		
		- Example: en → hi, en → fr, fr → de, etc.
		
		- **Note : ** The only catch: first run will download the translation model ***(unless we package it with the EXE)***.
		

# Fonts Download

	```
	import requests

	url_regular = "https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoSans/NotoSans-Regular.ttf"
	url_devanagari = "https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoSansDevanagari/NotoSansDevanagari-Medium.ttf"
	output_regular = "NotoSans-Regular.ttf"
	output_devanagari = "NotoSansDevanagari-Medium.ttf"

	def download_font(url, output_path):
		r = requests.get(url, allow_redirects=True)
		if r.status_code == 200:
			with open(output_path, "wb") as f:
				f.write(r.content)
			print(f"✅ Font downloaded: {output_path}")
			return True
		else:
			print(f"❌ Download failed for {url}: {r.status_code}")
			return False

	download_font(url_regular, output_regular)
	download_font(url_devanagari, output_devanagari)

	```
# Manual Font Download

``` # Browser open karo.

#URL open karo:

https://github.com/googlefonts/noto-fonts/blob/main/hinted/ttf/NotoSans/NotoSans-Regular.ttf

#Download raw file (button ya Download option hoga).

#File ko apne project folder me NotoSans-Regular.ttf ke naam se save kar lo.```


# 📌 Why Poppler is Required?

Some conversions in this tool (especially **PDF → Image**) rely on the `pdf2image` library.  
`pdf2image` internally uses **Poppler** to read and render PDF pages into images.  

👉 Without Poppler, you may see errors like:  
	| ⚠️ Error: Unable to get page count. Is poppler installed and in PATH?


So, installing Poppler is necessary for smooth PDF → PNG/JPG conversions.

---

## ⚙️ Poppler Installation Guide

	### 🖥 Windows
	
	1. Download latest Poppler build from:  
	   👉 [Poppler for Windows (Releases)](https://github.com/oschwartz10612/poppler-windows/releases/)
	   
	2. Extract the ZIP file (e.g., `poppler-xx/bin/`).
	
	3. Copy the **path of the `bin` folder** (e.g., `C:\poppler-xx\bin`).
	
	4. Add it to **System PATH**:  
	
	- Press `Win + R` → type `sysdm.cpl` → *Advanced* → *Environment Variables*.  

	   - Edit `Path` → Add new → Paste your Poppler `bin/` path.  
	
	5. Restart terminal / IDE. ✅

---

### 🍏 macOS

Poppler can be installed using [Homebrew](https://brew.sh/):  

	```bash
	brew install poppler
	```

### 🐧 Linux (Debian / Ubuntu)

Install via apt:
	```
	sudo apt-get update
	sudo apt-get install -y poppler-utils

	```

### ☁️ Google Colab

In Colab, just run this one-liner before using the tool:
	```
	!apt-get install -y poppler-utils

	```

### ✅ Verification

After installation, check if Poppler is working:
	```
	pdftoppm -h

	```


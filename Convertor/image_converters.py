# image_converters.py
import os
from typing import List, Union
from PIL import Image, ImageEnhance, ImageFilter
import pytesseract
from langdetect import detect
import langid
import easyocr # Import easyocr

# Initialize EasyOCR reader (load models once)
# Specify languages you expect to encounter, e.g., ['en', 'hi', 'es', 'fr', 'de']
# Add more languages as needed. 'en' is default.
# Download language models if they are not present.
_EASYOCR_READER = None
_EASYOCR_LANGS = []

def _get_easyocr(langs=['en']):
    global _EASYOCR_READER, _EASYOCR_LANGS
    # Re-initialize if reader is not set or if requested languages are different from currently loaded languages
    if _EASYOCR_READER is None or set(langs) != set(_EASYOCR_LANGS):
        try:
            import easyocr
            # The Reader constructor takes a list of languages directly
            _EASYOCR_READER = easyocr.Reader(langs, gpu=False) # Set gpu=True if you have a GPU
            _EASYOCR_LANGS = langs # Update the stored languages
            print(f"✅ Initialized EasyOCR with languages: {langs}")
        except Exception as e:
            print(f"Error initializing EasyOCR with languages {langs}: {e}")
            _EASYOCR_READER = None
            _EASYOCR_LANGS = [] # Clear languages on failure
    return _EASYOCR_READER

# Preprocessing function (optional but can help)
def preprocess_image(image_path, strong=False):
    img = Image.open(image_path).convert("RGB")
    if strong:
        img = img.filter(ImageFilter.MedianFilter(3))
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(2)
        enhancer = ImageEnhance.Brightness(img)
        img = enhancer.enhance(1.5)
    return img

def detect_language_from_image(img):
    """
    Quick heuristic: run a tiny OCR (English + few langs),
    then detect script with langdetect/langid.
    """
    try:
        # Use a basic config for quick detection
        tmp_text = pytesseract.image_to_string(img, lang="eng+hin", config="--psm 6 --oem 3")
        if not tmp_text.strip():
            return "en" # Default to English if no text detected

        # Use langid for better language detection on short text
        code, _ = langid.classify(tmp_text)
        return code
    except Exception as e:
        print(f"Language detection failed: {e}")
        return "en" # Fallback to English on error

def _langs_for_tesseract(easyocr_langs):
    """Maps EasyOCR language codes to Tesseract language codes."""
    tess_map = {
        'en': 'eng', 'hi': 'hin', 'es': 'spa', 'fr': 'fra', 'de': 'deu',
        'ru': 'rus', 'ja': 'jpn', 'ko': 'kor', 'ch_sim': 'chi_sim', 'ch_tra': 'chi_tra'
    }
    tess_langs = [tess_map.get(lang, lang) for lang in easyocr_langs]
    return "+".join(tess_langs)

def _parse_langs_for_easyocr(lang_input):
    """Parses comma-separated or auto language input for EasyOCR."""
    if lang_input.lower() == 'auto':
        # For auto-detection, still need to give EasyOCR *some* languages to load.
        # Let's default to English and Hindi for a common use case, but this could be expanded.
        return ['en', 'hi']
    return [l.strip() for l in lang_input.split(',')]


def image_to_txt_ocr(in_path, txt_path, lang="auto"):
    """
    Extracts text from an image using OCR. Supports multiple languages.
    Uses EasyOCR and Tesseract, prioritizing EasyOCR for some languages.

    Args:
        in_path (str): Path to the input image file.
        txt_path (str): Path to the output text file.
        lang (str): Language code(s) (e.g., 'en', 'hi', 'en,hi'). Use 'auto' for auto-detection.

    Returns:
        str: Path to the output text file if successful, None otherwise.
    """
    try:
        if not os.path.exists(in_path):
            raise FileNotFoundError(f"❌ File not found: {in_path}")

        # Preprocess image
        img = preprocess_image(in_path, strong=False)

        # Handle 'auto' language detection
        target_langs = []
        if lang.lower() == 'auto':
            # For auto-detection, load default EasyOCR languages and then try to detect
            target_langs_for_easyocr = _parse_langs_for_easyocr(lang)
            reader = _get_easyocr(target_langs_for_easyocr) # Load initial languages
            if reader:
                 # Now try to detect language from the image
                 detected = detect_language_from_image(img)
                 print(f"🌐 Detected language: {detected}")
                 # Re-initialize EasyOCR with the detected language + English for robustness
                 target_langs = list(set([detected, 'en'])) # Use set to avoid duplicates
                 # Ensure the detected language is in the supported list for EasyOCR if not English
                 if detected != 'en' and detected not in ['hi', 'es', 'fr', 'de', 'ru', 'ja', 'ko', 'ch_sim', 'ch_tra']: # Add more supported languages if needed
                      print(f"⚠️ Detected language {detected} might not be fully supported by EasyOCR. Using English.")
                      target_langs = ['en'] # Fallback to English
                 reader = _get_easyocr(target_langs) # Re-initialize with detected/fallback languages
            else:
                 # If EasyOCR couldn't even initialize with defaults
                 target_langs = ['en'] # Fallback to Tesseract with English
        else:
            target_langs = _parse_langs_for_easyocr(lang)
            reader = _get_easyocr(target_langs) # Initialize with specified languages


        text = ""
        used_engine = ""

        # --- Attempt with EasyOCR ---
        # Ensure reader is not None before using
        if reader:
            try:
                # Use paragraph=True for better formatting
                results = reader.readtext(img, detail=0, paragraph=True)
                text = "\n".join(results).strip()
                used_engine = "EasyOCR"
                print(f"✅ OCR extracted text using EasyOCR.")
            except Exception as e:
                print(f"⚠️ EasyOCR failed: {e}")


        # --- Fallback or additional pass with Tesseract ---
        if not text or used_engine != "EasyOCR": # If EasyOCR failed or was skipped
             tess_langs = _langs_for_tesseract(target_langs)
             print(f"Attempting Tesseract with languages: {tess_langs}")
             try:
                # Ensure tesseract_cmd is set if not in PATH (common in Colab)
                # pytesseract.tesseract_cmd = r'/usr/bin/tesseract' # Uncomment if needed
                text = pytesseract.image_to_string(img, lang=tess_langs, config="--oem 3 --psm 6").strip()
                used_engine = "Tesseract"
                print(f"✅ OCR extracted text using Tesseract.")
             except pytesseract.TesseractNotFoundError:
                 print("❌ Tesseract is not installed or not in PATH. Please install it.")
                 text = "" # Ensure text is empty if Tesseract is not found
             except Exception as e:
                print(f"⚠️ Tesseract failed for languages {tess_langs}: {e}")
                text = "" # Ensure text is empty if Tesseract also fails


        if not text:
            raise RuntimeError("OCR produced empty text from both engines or both failed.")

        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"✅ OCR extracted text saved: {txt_path} (Engine: {used_engine})")
        return txt_path

    except FileNotFoundError as e:
         print(f"❌ File not found: {e}")
         return None
    except Exception as e:
        print(f"❌ OCR failed: {e}")
        return None


# ---------- Image format helpers ----------
def image_to_image(in_path, out_path, fmt):
    try:
        # fmt: "PNG", "JPEG", etc.
        img = Image.open(in_path)
        # Convert mode if needed for JPEG
        if fmt.upper() in ["JPEG", "JPG"] and img.mode in ("RGBA", "LA"):
            bg = Image.new("RGB", img.size, (255,255,255))
            bg.paste(img, mask=img.split()[-1])
            img = bg
        img.convert("RGB").save(out_path, fmt.upper())
        return out_path

    except Exception as e:
        raise RuntimeError(f"❌ Image conversion failed: {e}")



# ---- Tiny wrappers (unchanged names) ----
def jpg_to_png(in_path, out_path): return image_to_image(in_path, out_path, "PNG")
def png_to_jpg(in_path, out_path): return image_to_image(in_path, out_path, "JPEG")
def jpg_to_jpeg(in_path, out_path): return image_to_image(in_path, out_path, "JPEG")
def png_to_jpeg(in_path, out_path): return image_to_image(in_path, out_path, "JPEG")
def jpeg_to_png(in_path, out_path): return image_to_image(in_path, out_path, "PNG")
def jpeg_to_jpg(in_path, out_path): return image_to_image(in_path, out_path, "JPEG")
def gif_to_png(in_path, out_path): return image_to_image(in_path, out_path, "PNG")

# OCR wrappers (note: lang now supports "auto" / "en,hi" / "eng+hin")
def jpg_to_txt(in_path, out_path, lang="auto"):  return image_to_txt_ocr(in_path, out_path, lang)
def jpeg_to_txt(in_path, out_path, lang="auto"): return image_to_txt_ocr(in_path, out_path, lang)
def png_to_txt(in_path, out_path, lang="auto"):  return image_to_txt_ocr(in_path, out_path, lang)
def gif_to_txt(in_path, out_path, lang="auto"):  return image_to_txt_ocr(in_path, out_path, lang)
def tiff_to_txt(in_path, out_path, lang="auto"): return image_to_txt_ocr(in_path, out_path, lang)
def bmp_to_txt(in_path, out_path, lang="auto"):  return image_to_txt_ocr(in_path, out_path, lang)


# ---------------- CLI MENU ---------------- #
def main():
    try:
        from google.colab import files
    except ImportError:
        files = None

    while True:
        print("\n==== Image File Converter =====")
        print("1. Convert Image File")
        print("2. OCR (Image → TXT)")
        print("3. Exit")
        choice = input("Enter choice: ").strip()

        if choice == "1":
            # Upload file (Colab) or enter path (offline)
            if files:
                uploaded = files.upload()
                img_file = list(uploaded.keys())[0]
            else:
                img_file = input("Enter path of image file (jpg/png/jpeg/gif): ").strip()
            
            print("\nSupported conversions: ")
            print("1. JPG → PNG")
            print("2. PNG → JPG")
            print("3. JPG → JPEG")
            print("4. PNG → JPEG")
            print("5. JPEG → PNG")
            print("6. JPEG → JPG")
            print("7. GIF → PNG")
            fmt_choice = input("Select target format (1-7): ").strip()

            base, _ = os.path.splitext(img_file)
            out_file = None
            try:
                if fmt_choice == "1": out_file = base + ".png";  jpg_to_png(img_file, out_file)
                elif fmt_choice == "2": out_file = base + ".jpg"; png_to_jpg(img_file, out_file)
                elif fmt_choice == "3": out_file = base + ".jpeg"; jpg_to_jpeg(img_file, out_file)
                elif fmt_choice == "4": out_file = base + ".jpeg"; png_to_jpeg(img_file, out_file)
                elif fmt_choice == "5": out_file = base + ".png";  jpeg_to_png(img_file, out_file)
                elif fmt_choice == "6": out_file = base + ".jpg";  jpeg_to_jpg(img_file, out_file)
                elif fmt_choice == "7": out_file = base + ".png";  gif_to_png(img_file, out_file)
                else:
                    print("❌ Invalid choice!")
                    continue
                print(f"✅ Converted successfully: {out_file}")
                if files and out_file: files.download(out_file)
            except Exception as e:
                print("❌ Conversion failed:", e)

        elif choice == "2":
            img_file = input("Enter path of image: ").strip()
            out_file = os.path.splitext(img_file)[0] + ".txt"
            lang_input =  "auto"
            try:
                res = image_to_txt_ocr(img_file, out_file, lang=lang_input) # Pass lang_input here
                if res:
                    print(f"✅ OCR done: {res}")
                    if files: files.download(res)
            except Exception as e:
                print("❌ OCR failed:", e)

        elif choice == "3":
            print("👋 Exiting...")
            break
        else:
            print("❌ Invalid choice!")

if __name__ == "__main__":
    main()
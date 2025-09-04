import pytest
import os
import shutil
import json
from unittest.mock import patch, MagicMock # Import MagicMock
import pytest
from unittest.mock import patch
import Converter.universal_converter as uc  # import your module

# Assuming the converter functions are in the current notebook cell (or imported)

# --- Helper to create dummy files ---
def create_dummy_csv(filename="dummy.csv"):
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["col1", "col2", "col3"])
        writer.writerow(["data1", "data2", "data3"])
        writer.writerow(["value 😅", "हिंदी", "another,value"])
    return filename

def create_dummy_xlsx(filename="dummy.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.append(["colA", "colB", "colC"])
    ws.append(["dataA", "dataB", "dataC"])
    ws.append(["value 😊", "मराठी", "one|two|three"])
    wb.save(filename)
    return filename

def create_dummy_txt(filename="dummy.txt"):
    content = """
This is a test text file.
It has multiple lines.
Including some Unicode characters: नमस्ते
And a blank line above.
"""
    with open(filename, "w", encoding="utf-8") as f:
        f.write(content)
    return filename

def create_dummy_docx(filename="dummy.docx"):
    doc = Document()
    doc.add_paragraph("This is a test document.")
    doc.add_paragraph("It also has text in Hindi: नमस्ते")
    table = doc.add_table(rows=2, cols=3)
    table.cell(0, 0).text = "Header 1"
    table.cell(0, 1).text = "Header 2"
    table.cell(0, 2).text = "Header 3"
    table.cell(1, 0).text = "Row 1 Data 1"
    table.cell(1, 1).text = "Row 1 Data 2 😊"
    table.cell(1, 2).text = "रो 1 डेटा 3"
    doc.save(filename)
    return filename

def create_dummy_pdf(filename="dummy.pdf"):
    # Using reportlab to create a simple PDF
    c = canvas.Canvas(filename, pagesize=A4)
    c.drawString(100, 750, "Test PDF Content")
    c.drawString(100, 735, "With some Hindi: नमस्ते")
    # Add a simple "table" like structure
    c.drawString(100, 700, "Col1 | Col2 | Col3")
    c.drawString(100, 685, "Data1 | Data2 | Data3")
    c.save()
    return filename

def create_dummy_json(filename="dummy.json"):
    data = [
        {"name": "Alice", "age": 30, "city": "New York"},
        {"name": "Bob", "age": 25, "city": "London", "greeting": "नमस्ते"}
    ]
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return filename

def create_dummy_image(filename="dummy.png"):
    img = Image.new('RGB', (100, 50), color = (255, 255, 255))
    d = ImageDraw.Draw(img)
    # Use a font that supports Unicode if possible
    try:
        font = ImageFont.truetype("/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf", 15)
    except Exception:
         font = ImageFont.load_default()
    d.text((10,10), "Test Image नमस्ते", fill=(0,0,0), font=font)
    img.save(filename)
    return filename

# --- Fixture for temporary directory ---
@pytest.fixture
def temp_dir(request):
    tmpdir = tempfile.mkdtemp()
    def cleanup():
        shutil.rmtree(tmpdir)
    request.addfinalizer(cleanup)
    return tmpdir

# --- Mock the schedule_delete function ---
@pytest.fixture(autouse=True)
def mock_schedule_delete():
    with patch('Converter.universal_converter.schedule_delete') as mock_delete:
        yield mock_delete

# --- Test Cases ---

# CSV Conversions
def test_csv_to_xls(temp_dir):
    csv_file = create_dummy_csv(os.path.join(temp_dir, "test.csv"))
    xls_file = os.path.join(temp_dir, "output.xlsx")
    result = csv_to_xls(csv_file, xls_file)
    assert os.path.exists(result)
    assert os.path.getsize(result) > 0
    # Basic check on content (optional, requires openpyxl)
    wb = load_workbook(result)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    assert len(rows) == 3
    assert rows[0] == ("col1", "col2", "col3")

def test_csv_to_pdf(temp_dir):
    csv_file = create_dummy_csv(os.path.join(temp_dir, "test.csv"))
    pdf_file = os.path.join(temp_dir, "output.pdf")
    result = csv_to_pdf(csv_file, pdf_file)
    assert os.path.exists(result)
    assert os.path.getsize(result) > 0

def test_csv_to_doc(temp_dir):
    csv_file = create_dummy_csv(os.path.join(temp_dir, "test.csv"))
    docx_file = os.path.join(temp_dir, "output.docx")
    result = csv_to_doc(csv_file, docx_file)
    assert os.path.exists(result)
    assert os.path.getsize(result) > 0
    # Basic check on content (optional, requires python-docx)
    # doc = Document(result)
    # assert len(doc.tables) > 0

def test_csv_to_txt(temp_dir):
    csv_file = create_dummy_csv(os.path.join(temp_dir, "test.csv"))
    txt_file = os.path.join(temp_dir, "output.txt")
    result = csv_to_txt(csv_file, txt_file, delimiter=",") # Test with comma delimiter
    assert os.path.exists(result)
    with open(result, "r", encoding="utf-8") as f:
        content = f.read()
    assert "col1,col2,col3" in content
    assert "value 😅,हिंदी,another,value" in content # Check for delimiter and content

def test_csv_to_json(temp_dir):
    csv_file = create_dummy_csv(os.path.join(temp_dir, "test.csv"))
    json_file = os.path.join(temp_dir, "output.json")
    result = csv_to_json(csv_file, json_file)
    assert os.path.exists(result)
    with open(result, "r", encoding="utf-8") as f:
        data = json.load(f)
    assert isinstance(data, list)
    assert len(data) == 2 # Header row is not included in data
    assert data[0]["col1"] == "data1"
    assert data[1]["col2"] == "हिंदी"

# XLS Conversions
def test_xls_to_csv(temp_dir):
    xls_file = create_dummy_xlsx(os.path.join(temp_dir, "test.xlsx"))
    csv_file = os.path.join(temp_dir, "output.csv")
    result = xls_to_csv(xls_file, csv_file)
    assert os.path.exists(result)
    assert os.path.getsize(result) > 0
    with open(result, "r", encoding="utf-8") as f:
        content = f.read()
    assert "colA,colB,colC" in content
    assert "value 😊,मराठी,one|two|three" in content

def test_xls_to_doc(temp_dir):
    xls_file = create_dummy_xlsx(os.path.join(temp_dir, "test.xlsx"))
    docx_file = os.path.join(temp_dir, "output.docx")
    result = xls_to_doc(xls_file, docx_file)
    assert os.path.exists(result)
    assert os.path.getsize(result) > 0

def test_xls_to_txt(temp_dir):
    xls_file = create_dummy_xlsx(os.path.join(temp_dir, "test.xlsx"))
    txt_file = os.path.join(temp_dir, "output.txt")
    result = xls_to_txt(xls_file, txt_file, delimiter="|") # Test with pipe delimiter
    assert os.path.exists(result)
    with open(result, "r", encoding="utf-8") as f:
        content = f.read()
    assert "colA|colB|colC" in content
    assert "value 😊|मराठी|one|two|three" in content

def test_xls_to_pdf(temp_dir):
    xls_file = create_dummy_xlsx(os.path.join(temp_dir, "test.xlsx"))
    pdf_file = os.path.join(temp_dir, "output.pdf")
    result = xls_to_pdf(xls_file, pdf_file)
    assert os.path.exists(result)
    assert os.path.getsize(result) > 0

def test_xls_to_json(temp_dir):
    xls_file = create_dummy_xlsx(os.path.join(temp_dir, "test.xlsx"))
    json_file = os.path.join(temp_dir, "output.json")
    result = xls_to_json(xls_file, json_file)
    assert os.path.exists(result)
    with open(result, "r", encoding="utf-8") as f:
        data = json.load(f)
    assert isinstance(data, list)
    assert len(data) == 2 # Header row is not included in data
    assert data[0]["colA"] == "dataA"
    assert data[1]["colB"] == "मराठी"

# Image Conversions (Basic tests, OCR accuracy varies)
@patch('__main__._get_easyocr', return_value=MagicMock()) # Mock EasyOCR initialization
@patch('pytesseract.image_to_string') # Mock Tesseract
def test_image_to_txt_ocr(mock_tesseract, mock_easyocr_reader, temp_dir):
    # Configure mocks
    mock_easyocr_reader.return_value.readtext.return_value = ["Mocked EasyOCR Text नमस्ते"]
    mock_tesseract.return_value = "Mocked Tesseract Text हिंदी"

    img_file = create_dummy_image(os.path.join(temp_dir, "test.png"))
    txt_file = os.path.join(temp_dir, "output.txt")

    # Test with EasyOCR (if enabled)
    result = image_to_txt_ocr(img_file, txt_file, lang='en,hi')
    assert os.path.exists(result)
    with open(result, "r", encoding="utf-8") as f:
        content = f.read()
    # Check if either engine's output is in the file
    assert "Mocked EasyOCR Text नमस्ते" in content or "Mocked Tesseract Text हिंदी" in content

    # Test fallback to Tesseract if EasyOCR fails (simulate EasyOCR failure)
    mock_easyocr_reader.return_value.readtext.side_effect = Exception("EasyOCR Error")
    result = image_to_txt_ocr(img_file, txt_file, lang='en,hi')
    assert os.path.exists(result)
    with open(result, "r", encoding="utf-8") as f:
        content = f.read()
    assert "Mocked Tesseract Text हिंदी" in content # Should contain Tesseract output

def test_image_to_image(temp_dir):
    img_file = create_dummy_image(os.path.join(temp_dir, "test.png"))
    jpg_file = os.path.join(temp_dir, "output.jpg")
    result = image_to_image(img_file, jpg_file, "JPEG")
    assert os.path.exists(result)
    assert os.path.getsize(result) > 0
    assert result.lower().endswith(".jpg")

# TXT Conversions
def test_txt_to_pdf(temp_dir):
    txt_file = create_dummy_txt(os.path.join(temp_dir, "test.txt"))
    pdf_file = os.path.join(temp_dir, "output.pdf")
    result = txt_to_pdf(txt_file, pdf_file)
    assert os.path.exists(result)
    assert os.path.getsize(result) > 0

@patch('__main__.Document') # Mock Document to avoid actual file creation/dependencies
def test_txt_to_doc(mock_document, temp_dir):
    txt_file = create_dummy_txt(os.path.join(temp_dir, "test.txt"))
    docx_file = os.path.join(temp_dir, "output.docx")
    # Configure the mock Document object
    mock_doc_instance = MagicMock()
    mock_document.return_value = mock_doc_instance

    result = txt_to_doc(txt_file, docx_file)

    # Assert that Document was called and save was called
    mock_document.assert_called_once()
    mock_doc_instance.save.assert_called_once_with(docx_file)
    # We can't easily assert file content without actual docx creation

def test_txt_to_json(temp_dir):
    txt_file = create_dummy_txt(os.path.join(temp_dir, "test.txt"))
    json_file = os.path.join(temp_dir, "output.json")
    result = txt_to_json(txt_file, json_file)
    assert os.path.exists(result)
    with open(result, "r", encoding="utf-8") as f:
        data = json.load(f)
    assert isinstance(data, list)
    assert "This is a test text file." in data

# DOC/DOCX Conversions (Require LibreOffice or heavy mocking)
@patch('__main__.convert_doc_to_docx_if_needed', return_value='mocked_input.docx')
@patch('__main__.convert_docx_to_pdf_libreoffice')
@patch('__main__.mammoth.convert_to_html')
@patch('__main__.pdfkit.from_file')
def test_doc_to_pdf_auto(mock_pdfkit, mock_mammoth, mock_libreoffice_pdf, mock_doc_to_docx, temp_dir):
    dummy_doc = os.path.join(temp_dir, "test.doc")
    output_pdf = os.path.join(temp_dir, "output.pdf")

    # Mock docx_path to exist for subsequent steps
    mock_doc_to_docx.return_value = os.path.join(temp_dir, "mocked_input.docx")
    # Create a dummy mocked docx file for the table check
    mock_docx_file = os.path.join(temp_dir, "mocked_input.docx")
    # We need a minimal docx file to avoid errors when python-docx tries to open it
    # A simple empty docx is enough if we mock the table check later
    Document().save(mock_docx_file)


    # Mock the table inspection within doc_to_pdf
    with patch('__main__.Document') as mock_Document_for_table_check:
        mock_doc_instance = MagicMock()
        mock_Document_for_table_check.return_value = mock_doc_instance

        # Case 1: No wide tables -> LibreOffice path
        mock_doc_instance.tables = [] # No tables detected
        mock_libreoffice_pdf.return_value = os.path.join(temp_dir, "libreoffice_output.pdf")
        # Create the mocked LibreOffice output file
        with open(os.path.join(temp_dir, "libreoffice_output.pdf"), 'w') as f:
            f.write("mock pdf content")

        result = doc_to_pdf(dummy_doc, output_pdf, mode="auto")
        mock_libreoffice_pdf.assert_called_once()
        assert os.path.exists(output_pdf)
        # Check if the content is from the mocked LibreOffice PDF
        with open(output_pdf, 'r') as f:
            assert "mock pdf content" in f.read()

        # Reset mocks for next case
        mock_libreoffice_pdf.reset_mock()
        mock_mammoth.reset_mock()
        mock_pdfkit.reset_mock()
        os.remove(output_pdf) # Clean up output

        # Case 2: Wide tables -> HTML path
        mock_doc_instance = MagicMock()
        mock_Document_for_table_check.return_value = mock_doc_instance
        # Simulate a wide table
        mock_table = MagicMock()
        mock_table.rows = [MagicMock()] # Need at least one row
        mock_table.rows[0].cells = [MagicMock()] * 5 # 5 columns
        mock_doc_instance.tables = [mock_table]

        mock_mammoth.return_value.value = "<html><body>Mock HTML Content</body></html>"
        mock_pdfkit.return_value = True # pdfkit returns True on success

        result = doc_to_pdf(dummy_doc, output_pdf, mode="auto")
        mock_mammoth.assert_called_once()
        mock_pdfkit.assert_called_once()
        # We can't easily check the final PDF content here as it's mocked

@patch('__main__.convert_doc_to_docx_if_needed', return_value='mocked_input.docx')
@patch('__main__.convert_docx_to_pdf_libreoffice')
@patch('__main__.convert_from_path')
@patch('PIL.Image.Image.save')
def test_doc_to_pdf_image(mock_image_save, mock_convert_from_path, mock_libreoffice_pdf, mock_doc_to_docx, temp_dir):
    dummy_doc = os.path.join(temp_dir, "test.doc")
    output_pdf = os.path.join(temp_dir, "output.pdf")

    # Mock docx_path to exist
    mock_doc_to_docx.return_value = os.path.join(temp_dir, "mocked_input.docx")

    # Mock LibreOffice PDF creation
    mock_libreoffice_pdf.return_value = os.path.join(temp_dir, "libreoffice_output.pdf")
    # Create the mocked LibreOffice output file
    with open(os.path.join(temp_dir, "libreoffice_output.pdf"), 'w') as f:
        f.write("mock pdf content")

    # Mock pdf2image conversion
    mock_image_instance = MagicMock()
    mock_convert_from_path.return_value = [mock_image_instance] # Simulate one page

    result = doc_to_pdf(dummy_doc, output_pdf, mode="image")

    mock_doc_to_docx.assert_called_once_with(dummy_doc)
    mock_libreoffice_pdf.assert_called_once()
    mock_convert_from_path.assert_called_once()
    mock_image_save.assert_called_once_with(output_pdf, save_all=True, append_images=[])
    assert os.path.exists(output_pdf) # Check if the final output file was created

# PDF Conversions
@patch('pdfplumber.open')
def test_pdf_to_txt(mock_pdfplumber_open, temp_dir):
    dummy_pdf = os.path.join(temp_dir, "test.pdf")
    output_txt = os.path.join(temp_dir, "output.txt")

    # Mock pdfplumber
    mock_pdf = MagicMock()
    mock_page1 = MagicMock()
    mock_page1.extract_text.return_value = "Page 1 Text नमस्ते"
    mock_page2 = MagicMock()
    mock_page2.extract_text.return_value = "Page 2 Text हिंदी"
    mock_pdf.pages = [mock_page1, mock_page2]
    mock_pdfplumber_open.return_value.__enter__.return_value = mock_pdf

    # Create a dummy input file for os.path.exists check (content doesn't matter)
    with open(dummy_pdf, 'w') as f: f.write("dummy")

    result = pdf_to_txt(dummy_pdf, output_txt)

    mock_pdfplumber_open.assert_called_once_with(dummy_pdf)
    mock_page1.extract_text.assert_called_once()
    mock_page2.extract_text.assert_called_once()
    assert os.path.exists(result)
    with open(result, 'r', encoding='utf-8') as f:
        content = f.read()
    assert "Page 1 Text नमस्ते" in content
    assert "Page 2 Text हिंदी" in content

@patch('pdfplumber.open')
@patch('__main__.Document')
def test_pdf_to_doc(mock_document, mock_pdfplumber_open, temp_dir):
    dummy_pdf = os.path.join(temp_dir, "test.pdf")
    output_docx = os.path.join(temp_dir, "output.docx")

    # Mock pdfplumber
    mock_pdf = MagicMock()
    mock_page1 = MagicMock()
    mock_page1.extract_text.return_value = "Page 1 Text नमस्ते"
    mock_page1.images = [] # No images for this test
    mock_page2 = MagicMock()
    mock_page2.extract_text.return_value = "Page 2 Text हिंदी"
    mock_page2.images = []
    mock_pdf.pages = [mock_page1, mock_page2]
    mock_pdfplumber_open.return_value.__enter__.return_value = mock_pdf
    mock_pdfplumber_open.return_value.__exit__.return_value = None # Mock context manager exit

    # Mock python-docx Document
    mock_doc_instance = MagicMock()
    mock_document.return_value = mock_doc_instance

    # Create a dummy input file for os.path.exists check
    with open(dummy_pdf, 'w') as f: f.write("dummy")

    result = pdf_to_doc(dummy_pdf, output_docx)

    mock_pdfplumber_open.assert_called_once_with(dummy_pdf)
    mock_document.assert_called_once()
    # Check if add_paragraph was called for each line (roughly)
    assert mock_doc_instance.add_paragraph.call_count >= 2 # At least one paragraph per page
    mock_doc_instance.save.assert_called_once_with(output_docx)
    # We can't easily assert file content without actual docx creation

@patch('pdf2image.convert_from_path')
@patch('PIL.Image.Image.save')
@patch('builtins.input', return_value='2') # Mock user input for multi-page choice
@patch('shutil.rmtree') # Mock rmtree to prevent cleanup issues
def test_pdf_to_image_multi_page_zip(mock_rmtree, mock_input, mock_image_save, mock_convert_from_path, temp_dir):
    dummy_pdf = os.path.join(temp_dir, "test.pdf")
    output_dir = os.path.join(temp_dir, "output_images")

    # Mock pdf2image to return multiple images
    mock_img1 = MagicMock()
    mock_img1.size = (100, 100)
    mock_img2 = MagicMock()
    mock_img2.size = (100, 100)
    mock_convert_from_path.return_value = [mock_img1, mock_img2]

    # Create a dummy input file
    with open(dummy_pdf, 'w') as f: f.write("dummy")

    # Mock zipfile.ZipFile to prevent actual zipping during the test
    with patch('zipfile.ZipFile') as mock_zipfile:
        mock_zip_instance = MagicMock()
        mock_zipfile.return_value.__enter__.return_value = mock_zip_instance

        result = pdf_to_image(dummy_pdf, output_dir, fmt="png")

        mock_convert_from_path.assert_called_once_with(dummy_pdf, dpi=150)
        # Check if save was called for each image
        assert mock_image_save.call_count == 2
        # Check if zipfile was called
        mock_zipfile.assert_called_once()
        # Check if rmtree was called for cleanup
        mock_rmtree.assert_called_once_with(output_dir)
        # Check if the result is the zip path (adjust if your function returns dir path)
        # Based on the code, it returns the directory path, which is then zipped in main.
        # Let's check for the existence of the output directory instead for this mock test.
        assert os.path.exists(output_dir) # Directory is created before rmtree mock

@patch('pdfplumber.open')
@patch('csv.writer')
@patch('openpyxl.Workbook')
def test_pdf_to_csv_xlsx(mock_workbook, mock_csv_writer, mock_pdfplumber_open, temp_dir):
    dummy_pdf = os.path.join(temp_dir, "test.pdf")
    output_csv = os.path.join(temp_dir, "output.csv")
    output_xlsx = os.path.join(temp_dir, "output.xlsx")

    # Mock pdfplumber
    mock_pdf = MagicMock()
    mock_page1 = MagicMock()
    # Simulate a table with rows
    mock_page1.extract_tables.return_value = [[["Header1", "Header2"], ["Data1", "Data2"]]]
    mock_page1.extract_text.return_value = "" # No plain text
    mock_pdf.pages = [mock_page1]
    mock_pdfplumber_open.return_value.__enter__.return_value = mock_pdf
    mock_pdfplumber_open.return_value.__exit__.return_value = None # Mock context manager exit

    # Mock csv writer
    mock_csv_writer_instance = MagicMock()
    mock_csv_writer.return_value = mock_csv_writer_instance

    # Mock openpyxl Workbook
    mock_wb_instance = MagicMock()
    mock_ws_instance = MagicMock()
    mock_workbook.return_value = mock_wb_instance
    mock_wb_instance.create_sheet.return_value = mock_ws_instance

    # Create dummy input file
    with open(dummy_pdf, 'w') as f: f.write("dummy")

    result = pdf_to_csv(dummy_pdf, csv_path=output_csv, xlsx_path=output_xlsx)

    mock_pdfplumber_open.assert_called_once_with(dummy_pdf)
    mock_csv_writer.assert_called_once()
    mock_workbook.assert_called_once()
    # Check if rows were written to CSV and XLSX
    assert mock_csv_writer_instance.writerow.call_count == 2 # Header + Data
    assert mock_ws_instance.append.call_count == 2 # Header + Data
    mock_wb_instance.save.assert_called_once_with(output_xlsx)
    assert os.path.exists(output_csv)
    assert os.path.exists(output_xlsx)

# JSON Conversions
@patch('pdfkit.from_file')
def test_json_to_pdf(mock_pdfkit_from_file, temp_dir):
    json_file = create_dummy_json(os.path.join(temp_dir, "test.json"))
    pdf_file = os.path.join(temp_dir, "output.pdf")

    mock_pdfkit_from_file.return_value = True # Simulate success

    result = json_to_pdf(json_file, pdf_file)

    mock_pdfkit_from_file.assert_called_once()
    assert result == pdf_file # Assuming success returns the output path

@patch('__main__.Document')
def test_json_to_doc(mock_document, temp_dir):
    json_file = create_dummy_json(os.path.join(temp_dir, "test.json"))
    docx_file = os.path.join(temp_dir, "output.docx")

    mock_doc_instance = MagicMock()
    mock_document.return_value = mock_doc_instance

    result = json_to_doc(json_file, docx_file)

    mock_document.assert_called_once()
    # Check if add_heading and add_paragraph were called
    mock_doc_instance.add_heading.assert_called_once()
    assert mock_doc_instance.add_paragraph.call_count >= 1 # At least one paragraph for the data
    mock_doc_instance.save.assert_called_once_with(docx_file)
    assert result == docx_file

@patch('__main__.save_html_as_image') # Mock the helper image saving function
@patch('builtins.input', return_value='1') # Mock user input for single image choice
def test_json_to_image_single(mock_input, mock_save_html_as_image, temp_dir):
    json_file = create_dummy_json(os.path.join(temp_dir, "test.json"))
    png_file = os.path.join(temp_dir, "output.png")

    mock_save_html_as_image.return_value = png_file # Simulate success

    result = json_to_image(json_file, png_file)

    mock_save_html_as_image.assert_called_once_with(
        json.dumps(json.load(open(json_file, encoding='utf-8')), indent=4, ensure_ascii=False),
        png_file,
        long_mode=True
    )
    assert result == png_file

def test_json_to_txt(temp_dir):
    json_file = create_dummy_json(os.path.join(temp_dir, "test.json"))
    txt_file = os.path.join(temp_dir, "output.txt")

    result = json_to_txt(json_file, txt_file)

    assert os.path.exists(result)
    with open(result, "r", encoding="utf-8") as f:
        content = f.read()
    # Check for pretty printed content
    assert '"name": "Alice"' in content
    assert '"greeting": "नमस्ते"' in content

@patch('csv.DictWriter')
def test_json_to_csv(mock_dict_writer, temp_dir):
    json_file = create_dummy_json(os.path.join(temp_dir, "test.json"))
    csv_file = os.path.join(temp_dir, "output.csv")

    mock_writer_instance = MagicMock()
    mock_dict_writer.return_value = mock_writer_instance

    result = json_to_csv(json_file, csv_file)

    mock_dict_writer.assert_called_once() # Should be called with fieldnames from the first object
    mock_writer_instance.writeheader.assert_called_once()
    assert mock_writer_instance.writerow.call_count == 2 # Called for each data object
    assert result == csv_file

@patch('openpyxl.Workbook')
def test_json_to_xls(mock_workbook, temp_dir):
    json_file = create_dummy_json(os.path.join(temp_dir, "test.json"))
    xlsx_file = os.path.join(temp_dir, "output.xlsx")

    mock_wb_instance = MagicMock()
    mock_ws_instance = MagicMock()
    mock_workbook.return_value = mock_wb_instance
    mock_wb_instance.create_sheet.return_value = mock_ws_instance

    result = json_to_xls(json_file, xlsx_file)

    mock_workbook.assert_called_once()
    mock_wb_instance.create_sheet.assert_called_once()
    # Should append header and then each flattened row
    assert mock_ws_instance.append.call_count == 3 # Header + 2 data rows
    mock_wb_instance.save.assert_called_once_with(xlsx_file)
    assert result == xlsx_file
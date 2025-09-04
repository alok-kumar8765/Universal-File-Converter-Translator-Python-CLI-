Here are test cases for the individual conversion functions using `pytest`.

**Note:**
1. You need to have `pytest` installed (`!pip install pytest`).
2. These tests require dummy input files (e.g., `dummy.csv`, `dummy.xlsx`, `dummy.txt`, `dummy.docx`, `dummy.pdf`, `dummy.json`, `dummy.png`) in the same directory as the script or provide the correct paths.
3. Some tests might require external dependencies like LibreOffice (`soffice`) for `.doc` to `.docx` conversion, or specific fonts.
4. The image conversion tests might be sensitive to font availability and exact rendering, so they might need adjustments based on your environment.
5. The `schedule_delete` function is mocked to prevent actual file deletion during tests.
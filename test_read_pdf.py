import fitz
import pytest

def test_read_old_pdf():
    print("\n================ OLD PDF ==================")
    doc = fitz.open(r'c:\Users\aguirre.maurin\Documents\GitHub\OFBilan-Plugin-QGIS\temp\Bilan_sd21_2024.pdf')
    for i, page in enumerate(doc):
        print(f"--- PAGE {i+1} ---")
        print(page.get_text())

def test_read_new_pdf():
    print("\n================ NEW PDF ==================")
    doc = fitz.open(r'c:\Users\aguirre.maurin\Documents\GitHub\OFBilan-Plugin-QGIS\data\out\bilan_global_21\bilan_global_21_ext.pdf')
    for i, page in enumerate(doc):
        print(f"--- PAGE {i+1} ---")
        print(page.get_text())

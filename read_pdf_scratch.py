import fitz
import sys

def read_pdf(pdf_path):
    print(f"=== {pdf_path} ===")
    try:
        doc = fitz.open(pdf_path)
        for i, page in enumerate(doc):
            print(f"--- PAGE {i+1} ---")
            print(page.get_text())
    except Exception as e:
        print(f"Error fitz: {e}")
        try:
            import pypdf
            reader = pypdf.PdfReader(pdf_path)
            for i, page in enumerate(reader.pages):
                print(f"--- PAGE {i+1} ---")
                print(page.extract_text())
        except Exception as e2:
            print(f"Error pypdf: {e2}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        read_pdf(sys.argv[1])

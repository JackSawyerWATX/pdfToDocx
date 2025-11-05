import pdfplumber
from docx import Document

pdf_path = input("Enter the path to the PDF file: ").strip().strip('"').strip("'")
pdf = pdfplumber.open(pdf_path)
doc = Document()

for page in pdf.pages:
    text = page.extract_text()
    if text:
        doc.add_paragraph(text)

doc.save("PDF-output.docx")
pdf.close()

print("Document extracted and saved as 'PDF-output.docx'.")
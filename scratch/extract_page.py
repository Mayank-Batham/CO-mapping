import fitz  # PyMuPDF
import os

doc = fitz.open("RPA LAB rubrics.pdf")
print("Total Pages:", len(doc))

os.makedirs("scratch", exist_ok=True)
page = doc[0]
pix = page.get_pixmap(dpi=150)
pix.save("scratch/page_0.png")
print("Saved page_0.png successfully!")

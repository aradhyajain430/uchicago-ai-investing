import pymupdf
from pathlib import Path

pdf = Path(r"C:\Users\Dell\Desktop\uchicagoaicomp\slideshow\1-UPenn-Merge (1).pdf")
out = Path(r"C:\Users\Dell\Desktop\uchicagoaicomp\research\extracts\sprouts_pages")
out.mkdir(parents=True, exist_ok=True)
doc = pymupdf.open(pdf)
# Render representative layout pages
for i in [0, 1, 3, 4, 10, 11, 12, 13]:
    if i < doc.page_count:
        page = doc[i]
        pix = page.get_pixmap(matrix=pymupdf.Matrix(1.4, 1.4))
        pix.save(str(out / f"page_{i+1:02d}.png"))
        print("saved", i + 1, page.rect)
print("done", doc.page_count)

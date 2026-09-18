from pathlib import Path
import fitz
from PIL import Image

root = Path(__file__).resolve().parents[1]
pdf = root / "final_slideshow" / "Final_Lumentum_Bull_Thesis.pdf"
out = root / "final_slideshow" / "qa"
out.mkdir(exist_ok=True)
doc = fitz.open(pdf)
thumbs = []
for i, page in enumerate(doc):
    path = out / f"slide_{i+1:02d}.png"
    page.get_pixmap(matrix=fitz.Matrix(1.2, 1.2), alpha=False).save(path)
    thumbs.append(Image.open(path).convert("RGB").resize((400, 300)))
sheet = Image.new("RGB", (1200, 1200), "white")
for i, im in enumerate(thumbs):
    sheet.paste(im, ((i % 3) * 400, (i // 3) * 300))
sheet.save(out / "contact_sheet.png")
print(out / "contact_sheet.png")

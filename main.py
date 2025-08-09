import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
import os

# نام فایل اکسل
excel_file = "data.xlsx"  # اینجا اسم فایل اکسل رو بگذار

# خواندن داده‌ها از اکسل
df = pd.read_excel(excel_file)

# نام فایل خروجی PDF
output_pdf = "output.pdf"

# ایجاد PDF
c = canvas.Canvas(output_pdf, pagesize=A4)
page_width, page_height = A4

for index, row in df.iterrows():
    name_shape = str(row['Name Shape'])
    WT = row['WT']
    HR = row['HR']
    WB = row['WB']
    HL = row['HL']
    img_path = row.iloc[5]  # ستون F

    # عنوان بالای صفحه
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(page_width / 2, page_height - 2 * cm, name_shape)

    # اضافه کردن تصویر
    if os.path.exists(img_path):
        img = ImageReader(img_path)
        img_width, img_height = img.getSize()
        max_width = 10 * cm
        max_height = 10 * cm

        scale = min(max_width / img_width, max_height / img_height)
        img_width *= scale
        img_height *= scale

        img_x = (page_width - img_width) / 2
        img_y = (page_height - img_height) / 2
        c.drawImage(img, img_x, img_y, img_width, img_height)

        # نوشتن اندازه‌ها
        c.setFont("Helvetica", 12)
        # بالا (WT)
        c.drawCentredString(page_width / 2, img_y + img_height + 0.5 * cm, f"WT: {WT}")
        # پایین (WB)
        c.drawCentredString(page_width / 2, img_y - 1 * cm, f"WB: {WB}")
        # راست (HR)
        c.drawString(img_x + img_width + 0.5 * cm, img_y + img_height / 2, f"HR: {HR}")
        # چپ (HL)
        c.drawRightString(img_x - 0.5 * cm, img_y + img_height / 2, f"HL: {HL}")

    else:
        c.setFont("Helvetica", 12)
        c.setFillColorRGB(1, 0, 0)
        c.drawCentredString(page_width / 2, page_height / 2, "Image not found")

    c.showPage()

c.save()
print(f"PDF ساخته شد: {output_pdf}")

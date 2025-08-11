import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
import os

# ---------- تنظیمات ----------
TEMPLATES_DIR = "templates"     # پوشه تصاویر نمونه
EXCEL_FILE = "data.xlsx"        # فایل اکسل ورودی
OUTPUT_PDF = "output_samples.pdf"

IMAGES_PER_PAGE = 4             # 4 یا 6 معمولاً
if IMAGES_PER_PAGE == 4:
    ROWS, COLS = 2, 2
elif IMAGES_PER_PAGE == 6:
    ROWS, COLS = 3, 2
else:
    # fallback: محاسبه سطر/ستون ساده (دو ستون پیش‌فرض)
    COLS = 2
    ROWS = (IMAGES_PER_PAGE + COLS - 1) // COLS

TITLE_FONT = ("Helvetica-Bold", 12)
TEXT_FONT = ("Helvetica", 9)
MAX_IMG_SCALE_CM = 6            # حداکثر "بعد" داخل هر سلول (می‌تونی تغییر بدی)
DEFAULT_TEMPLATE = os.path.join(TEMPLATES_DIR, "default.png")  # اگر مدل موجود نبود

# ثبت فونت یونی‌کد برای نمایش فارسی (در صورت نیاز)
try:
    pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
except Exception:
    pass  # اگر نشد، فونت پیش‌فرض به کار میرود

# ---------- توابع کمکی ----------
def find_template_for_name(name):
    """
    سعی می‌کند فایلی با اسم name را در فولدر templates پیدا کند.
    پشتیبانی از پسوندهای متداول png/jpg/jpeg.
    نام جستجو را به صورت case-insensitive بررسی می‌کند.
    """
    if name is None:
        return None
    name_str = str(name).strip()
    # try exact name with common extensions
    exts = [".png", ".jpg", ".jpeg", ".bmp"]
    for ext in exts:
        candidate = os.path.join(TEMPLATES_DIR, name_str + ext)
        if os.path.exists(candidate):
            return candidate
    # try case-insensitive search
    for fname in os.listdir(TEMPLATES_DIR):
        if fname.lower().startswith(name_str.lower()):
            return os.path.join(TEMPLATES_DIR, fname)
    # fallback
    return None

def draw_dimensions_on_cell(c, img_x, img_y, img_w, img_h, WT, WB, HR, HL):
    """
    با نوشتن WT, WB, HR, HL در اطراف تصویر، ابعاد رو نمایش میده.
    مکان‌ها شبیه به کد قبلی (بالا/پایین/چپ/راست).
    """
    c.setFont(TEXT_FONT[0], TEXT_FONT[1])
    # Top (WT)
    c.drawCentredString(img_x + img_w/2, img_y + img_h + 0.2*cm, f"WT: {WT}")
    # Bottom (WB)
    c.drawCentredString(img_x + img_w/2, img_y - 0.6*cm, f"WB: {WB}")
    # Right (HR)
    c.drawString(img_x + img_w + 0.2*cm, img_y + img_h/2, f"HR: {HR}")
    # Left (HL)
    c.drawRightString(img_x - 0.2*cm, img_y + img_h/2, f"HL: {HL}")

# ---------- خواندن اکسل ----------
df = pd.read_excel(EXCEL_FILE)

# ---------- ایجاد PDF ----------
c = canvas.Canvas(OUTPUT_PDF, pagesize=A4)
page_w, page_h = A4
cell_w = page_w / COLS
cell_h = page_h / ROWS

count = 0
for idx, row in df.iterrows():
    name = str(row.get('Name Shape', '')).strip()
    WT = row.get('WT', '')
    HR = row.get('HR', '')
    WB = row.get('WB', '')
    HL = row.get('HL', '')

    # تعیین موقعیت در شبکه
    col_idx = count % COLS
    row_idx = (count // COLS) % ROWS
    cell_x = col_idx * cell_w
    cell_y = page_h - (row_idx + 1) * cell_h

    # عنوان بالای هر سلول
    c.setFont(TITLE_FONT[0], TITLE_FONT[1])
    c.drawCentredString(cell_x + cell_w/2, cell_y + cell_h - 0.5*cm, name)

    # پیدا کردن تصویر نمونه برای این مدل
    tpl = find_template_for_name(name)
    if tpl is None:
        tpl = DEFAULT_TEMPLATE if os.path.exists(DEFAULT_TEMPLATE) else None

    if tpl and os.path.exists(tpl):
        try:
            img = ImageReader(tpl)
            img_w_px, img_h_px = img.getSize()
            # محاسبه اندازه نهایی در واحد points (reportlab) با مقیاس مناسب
            max_w = (cell_w - 1.0*cm)            # حاشیه افقی داخل سلول
            max_h = (cell_h - 1.2*cm)            # حاشیه عمودی (برای عنوان هم جا بگذار)
            scale = min(max_w / img_w_px, max_h / img_h_px, (MAX_IMG_SCALE_CM*cm) / max(img_w_px, img_h_px))
            if scale <= 0:
                scale = 1.0
            draw_w = img_w_px * scale
            draw_h = img_h_px * scale
            img_x = cell_x + (cell_w - draw_w) / 2
            img_y = cell_y + (cell_h - draw_h) / 2 - 0.2*cm
            c.drawImage(img, img_x, img_y, draw_w, draw_h, preserveAspectRatio=True)
            # نوشتن ابعاد
            draw_dimensions_on_cell(c, img_x, img_y, draw_w, draw_h, WT, WB, HR, HL)
        except Exception as e:
            # خطای خواندن تصویر
            c.setFont(TEXT_FONT[0], TEXT_FONT[1])
            c.setFillColorRGB(1, 0, 0)
            c.drawCentredString(cell_x + cell_w/2, cell_y + cell_h/2, "Image error")
            c.setFillColorRGB(0, 0, 0)
    else:
        # اگر تصویر نمونه وجود نداشت
        c.setFont(TEXT_FONT[0], TEXT_FONT[1])
        c.setFillColorRGB(1, 0, 0)
        c.drawCentredString(cell_x + cell_w/2, cell_y + cell_h/2, "Template not found")
        c.setFillColorRGB(0, 0, 0)

    count += 1
    if count % IMAGES_PER_PAGE == 0:
        c.showPage()

# اگر صفحهٔ آخر پر نشده، صفحه را تمام کن
if count % IMAGES_PER_PAGE != 0:
    c.showPage()

c.save()
print("Done ->", OUTPUT_PDF)

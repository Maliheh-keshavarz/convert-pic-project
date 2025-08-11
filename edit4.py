import os
import re
import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageChops

# تنظیمات اولیه
EXCEL_FILE = "data.xlsx"
TEMPLATES_DIR = "templates"   # مسیر فولدر قالب‌ها
OUTPUT_DIR = "out_images"

OUT_W, OUT_H = 1800, 2400

FONT_PATH = r"C:\Windows\Fonts\arial.ttf"  # فونت دلخواه 
TITLE_FONT_SIZE = 100  # یا هر عددی که میخوای
DIM_FONT_SIZE = 80
THICK_FONT_SIZE = 100

MARGIN = 60
TITLE_MAX_LINES = 6
TITLE_TOP_MARGIN = 10

# ستون‌های اصلی
COL_NCODE = "code"
COL_SECTION = "Section Name"
COL_SHAPE = "Shape"
COL_SUBSHAPE = "Subshape"

COL_WT = "WT"
COL_H = "H"
COL_WB = "WB"
COL_HR = "HR"
COL_THICK = "Thickness"

# تابع حذف فضای سفید اضافی اطراف تصویر (کروپ کردن)
def trim(im, border=5):
    if im.mode != "RGBA":
        im = im.convert("RGBA")
    bg = Image.new(im.mode, im.size, (255, 255, 255, 0))
    diff = ImageChops.difference(im, bg)
    bbox = diff.getbbox()
    if bbox:
        left = max(bbox[0] - border, 0)
        upper = max(bbox[1] - border, 0)
        right = min(bbox[2] + border, im.width)
        lower = min(bbox[3] + border, im.height)
        return im.crop((left, upper, right, lower))
    else:
        return im

# تابع برای جدا کردن متن طولانی به چند خط با محدودیت خط و طول هر خط
def split_text_multiline(text, max_lines=6, max_line_len=50):
    words = text.split()
    lines = []
    current_line = ""
    for w in words:
        if len(current_line + " " + w) <= max_line_len:
            current_line = (current_line + " " + w).strip()
        else:
            lines.append(current_line)
            current_line = w
            if len(lines) == max_lines:
                break
    if current_line and len(lines) < max_lines:
        lines.append(current_line)
    return "\n".join(lines)

def sanitize_filename(s):
    s = str(s).strip()
    # فقط حذف کاراکترهای نامجاز ویندوز
    s = re.sub(r'[\\/:"*?<>|]+', '', s)
    return s or "row"

def fmt2(v):
    try:
        return f"{float(v):.2f}"
    except Exception:
        return "" if pd.isna(v) else str(v)

def find_template(shape, subshape=None):
    if not shape or str(shape).strip() == "":
        return None
    s = str(shape).strip()
    subs = str(subshape).strip() if subshape and not pd.isna(subshape) else None
    candidates = []
    if subs:
        candidates.append(f"{s}.{subs}.png")
        candidates.append(f"{s}.{subs}.jpg")
    candidates += [f"{s}.png", f"{s}.jpg"]
    for c in candidates:
        path = os.path.join(TEMPLATES_DIR, c)
        if os.path.exists(path):
            return path
    try:
        files = os.listdir(TEMPLATES_DIR)
    except FileNotFoundError:
        return None
    low_s = s.lower()
    for fname in files:
        if fname.lower().startswith(low_s):
            return os.path.join(TEMPLATES_DIR, fname)
    return None

def load_font(ttf_path, size):
    try:
        if ttf_path and os.path.exists(ttf_path):
            return ImageFont.truetype(ttf_path, size)
    except Exception:
        pass
    try:
        return ImageFont.load_default()
    except Exception:
        return None

def text_size(draw_obj, text, font):
    try:
        bbox = draw_obj.textbbox((0, 0), text, font=font)
        width = bbox[2] - bbox[0]
        height = bbox[3] - bbox[1]
        return width, height
    except AttributeError:
        return font.getsize(text)

os.makedirs(OUTPUT_DIR, exist_ok=True)

title_font = load_font(FONT_PATH, TITLE_FONT_SIZE)
dim_font = load_font(FONT_PATH, DIM_FONT_SIZE)
thick_font = load_font(FONT_PATH, THICK_FONT_SIZE)

df = pd.read_excel(EXCEL_FILE, header=1, dtype=str)

cols_check = [COL_NCODE, COL_SECTION, COL_SHAPE, COL_SUBSHAPE]

df_valid = df[~(df[cols_check].isna().all(axis=1) | df[cols_check].apply(lambda row: all(str(x).strip() == '' for x in row), axis=1))]

print(f"Rows total: {len(df)}, Rows valid: {len(df_valid)}")

for idx, row in df_valid.iterrows():
    ncode = row.get(COL_NCODE, "").strip()
    section = row.get(COL_SECTION, "").strip()
    shape = row.get(COL_SHAPE, "")
    subshape = row.get(COL_SUBSHAPE, "")
    WT = row.get(COL_WT, "")
    H = row.get(COL_H, "")
    WB = row.get(COL_WB, "")
    HR = row.get(COL_HR, "")
    TH = row.get(COL_THICK, "")

    base_name = sanitize_filename(f"{ncode} _ {section}")
    out_path = os.path.join(OUTPUT_DIR, f"{base_name}.png")

    tpl_path = find_template(shape, subshape)
    if tpl_path is None:
        default_tpl = os.path.join(TEMPLATES_DIR, "default.png")
        tpl_path = default_tpl if os.path.exists(default_tpl) else None

    canvas_img = Image.new("RGB", (OUT_W, OUT_H), color=(255, 255, 255))
    draw = ImageDraw.Draw(canvas_img)

    # رسم عنوان با چند خط و فاصله کم از بالا
    title_full_text = f"{ncode} _ {section}"
    title_text = split_text_multiline(title_full_text, max_lines=TITLE_MAX_LINES, max_line_len=50)
    lines = title_text.split('\n')
    line_heights = []
    max_line_w = 0
    for line in lines:
        w, h = text_size(draw, line, title_font)
        line_heights.append(h)
        if w > max_line_w:
            max_line_w = w
    total_h = sum(line_heights)

    y_text = TITLE_TOP_MARGIN
    for line in lines:
        w, h = text_size(draw, line, title_font)
        x_text = (OUT_W - w) // 2
        draw.text((x_text, y_text), line, fill="black", font=title_font)
        y_text += h

    content_top = y_text + 5
    content_bottom = OUT_H - MARGIN
    content_h = content_bottom - content_top
    content_w = OUT_W - 2 * MARGIN

    if tpl_path:
        try:
            tpl = Image.open(tpl_path).convert("RGBA")
            tpl_cropped = trim(tpl, border=20)

            scale = min(content_w / tpl_cropped.width, content_h / tpl_cropped.height, 1.0)
            new_w = max(1, int(tpl_cropped.width * scale))
            new_h = max(1, int(tpl_cropped.height * scale))
            tpl_resized = tpl_cropped.resize((new_w, new_h), Image.LANCZOS)

            img_x = (OUT_W - new_w) // 2
            img_y = content_top + (content_h - new_h) // 2

            canvas_img.paste(tpl_resized, (img_x, img_y), tpl_resized)
        except Exception as e:
            print(f"[{idx + 1}] Warning: couldn't open template {tpl_path}: {e}")
            draw.text((MARGIN, content_top + 10), "Template open error", fill="red", font=dim_font)
            img_x = MARGIN
            img_y = content_top
            new_w = content_w
            new_h = content_h
    else:
        nf_text = "Template not found"
        nw, nh = text_size(draw, nf_text, dim_font)
        draw.text(((OUT_W - nw) // 2, content_top + (content_h - nh) // 2), nf_text, fill="red", font=dim_font)
        img_x = MARGIN
        img_y = content_top
        new_w = content_w
        new_h = content_h

    thick_font_use = thick_font or ImageFont.load_default()
    TH_s = fmt2(TH)
    if TH_s:
        tw_th, th_th = text_size(draw, TH_s, thick_font_use)
        thick_x = img_x + (new_w - tw_th) // 2
        thick_y = img_y + (new_h - th_th) // 2
        draw.rectangle([thick_x - 6, thick_y - 4, thick_x + tw_th + 6, thick_y + th_th + 4], fill=(255, 255, 255))
        draw.text((thick_x, thick_y), TH_s, fill="black", font=thick_font_use)

    dim_font_use = dim_font or ImageFont.load_default()

    WT_s = fmt2(WT)
    if WT_s:
        wt_label = f"{WT_s}"
        wtw, wth = text_size(draw, wt_label, dim_font_use)
        wt_x = img_x + (new_w - wtw) // 2
        wt_y = max(MARGIN + 10, img_y - wth - 6)
        draw.rectangle([wt_x - 4, wt_y - 3, wt_x + wtw + 4, wt_y + wth + 3], fill=(255, 255, 255))
        draw.text((wt_x, wt_y), wt_label, fill="black", font=dim_font_use)

    WB_s = fmt2(WB)
    if WB_s:
        wb_label = f"{WB_s}"
        wbw, wbh = text_size(draw, wb_label, dim_font_use)
        wb_x = img_x + (new_w - wbw) // 2
        wb_y = min(OUT_H - MARGIN - wbh - 4, img_y + new_h + 6)
        draw.rectangle([wb_x - 4, wb_y - 3, wb_x + wbw + 4, wb_y + wbh + 3], fill=(255, 255, 255))
        draw.text((wb_x, wb_y), wb_label, fill="black", font=dim_font_use)

    HR_s = fmt2(HR)
    if HR_s:
        hr_label = f"{HR_s}"
        hrw, hrh = text_size(draw, hr_label, dim_font_use)
        hr_x = min(OUT_W - MARGIN - hrw - 4, img_x + new_w + 8)
        hr_y = img_y + (new_h - hrh) // 2
        draw.rectangle([hr_x - 4, hr_y - 3, hr_x + hrw + 4, hr_y + hrh + 3], fill=(255, 255, 255))
        draw.text((hr_x, hr_y), hr_label, fill="black", font=dim_font_use)

    H_s = fmt2(H)
    if H_s:
        h_label = f"{H_s}"
        hw, hh = text_size(draw, h_label, dim_font_use)
        h_x = max(MARGIN + 4, img_x - hw - 12)
        h_y = img_y + (new_h - hh) // 2
        draw.rectangle([h_x - 4, h_y - 3, h_x + hw + 4, h_y + hh + 3], fill=(255, 255, 255))
        draw.text((h_x, h_y), h_label, fill="black", font=dim_font_use)

    try:
        canvas_img.save(out_path, format="PNG", quality=95)
        print(f"[{idx + 1}] Saved: {out_path}")
    except Exception as e:
        print(f"[{idx + 1}] Error saving {out_path}: {e}")

print("Done.")

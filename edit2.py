import os
import re
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# =============== تنظیمات ===============
EXCEL_FILE = "data.xlsx"
TEMPLATES_DIR = "templates"
OUTPUT_DIR = "out_images"

OUT_W, OUT_H = 1200, 1600

FONT_PATH = None  # اگر فونت دلخواه داری اینجا بگذار، مثلاً "fonts/IRANSans.ttf"
TITLE_FONT_SIZE = 30
DIM_FONT_SIZE = 20
THICK_FONT_SIZE = 36

# ستون‌هایی که باید برای فیلتر چک بشوند
COL_NCODE = "code"
COL_SECTION = "Section Name"
COL_SHAPE = "Shape"
COL_SUBSHAPE = "Subshape"

COL_WT = "WT"
COL_H = "H"
COL_WB = "WB"
COL_HR = "HR"
COL_THICK = "Thickness"

MARGIN = 40
TITLE_SPACE = 60
THICKNESS_OFFSET_PERC = 0.08

# =======================================

def sanitize_filename(s):
    s = str(s).strip()
    s = re.sub(r'[\\/:"*?<>|]+', '', s)
    s = re.sub(r'\s+', '_', s)
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
        candidates.append(f"{s}.{subs}.png")  # توجه به نام فایل‌های شما که با نقطه جدا شده‌اند
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

# خواندن اکسل با ردیف دوم به عنوان header (header=1)
df = pd.read_excel(EXCEL_FILE, header=1, dtype=str)

cols_check = [COL_NCODE, COL_SECTION, COL_SHAPE, COL_SUBSHAPE]

# حذف ردیف‌هایی که تمام 4 ستون بالا خالی یا فقط فاصله‌اند
df_valid = df[~(df[cols_check].isna().all(axis=1) | df[cols_check].apply(lambda row: all(str(x).strip() == '' for x in row), axis=1))]

print(f"Rows total: {len(df)}, Rows valid: {len(df_valid)}")

for idx, row in df_valid.iterrows():
    ncode = row.get(COL_NCODE, "")
    section = row.get(COL_SECTION, "")
    shape = row.get(COL_SHAPE, "")
    subshape = row.get(COL_SUBSHAPE, "")
    WT = row.get(COL_WT, "")
    H = row.get(COL_H, "")
    WB = row.get(COL_WB, "")
    HR = row.get(COL_HR, "")
    TH = row.get(COL_THICK, "")

    base_name = sanitize_filename(f"{ncode}_{section}")
    out_path = os.path.join(OUTPUT_DIR, f"{base_name}.png")

    tpl_path = find_template(shape, subshape)
    if tpl_path is None:
        default_tpl = os.path.join(TEMPLATES_DIR, "default.png")
        tpl_path = default_tpl if os.path.exists(default_tpl) else None

    canvas_img = Image.new("RGB", (OUT_W, OUT_H), color=(255, 255, 255))
    draw = ImageDraw.Draw(canvas_img)

    # متن عنوان وسط بالا
    title_text = f"{ncode}   {section}"
    tw, th = text_size(draw, title_text, title_font)
    draw.text(((OUT_W - tw) // 2, MARGIN), title_text, fill="black", font=title_font)

    content_top = MARGIN + TITLE_SPACE
    content_bottom = OUT_H - MARGIN
    content_h = content_bottom - content_top
    content_w = OUT_W - 2 * MARGIN

    img_x = MARGIN
    img_y = content_top
    draw_w = content_w
    draw_h = content_h

    if tpl_path:
        try:
            tpl = Image.open(tpl_path).convert("RGBA")
            scale = min(content_w / tpl.width, content_h / tpl.height, 1.0)
            new_w = max(1, int(tpl.width * scale))
            new_h = max(1, int(tpl.height * scale))
            tpl_resized = tpl.resize((new_w, new_h), Image.LANCZOS)
            img_x = (OUT_W - new_w) // 2
            img_y = content_top + (content_h - new_h) // 2
            draw_w, draw_h = new_w, new_h
            canvas_img.paste(tpl_resized, (img_x, img_y), tpl_resized)
        except Exception as e:
            print(f"[{idx + 1}] Warning: couldn't open template {tpl_path}: {e}")
            draw.text((MARGIN, content_top + 10), "Template open error", fill="red", font=dim_font)
    else:
        nf_text = "Template not found"
        nw, nh = text_size(draw, nf_text, dim_font)
        draw.text(((OUT_W - nw) // 2, content_top + (content_h - nh) // 2), nf_text, fill="red", font=dim_font)

    WT_s = fmt2(WT)
    H_s = fmt2(H)
    WB_s = fmt2(WB)
    HR_s = fmt2(HR)
    TH_s = fmt2(TH)

    thick_font_use = thick_font or ImageFont.load_default()
    if TH_s:
        tw_th, th_th = text_size(draw, TH_s, thick_font_use)
        thick_x = img_x + (draw_w - tw_th) // 2
        offset = int(draw_h * THICKNESS_OFFSET_PERC)
        thick_y = img_y + draw_h - offset - th_th
        thick_y = max(content_top + 10, min(thick_y, OUT_H - MARGIN - th_th))
        draw.rectangle([thick_x - 6, thick_y - 4, thick_x + tw_th + 6, thick_y + th_th + 4], fill=(255, 255, 255))
        draw.text((thick_x, thick_y), TH_s, fill="black", font=thick_font_use)

    dim_font_use = dim_font or ImageFont.load_default()

    # بالا - WT
    if WT_s:
        wt_label = f"WT: {WT_s}"
        wtw, wth = text_size(draw, wt_label, dim_font_use)
        wt_x = img_x + (draw_w - wtw) // 2
        wt_y = max(MARGIN + 10, img_y - wth - 6)
        draw.rectangle([wt_x - 4, wt_y - 3, wt_x + wtw + 4, wt_y + wth + 3], fill=(255, 255, 255))
        draw.text((wt_x, wt_y), wt_label, fill="black", font=dim_font_use)

    # پایین - WB
    if WB_s:
        wb_label = f"WB: {WB_s}"
        wbw, wbh = text_size(draw, wb_label, dim_font_use)
        wb_x = img_x + (draw_w - wbw) // 2
        wb_y = min(OUT_H - MARGIN - wbh - 4, img_y + draw_h + 6)
        draw.rectangle([wb_x - 4, wb_y - 3, wb_x + wbw + 4, wb_y + wbh + 3], fill=(255, 255, 255))
        draw.text((wb_x, wb_y), wb_label, fill="black", font=dim_font_use)

    # راست - HR
    if HR_s:
        hr_label = f"HR: {HR_s}"
        hrw, hrh = text_size(draw, hr_label, dim_font_use)
        hr_x = min(OUT_W - MARGIN - hrw - 4, img_x + draw_w + 8)
        hr_y = img_y + (draw_h - hrh) // 2
        draw.rectangle([hr_x - 4, hr_y - 3, hr_x + hrw + 4, hr_y + hrh + 3], fill=(255, 255, 255))
        draw.text((hr_x, hr_y), hr_label, fill="black", font=dim_font_use)

    # چپ - H
    if H_s:
        h_label = f"H: {H_s}"
        hw, hh = text_size(draw, h_label, dim_font_use)
        h_x = max(MARGIN + 4, img_x - hw - 12)
        h_y = img_y + (draw_h - hh) // 2
        draw.rectangle([h_x - 4, h_y - 3, h_x + hw + 4, h_y + hh + 3], fill=(255, 255, 255))
        draw.text((h_x, h_y), h_label, fill="black", font=dim_font_use)

    try:
        canvas_img.save(out_path, format="PNG")
        print(f"[{idx + 1}] Saved: {out_path}")
    except Exception as e:
        print(f"[{idx + 1}] Error saving {out_path}: {e}")

print("Done.")

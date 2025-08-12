import os
import re
import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageChops
import tkinter as tk
from tkinter import filedialog

# تنظیمات اولیه
EXCEL_FILE = "data.xlsx"
TEMPLATES_DIR = "templates"

# اندازه کل تصویر خروجی (canvas)
OUT_W, OUT_H = 1400, 1800

# فونت‌ها و اندازه‌ها
FONT_PATH = r"C:\Windows\Fonts\arial.ttf"
TITLE_FONT_SIZE = 50
DIM_FONT_SIZE = 80
THICK_FONT_SIZE = 80

# حاشیه سفید اطراف قالب
MARGIN = 20

# رنگ‌ها
COLOR_TEXT = (0, 0, 0)
COLOR_RECT_BG = (255, 255, 255)

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

def select_output_dir(default="out_images"):
    root = tk.Tk()
    root.withdraw()
    folder_selected = filedialog.askdirectory(title="Select base output directory")
    if not folder_selected:
        folder_selected = os.getcwd()
    root.destroy()
    final_output_dir = os.path.join(folder_selected, default)
    os.makedirs(final_output_dir, exist_ok=True)
    return final_output_dir

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

def replace_gray_with_white(img, threshold=200):
    pixels = img.load()
    w, h = img.size
    for x in range(w):
        for y in range(h):
            r, g, b, a = pixels[x, y]
            if a > 0:
                if abs(r - g) < 10 and abs(g - b) < 10 and r > threshold and g > threshold and b > threshold:
                    pixels[x, y] = (255, 255, 255, a)
    return img

def sanitize_filename(s):
    s = str(s).strip()
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

def resize_template_with_margin(img, out_w, out_h, margin):
    max_w = out_w - 2 * margin
    max_h = out_h - 2 * margin
    w, h = img.size

    scale = min(max_w / w, max_h / h, 1.0)
    new_w = int(w * scale)
    new_h = int(h * scale)
    resized_img = img.resize((new_w, new_h), Image.LANCZOS)
    return resized_img

def main():
    output_dir = select_output_dir()

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
        out_path = os.path.join(output_dir, f"{base_name}.png")

        tpl_path = find_template(shape, subshape)
        if tpl_path is None:
            default_tpl = os.path.join(TEMPLATES_DIR, "default.png")
            tpl_path = default_tpl if os.path.exists(default_tpl) else None

        canvas_img = Image.new("RGB", (OUT_W, OUT_H), color=(255, 255, 255))
        draw = ImageDraw.Draw(canvas_img)

        # رسم عنوان بالای قالب وسط
        title_line = section
        w, h = text_size(draw, title_line, title_font)
        x = (OUT_W - w) // 2
        y = 50  # فاصله از بالا
        draw.text((x, y), title_line, fill=COLOR_TEXT, font=title_font)

        if tpl_path:
            try:
                tpl = Image.open(tpl_path).convert("RGBA")
                tpl = replace_gray_with_white(tpl)
                tpl_cropped = trim(tpl, border=20)
                tpl_resized = resize_template_with_margin(tpl_cropped, OUT_W, OUT_H, MARGIN)

                img_x = (OUT_W - tpl_resized.width) // 2
                img_y = (OUT_H - tpl_resized.height) // 2

                canvas_img.paste(tpl_resized, (img_x, img_y), tpl_resized)
            except Exception as e:
                print(f"[{idx + 1}] Warning: couldn't open template {tpl_path}: {e}")
                tpl_resized = None
                img_x = MARGIN
                img_y = MARGIN
        else:
            tpl_resized = None
            img_x = MARGIN
            img_y = MARGIN

        # ضخامت وسط قالب
        TH_s = fmt2(TH)
        if TH_s:
            thick_label = f"Th: {TH_s}"
            tw_th, th_th = text_size(draw, thick_label, thick_font)
            thick_x = OUT_W // 2 - tw_th // 2
            thick_y = img_y + (tpl_resized.height // 2 if tpl_resized else OUT_H // 2) - th_th // 2
            draw.rectangle([thick_x - 6, thick_y - 4, thick_x + tw_th + 6, thick_y + th_th + 4], fill=COLOR_RECT_BG)
            draw.text((thick_x, thick_y), thick_label, fill=COLOR_TEXT, font=thick_font)

        # موقعیت های WT و H بر اساس Shape
        if str(shape).strip().lower() == "step beam":
            # موقعیت WT و H متناسب با اندازه قالب (در قسمت Step)
            wt_x = img_x + int(tpl_resized.width * 0.5)
            wt_y = img_y + int(tpl_resized.height * 0.25)

            h_x = img_x + int(tpl_resized.width * 0.5)
            h_y = img_y + int(tpl_resized.height * 0.75)
        else:
            # موقعیت ثابت برای Box Beam و بقیه
            wt_x = OUT_W // 2
            wt_y = 300

            h_x = 200
            h_y = OUT_H // 2

        # موقعیت ثابت WB و HR
        wb_x = OUT_W // 2
        wb_y = OUT_H - 190

        hr_x = OUT_W - 200
        hr_y = OUT_H // 2

        # تابع کمکی رسم متن با پس‌زمینه سفید
        def draw_text_centered(label, cx, cy, font):
            w, h = text_size(draw, label, font)
            x = cx - w // 2
            y = cy - h // 2
            draw.rectangle([x - 4, y - 3, x + w + 4, y + h + 3], fill=COLOR_RECT_BG)
            draw.text((x, y), label, fill=COLOR_TEXT, font=font)

        WT_s = fmt2(WT)
        if WT_s:
            draw_text_centered(WT_s, wt_x, wt_y, dim_font)

        H_s = fmt2(H)
        if H_s:
            draw_text_centered(H_s, h_x, h_y, dim_font)

        WB_s = fmt2(WB)
        if WB_s:
            draw_text_centered(WB_s, wb_x, wb_y, dim_font)

        HR_s = fmt2(HR)
        if HR_s:
            draw_text_centered(HR_s, hr_x, hr_y, dim_font)

        try:
            canvas_img.save(out_path, format="PNG", quality=95)
            print(f"[{idx + 1}] Saved: {out_path}")
        except Exception as e:
            print(f"[{idx + 1}] Error saving {out_path}: {e}")

    print("Done.")

if __name__ == "__main__":
    main()

import os
import re
import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageChops
import tkinter as tk
from tkinter import filedialog

# تنظیمات اولیه
EXCEL_FILE = "data.xlsx"
TEMPLATES_DIR = "templates"

OUT_W, OUT_H = 1800, 2400

FONT_PATH = r"C:\Windows\Fonts\arial.ttf"
TITLE_FONT_SIZE = 60
DIM_FONT_SIZE = 80
THICK_FONT_SIZE = 100

MARGIN = 60
TITLE_TOP_MARGIN = 100

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

# رنگ‌ها
COLOR_TEXT = (0, 0, 0)
COLOR_RECT_BG = (255, 255, 255)

# جای مقادیر ابعاد (اعداد، می‌تونی تغییر بدی)
positions = {
    COL_WT: ("center", "top"),
    COL_WB: ("center", "bottom"),
    COL_HR: ("right", "center"),
    COL_H: ("left", "center"),
}

def select_output_dir(default="out_images"):
    root = tk.Tk()
    root.withdraw()
    folder_selected = filedialog.askdirectory(title="Select base output directory")
    if not folder_selected:
        folder_selected = os.getcwd()  # اگر انتخاب نکردی، فولدر جاری رو بگیر
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

def get_pos(base_x, base_y, base_w, base_h, text_w, text_h, pos):
    x, y = base_x, base_y
    if pos[0] == "center":
        x = base_x + (base_w - text_w) // 2
    elif pos[0] == "left":
        x = base_x - text_w - 12
    elif pos[0] == "right":
        x = base_x + base_w + 8

    if pos[1] == "center":
        y = base_y + (base_h - text_h) // 2
    elif pos[1] == "top":
        y = max(MARGIN + 10, base_y - text_h - 6)
    elif pos[1] == "bottom":
        y = min(OUT_H - MARGIN - text_h - 4, base_y + base_h + 6)
    return x, y

def draw_dims(draw, dim_font, img_x, img_y, new_w, new_h, dims):
    for key, val in dims.items():
        val_str = fmt2(val)
        if not val_str:
            continue
        if key == COL_THICK:
            label = f"Th: {val_str}"
            font = load_font(FONT_PATH, THICK_FONT_SIZE)
            tw, th = text_size(draw, label, font)
            x = img_x + (new_w - tw) // 2
            y = img_y + (new_h - th) // 2
            draw.rectangle([x - 6, y - 4, x + tw + 6, y + th + 4], fill=COLOR_RECT_BG)
            draw.text((x, y), label, fill=COLOR_TEXT, font=font)
        else:
            label = val_str
            font = dim_font
            tw, th = text_size(draw, label, font)
            pos = positions.get(key, ("center", "top"))
            x, y = get_pos(img_x, img_y, new_w, new_h, tw, th, pos)
            draw.rectangle([x - 6, y - 4, x + tw + 6, y + th + 4], fill=COLOR_RECT_BG)
            draw.text((x, y), label, fill=COLOR_TEXT, font=font)

def main():
    output_dir = select_output_dir()

    title_font = load_font(FONT_PATH, TITLE_FONT_SIZE)
    dim_font = load_font(FONT_PATH, DIM_FONT_SIZE)

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

        title_line = section
        w, h = text_size(draw, title_line, title_font)
        x = (OUT_W - w) // 2
        y = TITLE_TOP_MARGIN
        draw.text((x, y), title_line, fill=COLOR_TEXT, font=title_font)

        content_top = y + h + 10
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

        dims = {
            COL_WT: WT,
            COL_WB: WB,
            COL_HR: HR,
            COL_H: H,
            COL_THICK: TH,
        }

        draw_dims(draw, dim_font, img_x, img_y, new_w, new_h, dims)

        try:
            canvas_img.save(out_path, format="PNG", quality=95)
            print(f"[{idx + 1}] Saved: {out_path}")
        except Exception as e:
            print(f"[{idx + 1}] Error saving {out_path}: {e}")

    print("Done.")

if __name__ == "__main__":
    main()

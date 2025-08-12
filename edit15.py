import os
import re
import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageChops
import tkinter as tk
from tkinter import filedialog

EXCEL_FILE = "data.xlsx"
TEMPLATES_DIR = "templates"

OUT_W, OUT_H = 1600, 2200

FONT_PATH = r"C:\Windows\Fonts\arial.ttf"
TITLE_FONT_SIZE = 60
DIM_FONT_SIZE = 80
THICK_FONT_SIZE = 100

MARGIN = -50
TITLE_TOP_MARGIN = 100

COL_NCODE = "code"
COL_SECTION = "Section Name"
COL_SHAPE = "Shape"
COL_SUBSHAPE = "Subshape"

COL_WT = "WT"
COL_H = "H"
COL_WB = "WB"
COL_HR = "HR"
COL_THICK = "Thickness"

COLOR_TEXT = (0, 0, 0)
COLOR_RECT_BG = (255, 255, 255)

wt_pos = (700, 300)
wb_pos = (700, 2200)
hr_pos = (1000, 1300)
h_pos  = (200, 1300)

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

def main():
    output_dir = select_output_dir()

    title_font = load_font(FONT_PATH, TITLE_FONT_SIZE)
    dim_font = load_font(FONT_PATH, DIM_FONT_SIZE)
    thick_font = load_font(FONT_PATH, THICK_FONT_SIZE)

    df = pd.read_excel(EXCEL_FILE, header=1, dtype=str)

    cols_check = [COL_NCODE, COL_SECTION, COL_SHAPE, COL_SUBSHAPE]
    df_valid = df[~(df[cols_check].isna().all(axis=1) | df[cols_check].apply(lambda row: all(str(x).strip() == '' for x in row), axis=1))]

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
                tpl = replace_gray_with_white(tpl)
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
                img_x = MARGIN
                img_y = content_top
                new_w = content_w
                new_h = content_h
        else:
            img_x = MARGIN
            img_y = content_top
            new_w = content_w
            new_h = content_h

        # ضخامت
        TH_s = fmt2(TH)
        if TH_s:
            thick_label = f"Th: {TH_s}"
            tw_th, th_th = text_size(draw, thick_label, thick_font)
            thick_x = img_x + (new_w - tw_th) // 2
            thick_y = img_y + (new_h - th_th) // 2
            draw.rectangle([thick_x - 6, thick_y - 4, thick_x + tw_th + 6, thick_y + th_th + 4], fill=COLOR_RECT_BG)
            draw.text((thick_x, thick_y), thick_label, fill=COLOR_TEXT, font=thick_font)

        # موقعیت WT و HR برای Step Beam داینامیک محاسبه شود
        if shape.strip().lower() == "step beam":
            try:
                tpl_width, tpl_height = tpl_resized.size
            except NameError:
                tpl_width, tpl_height = 1000, 1000

            wt_x = img_x + tpl_width // 2
            wt_y = img_y + 300  # این عدد قابل تغییر است

            hr_x = img_x + tpl_width - 200  # این عدد قابل تغییر است
            hr_y = img_y + tpl_height // 2
        else:
            wt_x, wt_y = wt_pos
            hr_x, hr_y = hr_pos

        # WT
        WT_s = fmt2(WT)
        if WT_s:
            wt_label = f"{WT_s}"
            wtw, wth = text_size(draw, wt_label, dim_font)
            wt_draw_x = wt_x - wtw // 2
            wt_draw_y = wt_y - wth // 2
            draw.rectangle([wt_draw_x - 4, wt_draw_y - 3, wt_draw_x + wtw + 4, wt_draw_y + wth + 3], fill=COLOR_RECT_BG)
            draw.text((wt_draw_x, wt_draw_y), wt_label, fill=COLOR_TEXT, font=dim_font)

        # WB (همیشه وسط است)
        WB_s = fmt2(WB)
        if WB_s:
            wb_label = f"{WB_s}"
            wbw, wbh = text_size(draw, wb_label, dim_font)
            wb_x = wb_pos[0] - wbw // 2
            wb_y = wb_pos[1] - wbh // 2
            draw.rectangle([wb_x - 4, wb_y - 3, wb_x + wbw + 4, wb_y + wbh + 3], fill=COLOR_RECT_BG)
            draw.text((wb_x, wb_y), wb_label, fill=COLOR_TEXT, font=dim_font)

        # HR
        HR_s = fmt2(HR)
        if HR_s:
            hr_label = f"{HR_s}"
            hrw, hrh = text_size(draw, hr_label, dim_font)
            hr_draw_x = hr_x - hrw // 2
            hr_draw_y = hr_y - hrh // 2
            draw.rectangle([hr_draw_x - 4, hr_draw_y - 3, hr_draw_x + hrw + 4, hr_draw_y + hrh + 3], fill=COLOR_RECT_BG)
            draw.text((hr_draw_x, hr_draw_y), hr_label, fill=COLOR_TEXT, font=dim_font)

        # H (همیشه وسط است)
        H_s = fmt2(H)
        if H_s:
            h_label = f"{H_s}"
            hw, hh = text_size(draw, h_label, dim_font)
            h_x = h_pos[0] - hw // 2
            h_y = h_pos[1] - hh // 2
            draw.rectangle([h_x - 4, h_y - 3, h_x + hw + 4, h_y + hh + 3], fill=COLOR_RECT_BG)
            draw.text((h_x, h_y), h_label, fill=COLOR_TEXT, font=dim_font)

        try:
            canvas_img.save(out_path, format="PNG", quality=95)
            print(f"[{idx + 1}] Saved: {out_path}")
        except Exception as e:
            print(f"[{idx + 1}] Error saving {out_path}: {e}")

    print("Done.")

if __name__ == "__main__":
    main()

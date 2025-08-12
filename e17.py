import ezdxf
from PIL import Image
import os

# تابع حذف پس‌زمینه خاکستری
def replace_gray_with_white(img, threshold=200):
    pixels = img.load()
    w, h = img.size
    for x in range(w):
        for y in range(h):
            r, g, b, a = pixels[x, y]
            if a > 0:
                if abs(r - g) < 10 and abs(g - b) < 10 and r > threshold and g > threshold and b > threshold:
                    pixels[x, y] = (255, 255, 255, 0)
    return img

# خواندن ابعاد واقعی از فایل DXF
def get_dimensions_from_dxf(file_path):
    doc = ezdxf.readfile(file_path)
    msp = doc.modelspace()
    min_x, min_y, max_x, max_y = None, None, None, None

    for e in msp.query("LINE CIRCLE ARC LWPOLYLINE"):
        for p in e.vertices() if hasattr(e, "vertices") else []:
            x, y = p[0], p[1]
            if min_x is None or x < min_x: min_x = x
            if min_y is None or y < min_y: min_y = y
            if max_x is None or x > max_x: max_x = x
            if max_y is None or y > max_y: max_y = y

    width = max_x - min_x
    height = max_y - min_y
    return width, height

# پردازش داده‌ها
def process_beam(shape, subshape, template_path, output_folder):
    width, height = get_dimensions_from_dxf(template_path)

    if shape.lower() == "box beam":
        WT, WB, HR, H = width, width, height, height
    elif shape.lower() == "step beam":
        WB, H = width, height
        WT = width / 2  # نمونه: وسط
        HR = height / 2
    else:
        return

    # تبدیل CAD به تصویر (نیاز به ابزار جانبی)
    # اینجا فرض می‌کنیم قبلاً به PNG تبدیل شده
    cad_png = template_path.replace(".dxf", ".png")
    if not os.path.exists(cad_png):
        print(f"⚠ فایل PNG از CAD یافت نشد: {cad_png}")
        return

    img = Image.open(cad_png).convert("RGBA")
    img = replace_gray_with_white(img)

    output_path = os.path.join(output_folder, f"{shape}_{subshape}.png")
    img.save(output_path)
    print(f"✅ ذخیره شد: {output_path}")

# تست روی 5 ردیف اول
data = [
    ("Box Beam", "Sub1", r"C:\templates\box1.dxf"),
    ("Step Beam", "Sub2", r"C:\templates\step1.dxf"),
    ("Box Beam", "Sub3", r"C:\templates\box2.dxf"),
    ("Step Beam", "Sub4", r"C:\templates\step2.dxf"),
    ("Box Beam", "Sub5", r"C:\templates\box3.dxf"),
]

for shape, subshape, path in data:
    process_beam(shape, subshape, path, r"C:\output")

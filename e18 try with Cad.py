import pandas as pd
import ezdxf
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
import matplotlib.pyplot as plt
import os

# خواندن دیتابیس
df = pd.read_excel("data.xlsx", header=1)
df = df.dropna(subset=["Shape", "Subshape", "WT", "H", "WB", "HR", "Thickness"])

# ایجاد پوشه خروجی
os.makedirs("output", exist_ok=True)

for _, row in df.iterrows():
    section_name = str(row["Section Name"]).strip()
    shape = str(row["Shape"]).strip()
    subshape = str(row["Subshape"]).strip()

    dxf_path = f"templates/{shape}.{subshape}.dxf"
    if not os.path.exists(dxf_path):
        print(f"DXF not found: {dxf_path}")
        continue

    WT, H, WB, HR, TH = row["WT"], row["H"], row["WB"], row["HR"], row["Thickness"]

    # باز کردن فایل DXF
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    # محاسبه bounding box و جمع آوری تمام نقاط
    min_x = min_y = float('inf')
    max_x = max_y = float('-inf')
    points_x = []
    points_y = []

    for e in msp:
        try:
            if e.dxftype() in ('LINE', 'LWPOLYLINE', 'POLYLINE'):
                pts = []
                if e.dxftype() == 'LINE':
                    pts = [e.dxf.start, e.dxf.end]
                else:
                    pts = [p[:2] for p in e.get_points()]
                for x, y in pts:
                    min_x = min(min_x, x)
                    max_x = max(max_x, x)
                    min_y = min(min_y, y)
                    max_y = max(max_y, y)
                    points_x.append(x)
                    points_y.append(y)
        except Exception:
            pass

    # اگر bounding box معتبر نیست، از مقادیر پیشفرض استفاده کنیم
    if min_x == float('inf') or max_x == float('-inf'):
        min_x, max_x = 0, WT
    if min_y == float('inf') or max_y == float('-inf'):
        min_y, max_y = 0, HR

    width = max_x - min_x
    height = max_y - min_y
    if width == 0: width = 1
    if height == 0: height = 1

    # مرکز واقعی شکل به صورت مستقل برای هر محور
    center_x = sum(points_x)/len(points_x) if points_x else (min_x + max_x)/2
    center_y = sum(points_y)/len(points_y) if points_y else (min_y + max_y)/2

    # رندر با Matplotlib
    fig, ax = plt.subplots(figsize=(6, 6))
    ax.set_aspect("equal")
    ax.axis("off")

    ctx = RenderContext(doc)
    frontend = Frontend(ctx, MatplotlibBackend(ax))
    frontend.draw_layout(msp)

    # تنظیم محدوده محور با حاشیه
    margin_x = width * 0.1
    margin_y = height * 0.1
    ax.set_xlim(min_x - margin_x, max_x + margin_x)
    ax.set_ylim(min_y - margin_y, max_y + margin_y)

    # افزودن متن‌ها نسبت به bounding box و مرکز واقعی
    fontsize = min(width, height) * 3  # اندازه متن نسبی به شکل
    ax.text(min_x + 0.05*width, min_y + 0.5*height, f"HR: {HR:.2f}", ha='left', va='center', fontsize=fontsize)
    ax.text(center_x, max_y - 0.05*height, f"WT: {WT:.2f}", ha='center', va='top', fontsize=fontsize)
    ax.text(max_x - 0.05*width, center_y, f"H: {H:.2f}", ha='right', va='center', fontsize=fontsize)
    ax.text(min_x + 0.5*width, min_y + 0.05*height, f"WB: {WB:.2f}", ha='center', va='bottom', fontsize=fontsize)
    ax.text(center_x, center_y, f"Th: {TH:.2f}", ha='center', va='center', fontsize=fontsize)
    ax.text(min_x + 0.05*width, max_y + 0.05*height, section_name, ha='left', va='bottom', fontsize=fontsize*1.2)

    # ذخیره PNG
    plt.savefig(f"output/{section_name}.png", dpi=300, bbox_inches="tight")
    plt.close(fig)

print("Matplotlib rendering with shape and accurately centered WT/H finished safely.")

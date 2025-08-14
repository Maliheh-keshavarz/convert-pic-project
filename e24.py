import pandas as pd
import ezdxf
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
import matplotlib.pyplot as plt
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# ---------- پنجره انتخاب فایل Excel ----------
Tk().withdraw()  # پنهان کردن پنجره اصلی Tkinter
excel_path = askopenfilename(title="Select Excel Database", filetypes=[("Excel files", "*.xlsx *.xls")])
if not excel_path:
    print("No Excel file selected. Exiting.")
    exit()

## ---------- خواندن دیتابیس ----------
df = pd.read_excel(excel_path, header=1)

## اگر ستون‌های xl و yb موجود نباشند، مقدار پیش‌فرض صفر بده
if "xl  =" not in df.columns:
    df["xl  ="] = 0
if "yb  =" not in df.columns:
    df["yb  ="] = 0

# ستون مسیر فایل SCT (ستون تک-مسیر)؛ اسم ستون رو با اسم واقعی خودت هماهنگ کن
# مثلا "File Address" یا هر چیزی که گذاشتی
FILE_COL = "File Address"

## حذف ردیف‌های ناقص
req_cols = ["Shape", "Subshape", "WT", "H", "WB", "HR", "Thickness", FILE_COL]
df = df.dropna(subset=[c for c in req_cols if c in df.columns])

for _, row in df.iterrows():
    section_name = str(row.get("Section Name", "")).strip()
    shape = str(row["Shape"]).strip()
    subshape = str(row["Subshape"]).strip()

    ## ---------- ساخت مسیر DXF و PNG از روی مسیر SCT ----------
    ## مسیر SCT رو از دیتابیس بخون
    sct_path_raw = str(row[FILE_COL]).strip()

    ## نرمال‌سازی مسیر (برای UNC هم جواب می‌ده)
    sct_path = os.path.normpath(sct_path_raw)

    ## پایه‌ی نام فایل بدون پسوند
    base_path, _ = os.path.splitext(sct_path)

    ## ورودی DXF و خروجی PNG کنار همان فایل
    dxf_path = base_path + ".dxf"
    png_path = base_path + ".png"

    ## اگر DXF موجود نیست، رد شو
    if not os.path.exists(dxf_path):
        print(f"DXF file not found: {dxf_path}")
        continue

    ## ---------- خواندن مقادیر از دیتابیس ----------
    WT = float(row["WT"])
    H  = float(row["H"])
    WB = float(row["WB"])
    HR = float(row["HR"])
    TH = float(row["Thickness"])
    XL = float(row["xl  ="])
    YB = float(row["yb  ="])
    WO = float(row.get("Brace Entering", 0))  # اگر نبود، صفر

    ## ---------- باز کردن DXF ----------
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    ## ---------- محاسبه مختصات متن‌ها (منطق فعلی مورد تایید شما) ----------
    ## H
    H_x = -XL - (2 * TH)
    H_y = (H / 2) - YB

    ## WB
    WB_x = -XL + (WB / 2)
    WB_y = -YB - (2 * TH)

    ## HR
    HR_x = (-XL + WB) + (2 * TH)
    HR_y = -YB + (HR / 2)

    ## TH (مرکز ارتفاع HR)
    TH_x = 0
    TH_y = -YB + (HR / 2)

    ## WT
    WT_x = -XL + (WT / 2)
    WT_y = (H - YB) + (2 * TH)

    ## Subshape
    Subshape_x = 0
    Subshape_y = TH_y + 0.2

    ## Title
    SectionName_x = -XL + (WT / 2)
    SectionName_y = (H / 2) + 0.4

    ## ---------- چاپ مختصات برای بررسی ----------
    # print()
    # print(f"Section: {section_name or os.path.basename(base_path)}")
    # print(f"DXF: {dxf_path}")
    # print(f"PNG: {png_path}")
    # print(f"H:  ({H_x:.3f}, {H_y:.3f})")
    # print(f"HR: ({HR_x:.3f}, {HR_y:.3f}) -> value={HR:.2f}")
    # print(f"WT: ({WT_x:.3f}, {WT_y:.3f}) -> value={WT:.2f}")
    # print(f"Th: ({TH_x:.3f}, {TH_y:.3f}) -> value={TH:.2f}")
    # # اگر shape از نوع Brace/Post باشد، WO چاپ می‌شود
    # if shape in ["Brace", "Post"]:
    #     print(f"WO: ({WB_x:.3f}, {WB_y:.3f}) -> value={WO:.2f}")
    # else:
    #     print(f"WB: ({WB_x:.3f}, {WB_y:.3f}) -> value={WB:.2f}")

    ## ---------- رندر ----------
    fig, ax = plt.subplots(figsize=(6, 6))
    ax.set_aspect("equal")
    ax.axis("off")

    frontend = Frontend(RenderContext(doc), MatplotlibBackend(ax))
    frontend.draw_layout(msp)

    ## اضافه کردن متن‌ها روی شکل
    ## H
    ax.text(H_x, H_y, f"H: {H:.2f}", ha='right', va='center', fontsize=12, color='red')
    ## WB یا WO
    if shape in ["Brace", "Post"]:
        ax.text(WB_x, WB_y, f"WO: {WO:.2f}", ha='center', va='center', fontsize=12, color='blue')
    else:
        ax.text(WB_x, WB_y, f"WB: {WB:.2f}", ha='center', va='center', fontsize=12, color='blue')
    ## HR
    ax.text(HR_x, HR_y, f"HR: {HR:.2f}", ha='left', va='center', fontsize=12, color='green')
    ## WT
    ax.text(WT_x, WT_y, f"WT: {WT:.2f}", ha='center', va='center', fontsize=12, color='orange')
    ## Th
    ax.text(TH_x, TH_y, f"Th: {TH:.2f}", ha='center', va='center', fontsize=14, color='purple')
    ## Title + Subshape
    ax.text(SectionName_x, SectionName_y, f"\n{section_name}", ha='center', va='bottom', fontsize=14, color='black')
    ax.text(Subshape_x, Subshape_y, subshape, ha='center', va='center', fontsize=30, color='black')

    ## اطمینان از وجود فولدر مقصد (اگر نباشد می‌سازد)
    os.makedirs(os.path.dirname(png_path), exist_ok=True)

    ## ذخیره PNG کنار همان DXF/SCT
    plt.savefig(png_path, dpi=300, bbox_inches="tight")
    plt.close(fig)

print("\nAll shapes rendered with H, WB/WO, HR, WT, Th, and Section Name correctly positioned, using single-path column for DXF/PNG.")

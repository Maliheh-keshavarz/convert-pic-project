## ========================================================================================================
## SUMMARY / خلاصه
## 
## This script reads structural data from an Excel database, opens corresponding
## DXF files, calculates label positions (H, WB/WO, HR, WT, TH, Subshape, Section Name)
## with respect to XL and YB shifts, renders the DXF using matplotlib, adds text
## annotations on the shapes, and saves the annotated images as PNG files
## alongside the DXF.
##
##
## INPUT / ورودی:
## - Excel file with columns: Shape, Subshape, WT, H, WB, HR, Thickness, File Address, xl  =, yb  =
## - DXF files located at the paths specified in the Excel 'File Address' column
##
## OUTPUT / خروجی:
## - PNG images of each shape with annotated labels, saved alongside DXF files
##
## PROCESS / فرآیند اصلی:
## 1. Read Excel database and set default values for missing XL/YB
## 2. Iterate through each row, construct DXF and PNG paths
## 3. Read dimensions from database (H, WT, WB, HR, TH, XL, YB)
## 4. Open DXF, calculate text label coordinates
## 5. Render DXF and overlay text labels
## 6. Ensure output folder exists and save PNG
## ========================================================================================================

import pandas as pd
import ezdxf
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
import matplotlib.pyplot as plt
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename

## ---------- Excel File Selection / انتخاب فایل Excel ----------
Tk().withdraw()  ## Hide the main Tkinter window / پنهان کردن پنجره اصلی Tkinter
excel_path = askopenfilename(
    title="Select Excel Database", 
    filetypes=[("Excel files", "*.xlsx *.xls")]
)
if not excel_path:
    print("No Excel file selected. Exiting.")  ## If no file chosen, exit / اگر فایلی انتخاب نشد، خروج
    exit()

## ---------- Read Database / خواندن دیتابیس ----------
df = pd.read_excel(excel_path, header=1)

## ---------- Set default values if columns missing / اگر ستون‌های xl و yb موجود نبود، صفر بده ----------
if "xl  =" not in df.columns:
    df["xl  ="] = 0
if "yb  =" not in df.columns:
    df["yb  ="] = 0

## ---------- File path column / ستون مسیر فایل ----------
FILE_COL = "File Address"  ## Adjust to your column name / اسم ستون خودت رو وارد کن

## ---------- Remove incomplete rows / حذف ردیف‌های ناقص ----------
req_cols = ["Shape", "Subshape", "WT", "H", "WB", "HR", "Thickness", FILE_COL]
df = df.dropna(subset=[c for c in req_cols if c in df.columns])

## ---------- Iterate through each row / پردازش هر ردیف ----------
for _, row in df.iterrows():
    section_name = str(row.get("Section Name", "")).strip()
    shape = str(row["Shape"]).strip()
    subshape = str(row["Subshape"]).strip()

    ## ---------- Construct DXF and PNG paths / ساخت مسیر DXF و PNG ----------
    sct_path_raw = str(row[FILE_COL]).strip()
    sct_path = os.path.normpath(sct_path_raw)  ## Normalize path / نرمال‌سازی مسیر
    base_path, _ = os.path.splitext(sct_path)
    dxf_path = base_path + ".dxf"  ## DXF file path / مسیر DXF
    png_path = base_path + ".png"  ## PNG output path / مسیر خروجی PNG

    if not os.path.exists(dxf_path):
        print(f"DXF file not found: {dxf_path}")  ## Skip if DXF missing / رد کردن فایل اگر DXF موجود نباشد
        continue

    ## ---------- Read dimensions from database / خواندن ابعاد از دیتابیس ----------
    WT = float(row["WT"])         ## Width of the top 
    H  = float(row["H"])          ## height Right 
    WB = float(row["WB"])         ## width Bottom
    HR = float(row["HR"])         ## Height Left
    TH = float(row["Thickness"])  ## Thickness 
    XL = float(row["xl  ="])      ## Horizontal shift 
    YB = float(row["yb  ="])      ## Vertical shift 
    WO = float(row.get("Brace Entering", 0))  ## Brace width (if shape is Brace/Post) / عرض مهاربند یا ستون، صفر اگر موجود نباشد

    ## ---------- Open DXF / باز کردن DXF ----------
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    ## ---------- Calculate text coordinates / محاسبه مختصات متن ----------

    ## For all labels :
    ## X coordinate = base horizontal position, shifted left or centered according to XL and spacing
    ## Y coordinate = base vertical position, adjusted relative to reference height and YB offset
    
    ## H label 
    H_x = -XL - (2 * TH)
    H_y = (H / 2) - YB

    ## WB label 
    WB_x = -XL + (WB / 2)
    WB_y = -YB - (2 * TH)

    ## HR label 
    HR_x = (-XL + WB) + (2 * TH)
    HR_y = -YB + (HR / 2)

    ## TH label 
    TH_x = 0
    TH_y = -YB + (HR / 2)

    ## WT label
    WT_x = -XL + (WT / 2)
    WT_y = (H - YB) + (2 * TH)

    ## Subshape label 
    Subshape_x = 0
    Subshape_y = TH_y + 0.2

    ## Section Name 
    SectionName_x = -XL + (WT / 2)
    SectionName_y = (H / 2) + 0.4

    ## ---------- Print coordinates for review / چاپ مختصات برای بررسی ----------
    # print()
    # print(f"Section: {section_name or os.path.basename(base_path)}")
    # print(f"DXF: {dxf_path}")
    # print(f"PNG: {png_path}")
    # print(f"H:  ({H_x:.3f}, {H_y:.3f})")
    # print(f"HR: ({HR_x:.3f}, {HR_y:.3f}) -> value={HR:.2f}")
    # print(f"WT: ({WT_x:.3f}, {WT_y:.3f}) -> value={WT:.2f}")
    # print(f"Th: ({TH_x:.3f}, {TH_y:.3f}) -> value={TH:.2f}")
    # if shape in ["Brace", "Post"]:
    #     print(f"WO: ({WB_x:.3f}, {WB_y:.3f}) -> value={WO:.2f}")
    # else:
    #     print(f"WB: ({WB_x:.3f}, {WB_y:.3f}) -> value={WB:.2f}")

    ## ---------- Rendering / رندر ----------
    fig, ax = plt.subplots(figsize=(6, 6))
    ax.set_aspect("equal")
    ax.axis("off")
    frontend = Frontend(RenderContext(doc), MatplotlibBackend(ax))
    frontend.draw_layout(msp)

    ## ---------- Add text labels / اضافه کردن متن‌ها ----------
    ax.text(H_x, H_y, f"HL: {H:.2f}", ha='right', va='center', fontsize=12, color='red')
    if shape in ["Brace", "Post"]:
        ax.text(WB_x, WB_y, f"WO: {WO:.2f}", ha='center', va='center', fontsize=12, color='blue')
    else:
        ax.text(WB_x, WB_y, f"WB: {WB:.2f}", ha='center', va='center', fontsize=12, color='blue')
    ax.text(HR_x, HR_y, f"HR: {HR:.2f}", ha='left', va='center', fontsize=12, color='green')
    ax.text(WT_x, WT_y, f"WT: {WT:.2f}", ha='center', va='center', fontsize=12, color='orange')
    ax.text(TH_x, TH_y, f"Th: {TH:.2f}", ha='center', va='center', fontsize=14, color='purple')
    ax.text(SectionName_x, SectionName_y, f"\n{section_name}", ha='center', va='bottom', fontsize=14, color='black')
    ax.text(Subshape_x, Subshape_y, subshape, ha='center', va='center', fontsize=30, color='black')

    ## ---------- Ensure output folder exists / اطمینان از وجود فولدر مقصد ----------
    os.makedirs(os.path.dirname(png_path), exist_ok=True)

    ## ---------- Save PNG / ذخیره PNG ----------
    plt.savefig(png_path, dpi=300, bbox_inches="tight")
    plt.close(fig)

print("\n ✅ All shapes rendered correctly using single-path column for DXF/PNG.")

import pandas as pd
import ezdxf
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
import matplotlib.pyplot as plt
import os

# ---------- خواندن دیتابیس ----------
df = pd.read_excel("data.xlsx", header=1)

# بررسی ستون‌های xl و yb
if "xl  =" not in df.columns:
    df["xl  ="] = 0
if "yb  =" not in df.columns:
    df["yb  ="] = 0

# حذف ردیف‌هایی که مقادیر ضروری ندارند
df = df.dropna(subset=["Shape", "Subshape", "WT", "H", "WB", "HR", "Thickness"])

# پوشه خروجی
os.makedirs("output", exist_ok=True)

for _, row in df.iterrows():
    section_name = str(row["Section Name"]).strip()
    shape = str(row["Shape"]).strip()
    subshape = str(row["Subshape"]).strip()

    dxf_path = f"templates/{section_name}.dxf"
    if not os.path.exists(dxf_path):
        print(f"DXF file not found: {dxf_path}")
        continue

    # خواندن مقادیر از دیتابیس
    WT, H, WB, HR, TH = row["WT"], row["H"], row["WB"], row["HR"], row["Thickness"]
    XL, YB = row["xl  ="], row["yb  ="]

    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    # ---------- جمع آوری نقاط ----------
    points_x, points_y = [], []
    for e in msp:
        try:
            if e.dxftype() in ('LINE', 'LWPOLYLINE', 'POLYLINE'):
                pts = [e.dxf.start, e.dxf.end] if e.dxftype() == 'LINE' else [p[:2] for p in e.get_points()]
                for x, y in pts:
                    points_x.append(x)
                    points_y.append(y)
        except Exception:
            pass

    # ---------- محاسبه مختصات متن‌ها طبق دیتابیس ----------
    # H
    H_x = -XL-(2*TH)
    H_y = (H / 2) - YB

    # WB
    WB_x = -XL + WB/2
    WB_y = -YB -(2*TH)

    # HR 
    HR_x = (-XL + WB)+(2*TH)
    HR_y = -YB + HR/2

    #TH
    TH_x = 0
    TH_y = -YB + HR/2
    
    # WT 
    WT_x = -XL + WT/2
    WT_y = (H - YB)+(2*TH)

    # Subshape
    Subshape_x = 0
    Subshape_y = TH_y + 0.2
    
    #title
    SectionName_x = -XL + WT/2
    SectionName_y = H/2 + 0.4

    # ---------- چاپ مختصات برای بررسی ----------
    # print(f"\nSection: {section_name}")
    # print(f"Calculated H: ({H_x:.3f}, {H_y:.3f})")
    # print(f"Calculated WB: ({WB_x:.3f}, {WB_y:.3f})")
    # print(f"Calculated HR: ({HR_x:.3f}, {HR_y:.3f})")
    # print(f"Calculated WT: ({WT_x:.3f}, {WT_y:.3f})")
    # print(f"Calculated TH: ({TH_x:.3f}, {TH_y:.3f})")

    # ---------- رندر ----------
    fig, ax = plt.subplots(figsize=(6,6))
    ax.set_aspect("equal")
    ax.axis("off")

    frontend = Frontend(RenderContext(doc), MatplotlibBackend(ax))
    frontend.draw_layout(msp)

    # اضافه کردن متن‌ها روی شکل
    ax.text(H_x, H_y, f"H: {H:.2f}", ha='right', va='center', fontsize=12, color='red')
    ax.text(WB_x, WB_y, f"WB: {WB:.2f}", ha='center', va='center', fontsize=12, color='blue')
    ax.text(HR_x, HR_y, f"HR: {HR:.2f}", ha='left', va='center', fontsize=12, color='green')
    ax.text(WT_x, WT_y, f"WT: {WT:.2f}", ha='center', va='center', fontsize=12, color='orange')
    ax.text(TH_x, TH_y, f"Th: {TH:.2f}", ha='center', va='center', fontsize=14, color='purple')
    ax.text(SectionName_x, SectionName_y, f"\n{section_name}", ha='center', va='bottom', fontsize=14, color='black')
    ax.text(Subshape_x, Subshape_y, subshape , ha='center', va='center', fontsize=30, color='black')


    # ذخیره PNG
    plt.savefig(f"output/{section_name}.png", dpi=300, bbox_inches="tight")
    plt.close(fig)

print("\nAll shapes rendered with H, WB, HR, WT, Th, and Section Name correctly positioned according to database formulas.")

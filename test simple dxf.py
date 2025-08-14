import ezdxf
import matplotlib.pyplot as plt
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
import io
from PIL import Image

def replace_blue_with_black(img):
    pixels = img.load()
    w, h = img.size
    for x in range(w):
        for y in range(h):
            r, g, b, a = pixels[x, y]
            if a > 0:
                if b > 100 and b > r + 20 and b > g + 20:
                    pixels[x, y] = (0, 0, 0, a)
    return img

def render_dxf_to_png(dxf_path):
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    fig, ax = plt.subplots(figsize=(6,6))
    ax.set_aspect('equal')
    ax.axis('off')

    ax.set_facecolor('white')
    fig.patch.set_facecolor('white')

    ctx = RenderContext(doc)
    frontend = Frontend(ctx, MatplotlibBackend(ax))

    frontend.draw_layout(msp)

    buf = io.BytesIO()
    fig.savefig(buf, format='png', transparent=False)
    plt.close(fig)
    buf.seek(0)

    img = Image.open(buf).convert("RGBA")
    img = replace_blue_with_black(img)
    #img.show()   #just for test 

if __name__ == "__main__":
    render_dxf_to_png("templates/Box Beam.A.dxf")

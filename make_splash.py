"""Generate splash.png for PyInstaller build."""
from PIL import Image, ImageDraw, ImageFont

W, H = 400, 200
BG     = (245, 245, 245)
BORDER = (80,  80,  80)
ACCENT = (60,  120, 200)
TEXT   = (40,  40,  40)

img  = Image.new("RGB", (W, H), BG)
draw = ImageDraw.Draw(img)

# border
draw.rectangle([0, 0, W-1, H-1], outline=BORDER, width=3)

# accent bar
bar_x0, bar_y0, bar_x1, bar_y1 = 40, 110, W-40, 130
for x in range(bar_x0, bar_x1):
    t = (x - bar_x0) / (bar_x1 - bar_x0)
    r = int(ACCENT[0] * t + BG[0] * (1 - t))
    g = int(ACCENT[1] * t + BG[1] * (1 - t))
    b = int(ACCENT[2] * t + BG[2] * (1 - t))
    draw.line([(x, bar_y0), (x, bar_y1)], fill=(r, g, b))

# text
try:
    font = ImageFont.truetype("C:/Windows/Fonts/msjh.ttc", 28)  # 微軟正黑體
except Exception:
    font = ImageFont.load_default()

label = "啟動中..."
bbox  = draw.textbbox((0, 0), label, font=font)
tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
draw.text(((W - tw) // 2, 60), label, font=font, fill=TEXT)

img.save("splash.png")
print("splash.png 建立完成")

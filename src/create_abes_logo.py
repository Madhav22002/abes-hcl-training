"""
Generate high-resolution ABES Engineering College official crest and emblem for project cover page.
"""

from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
import math

BASE_DIR = Path(__file__).resolve().parent.parent
DOCS_DIR = BASE_DIR / "docs"
LOGO_PATH = DOCS_DIR / "abes_logo.png"

def create_abes_logo():
    # 800 x 800 high-res logo
    size = (800, 800)
    img = Image.new("RGBA", size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)

    # Fonts
    try:
        font_large = ImageFont.truetype("arialbd.ttf", 46)
        font_med = ImageFont.truetype("arialbd.ttf", 30)
        font_small = ImageFont.truetype("arial.ttf", 22)
        font_estd = ImageFont.truetype("arialbd.ttf", 24)
    except Exception:
        font_large = ImageFont.load_default()
        font_med = ImageFont.load_default()
        font_small = ImageFont.load_default()
        font_estd = ImageFont.load_default()

    # Outer Circular Border - Navy Blue (#0F2B5C)
    navy = (15, 43, 92, 255)
    gold = (212, 160, 23, 255)
    maroon = (139, 21, 56, 255)

    center = (400, 400)
    outer_r = 370
    inner_r = 350
    ring_r = 250

    # Concentric circles
    draw.ellipse([center[0]-outer_r, center[1]-outer_r, center[0]+outer_r, center[1]+outer_r], fill=navy)
    draw.ellipse([center[0]-inner_r, center[1]-inner_r, center[0]+inner_r, center[1]+inner_r], fill=(255, 255, 255, 255))
    draw.ellipse([center[0]-ring_r, center[1]-ring_r, center[0]+ring_r, center[1]+ring_r], fill=navy)

    # Center Shield / Emblem area
    center_shield_r = 230
    draw.ellipse([center[0]-center_shield_r, center[1]-center_shield_r, center[0]+center_shield_r, center[1]+center_shield_r], fill=(248, 250, 252, 255))

    # Inner Gold Border
    draw.ellipse([center[0]-center_shield_r+8, center[1]-center_shield_r+8, center[0]+center_shield_r-8, center[1]+center_shield_r-8], outline=gold, width=6)

    # Text in the ring
    # Top text: "ABES ENGINEERING COLLEGE"
    draw.text((150, 65), "ABES ENGINEERING COLLEGE", fill=navy, font=font_med)
    draw.text((320, 690), "GHAZIABAD", fill=navy, font=font_med)

    # Draw Center Acronym "ABES" in bold maroon/gold
    draw.text((235, 270), "ABES", fill=maroon, font=font_large)
    
    # Subtitle under ABES
    draw.line([(240, 345), (560, 345)], fill=gold, width=4)
    draw.text((285, 365), "LEADERS IN TECH", fill=navy, font=font_small)
    draw.text((335, 410), "ESTD. 2000", fill=maroon, font=font_estd)

    # Tech icon (Gear / book motif)
    # Gear teeth circle
    gear_center = (400, 520)
    gear_r = 60
    for deg in range(0, 360, 30):
        rad = math.radians(deg)
        x = gear_center[0] + (gear_r + 14) * math.cos(rad)
        y = gear_center[1] + (gear_r + 14) * math.sin(rad)
        draw.rectangle([x-6, y-6, x+6, y+6], fill=gold)
    draw.ellipse([gear_center[0]-gear_r, gear_center[1]-gear_r, gear_center[0]+gear_r, gear_center[1]+gear_r], fill=navy)
    draw.ellipse([gear_center[0]-25, gear_center[1]-25, gear_center[0]+25, gear_center[1]+25], fill=(255, 255, 255, 255))

    img.save(LOGO_PATH, "PNG")
    print(f"ABES logo generated and saved to {LOGO_PATH}")

if __name__ == "__main__":
    create_abes_logo()

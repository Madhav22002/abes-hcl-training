"""
Render ER Diagram PNG from Database Schema using PIL.
"""

from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

DOCS_DIR = Path(__file__).resolve().parent.parent / "docs"
PNG_PATH = DOCS_DIR / "09_ER_Diagram.png"

def draw_er_diagram():
    # 1600 x 1000 canvas
    width, height = 1600, 1050
    img = Image.new("RGB", (width, height), color="#F8FAFC")
    draw = ImageDraw.Draw(img)

    # Fonts
    try:
        title_font = ImageFont.truetype("segoeui.ttf", 28)
        header_font = ImageFont.truetype("segoeuib.ttf", 16)
        text_font = ImageFont.truetype("segoeui.ttf", 13)
        bold_font = ImageFont.truetype("segoeuib.ttf", 13)
        note_font = ImageFont.truetype("segoeui.ttf", 12)
    except Exception:
        title_font = ImageFont.load_default()
        header_font = ImageFont.load_default()
        text_font = ImageFont.load_default()
        bold_font = ImageFont.load_default()
        note_font = ImageFont.load_default()

    # Title Banner
    draw.rectangle([(0, 0), (width, 80)], fill="#1E3A8A")
    draw.text((40, 24), "Entity Relationship Diagram — Marketing Campaign Analytics Schema", fill="#FFFFFF", font=title_font)

    # Box coordinates
    # 4 dimension tables on left and right, 1 fact table in center
    boxes = {
        "campaigns": {"x": 50, "y": 140, "w": 380, "h": 260, "color": "#2563EB", "title": "campaigns (Dimension Master)"},
        "channels": {"x": 50, "y": 480, "w": 380, "h": 200, "color": "#0D9488", "title": "channels (Dimension Master)"},
        "audience_segments": {"x": 1170, "y": 140, "w": 380, "h": 230, "color": "#7C3AED", "title": "audience_segments (Dimension Master)"},
        "creatives": {"x": 1170, "y": 480, "w": 380, "h": 220, "color": "#EA580C", "title": "creatives (Dimension Master)"},
        "daily_performance": {"x": 540, "y": 240, "w": 520, "h": 460, "color": "#DC2626", "title": "daily_performance (Fact Table)"}
    }

    table_fields = {
        "campaigns": [
            ("campaign_id", "TEXT", "PK"),
            ("campaign_name", "TEXT", "NOT NULL"),
            ("objective", "TEXT", "NOT NULL"),
            ("start_date", "TEXT", "NOT NULL"),
            ("end_date", "TEXT", "NOT NULL"),
            ("budget_inr", "REAL", "NOT NULL CHECK > 0"),
            ("status", "TEXT", "Active/Completed/Paused")
        ],
        "channels": [
            ("channel_id", "TEXT", "PK"),
            ("channel_name", "TEXT", "NOT NULL"),
            ("channel_type", "TEXT", "NOT NULL"),
            ("cost_model", "TEXT", "CPC/CPM/CPA/CPV")
        ],
        "audience_segments": [
            ("segment_id", "TEXT", "PK"),
            ("segment_name", "TEXT", "NOT NULL"),
            ("region", "TEXT", "NOT NULL (Indian Territory)"),
            ("target_age_band", "TEXT", "NOT NULL"),
            ("primary_device", "TEXT", "Desktop/Mobile/Multi")
        ],
        "creatives": [
            ("creative_id", "TEXT", "PK"),
            ("creative_name", "TEXT", "NOT NULL"),
            ("creative_format", "TEXT", "PDF/Video/Carousel/Tool"),
            ("messaging_theme", "TEXT", "NOT NULL"),
            ("offer_type", "TEXT", "Download/Demo/Trial")
        ],
        "daily_performance": [
            ("fact_id", "INTEGER", "PK (AUTOINCREMENT)"),
            ("report_date", "TEXT", "NOT NULL (Grain Component)"),
            ("campaign_id", "TEXT", "FK -> campaigns.campaign_id"),
            ("channel_id", "TEXT", "FK -> channels.channel_id"),
            ("segment_id", "TEXT", "FK -> audience_segments.segment_id"),
            ("creative_id", "TEXT", "FK -> creatives.creative_id"),
            ("impressions", "INTEGER", "NOT NULL CHECK >= 0"),
            ("clicks", "INTEGER", "NOT NULL CHECK >= 0"),
            ("conversions", "INTEGER", "NOT NULL CHECK >= 0"),
            ("spend_inr", "REAL", "NOT NULL CHECK >= 0"),
            ("revenue_inr", "REAL", "NOT NULL CHECK >= 0"),
            ("reach", "INTEGER", "NOT NULL CHECK >= 0")
        ]
    }

    # Draw Boxes
    for t_key, box in boxes.items():
        bx, by, bw, bh = box["x"], box["y"], box["w"], box["h"]
        # Outer Border
        draw.rounded_rectangle([(bx, by), (bx + bw, by + bh)], radius=8, fill="#FFFFFF", outline="#CBD5E1", width=2)
        # Header Bar
        draw.rounded_rectangle([(bx, by), (bx + bw, by + 36)], radius=8, fill=box["color"])
        draw.rectangle([(bx, by + 28), (bx + bw, by + 36)], fill=box["color"])
        draw.text((bx + 12, by + 8), box["title"], fill="#FFFFFF", font=header_font)

        # Fields
        cur_y = by + 46
        fields = table_fields[t_key]
        for f_name, f_type, f_con in fields:
            is_pk = "PK" in f_con
            is_fk = "FK" in f_con
            badge = "[PK]" if is_pk else ("[FK]" if is_fk else "    ")
            b_color = "#DC2626" if is_pk else ("#2563EB" if is_fk else "#64748B")
            
            draw.text((bx + 12, cur_y), badge, fill=b_color, font=bold_font)
            draw.text((bx + 52, cur_y), f_name, fill="#1E293B", font=bold_font if (is_pk or is_fk) else text_font)
            draw.text((bx + 210, cur_y), f_type, fill="#475569", font=text_font)
            draw.text((bx + 280, cur_y), f_con[:22], fill="#64748B", font=note_font)
            cur_y += 24

    # Draw Connector Lines (1:N relationships)
    # 1. campaigns -> daily_performance
    draw.line([(430, 270), (540, 290)], fill="#2563EB", width=3)
    draw.polygon([(535, 285), (540, 290), (535, 295)], fill="#2563EB")
    draw.text((445, 260), "1 : N (generates)", fill="#2563EB", font=note_font)

    # 2. channels -> daily_performance
    draw.line([(430, 560), (540, 480)], fill="#0D9488", width=3)
    draw.polygon([(535, 475), (540, 480), (535, 485)], fill="#0D9488")
    draw.text((445, 530), "1 : N (delivers)", fill="#0D9488", font=note_font)

    # 3. audience_segments -> daily_performance
    draw.line([(1170, 260), (1060, 320)], fill="#7C3AED", width=3)
    draw.polygon([(1065, 315), (1060, 320), (1065, 325)], fill="#7C3AED")
    draw.text((1075, 275), "1 : N (targets)", fill="#7C3AED", font=note_font)

    # 4. creatives -> daily_performance
    draw.line([(1170, 560), (1060, 500)], fill="#EA580C", width=3)
    draw.polygon([(1065, 495), (1060, 500), (1065, 505)], fill="#EA580C")
    draw.text((1075, 540), "1 : N (deploys)", fill="#EA580C", font=note_font)

    # Bottom Footer
    footer_text = "Enforced: PRAGMA foreign_keys = ON | Composite Grain Rule: UNIQUE(report_date, campaign_id, channel_id, segment_id, creative_id)"
    draw.rectangle([(50, 930), (1550, 980)], fill="#E2E8F0", outline="#CBD5E1")
    draw.text((70, 946), footer_text, fill="#334155", font=bold_font)

    img.save(PNG_PATH, "PNG")
    print(f"ER Diagram PNG successfully saved to {PNG_PATH}")

if __name__ == "__main__":
    draw_er_diagram()

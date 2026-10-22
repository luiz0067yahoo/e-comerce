"""
Slide builder helper utilities for rendering high-res graphics
"""

import os
import sys
import math
from PIL import Image, ImageDraw, ImageFont
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# Ensure temp directory exists
TEMP_DIR = os.path.join(os.path.dirname(__file__), "temp_slides")
os.makedirs(TEMP_DIR, exist_ok=True)

# Fonts
FONT_HERO = ImageFont.truetype(r"C:\Windows\Fonts\segoeuib.ttf", 64)
FONT_TITLE = ImageFont.truetype(r"C:\Windows\Fonts\segoeuib.ttf", 46)
FONT_SUBTITLE = ImageFont.truetype(r"C:\Windows\Fonts\segoeui.ttf", 24)
FONT_HEADING = ImageFont.truetype(r"C:\Windows\Fonts\segoeuib.ttf", 28)
FONT_BODY = ImageFont.truetype(r"C:\Windows\Fonts\segoeui.ttf", 20)
FONT_BODY_BOLD = ImageFont.truetype(r"C:\Windows\Fonts\segoeuib.ttf", 20)
FONT_SMALL = ImageFont.truetype(r"C:\Windows\Fonts\segoeui.ttf", 16)
FONT_SMALL_BOLD = ImageFont.truetype(r"C:\Windows\Fonts\segoeuib.ttf", 16)
FONT_CODE = ImageFont.truetype(r"C:\Windows\Fonts\consola.ttf", 18)

# Color Palettes
C_DARK_BG = (15, 23, 42)      # Slate 900
C_CARD_BG = (30, 41, 59)      # Slate 800
C_CARD_BORDER = (51, 65, 85)  # Slate 700
C_TEXT_LIGHT = (248, 250, 252)
C_TEXT_MUTED = (148, 163, 184)
C_WHITE = (255, 255, 255)

# Module accents
ACCENTS = {
    "woo": {"primary": (127, 84, 179), "secondary": (155, 114, 207), "light": (243, 238, 248), "name": "WooCommerce"},
    "meli": {"primary": (255, 214, 0), "secondary": (0, 158, 227), "light": (255, 249, 219), "name": "Mercado Livre & Mercado Pago"},
    "frete": {"primary": (14, 116, 144), "secondary": (245, 158, 11), "light": (236, 254, 255), "name": "Frete & Logística"},
    "prod": {"primary": (79, 70, 229), "secondary": (16, 185, 129), "light": (238, 242, 255), "name": "Produtos & Catálogo"},
    "compra": {"primary": (16, 185, 129), "secondary": (37, 99, 235), "light": (236, 253, 245), "name": "Simulação de Compra"}
}

def create_base_canvas(w=1920, h=1080, bg=C_DARK_BG):
    img = Image.new("RGB", (w, h), color=bg)
    draw = ImageDraw.Draw(img)
    return img, draw

def draw_header(draw, tag, title, subtitle, accent_color=(127, 84, 179)):
    draw.rounded_rectangle([80, 50, 80 + len(tag)*13 + 30, 88], radius=8, fill=accent_color)
    tag_text_color = (0, 0, 0) if (accent_color[0]*0.299 + accent_color[1]*0.587 + accent_color[2]*0.114) > 150 else C_WHITE
    draw.text((95, 58), tag.upper(), font=FONT_SMALL_BOLD, fill=tag_text_color)
    
    draw.text((80, 102), title, font=FONT_TITLE, fill=C_TEXT_LIGHT)
    draw.text((80, 168), subtitle, font=FONT_SUBTITLE, fill=C_TEXT_MUTED)
    draw.line([80, 215, 1840, 215], fill=C_CARD_BORDER, width=2)

def draw_browser_chrome(draw, x, y, w, h, title="Painel WordPress", url="https://meusite.com/wp-admin/"):
    draw.rounded_rectangle([x, y, x + w, y + h], radius=16, fill=(24, 30, 42), outline=(56, 68, 88), width=2)
    draw.rounded_rectangle([x, y, x + w, y + 48], radius=16, fill=(18, 22, 32))
    draw.rectangle([x, y + 32, x + w, y + 48], fill=(18, 22, 32))
    draw.ellipse([x + 18, y + 17, x + 32, y + 31], fill=(239, 68, 68))
    draw.ellipse([x + 38, y + 17, x + 52, y + 31], fill=(245, 158, 11))
    draw.ellipse([x + 58, y + 17, x + 72, y + 31], fill=(16, 185, 129))
    draw.rounded_rectangle([x + 90, y + 10, x + w - 240, y + 38], radius=8, fill=(34, 43, 60))
    draw.text((x + 105, y + 14), url, font=FONT_SMALL, fill=(180, 195, 215))
    draw.text((x + w - 220, y + 14), title, font=FONT_SMALL_BOLD, fill=(140, 155, 175))

def draw_wp_sidebar(draw, x, y, w, h, active_item="Plugins"):
    items = [
        "Painel", "Posts", "Mídia", "Páginas", "Comentários",
        "WooCommerce", "Produtos", "Aparência", "Plugins", "Usuários", "Ferramentas", "Configurações"
    ]
    draw.rectangle([x, y, x + w, y + h], fill=(29, 35, 39))
    item_y = y + 20
    for item in items:
        is_active = (item.lower() in active_item.lower())
        if is_active:
            draw.rectangle([x, item_y - 6, x + w, item_y + 30], fill=(34, 113, 177))
            draw.text((x + 24, item_y), item, font=FONT_BODY_BOLD, fill=C_WHITE)
        else:
            draw.text((x + 24, item_y), item, font=FONT_BODY, fill=(180, 190, 200))
        item_y += 42

def draw_arrow_pointer(draw, start, end, color=(239, 68, 68), width=5):
    x1, y1 = start
    x2, y2 = end
    draw.line([x1, y1, x2, y2], fill=color, width=width)
    angle = math.atan2(y2 - y1, x2 - x1)
    arrow_len = 22
    arrow_angle = math.pi / 6
    x3 = x2 - arrow_len * math.cos(angle - arrow_angle)
    y3 = y2 - arrow_len * math.sin(angle - arrow_angle)
    x4 = x2 - arrow_len * math.cos(angle + arrow_angle)
    y4 = y2 - arrow_len * math.sin(angle + arrow_angle)
    draw.polygon([(x2, y2), (x3, y3), (x4, y4)], fill=color)

def draw_check_icon(draw, cx, cy, radius=12, bg=(16, 185, 129), fg=C_WHITE):
    draw.ellipse([cx - radius, cy - radius, cx + radius, cy + radius], fill=bg)
    # Check lines
    p1 = (cx - radius*0.45, cy)
    p2 = (cx - radius*0.1, cy + radius*0.4)
    p3 = (cx + radius*0.5, cy - radius*0.35)
    draw.line([p1, p2, p3], fill=fg, width=int(radius*0.25))

def draw_cross_icon(draw, cx, cy, radius=12, bg=(239, 68, 68), fg=C_WHITE):
    draw.ellipse([cx - radius, cy - radius, cx + radius, cy + radius], fill=bg)
    d = radius * 0.4
    draw.line([(cx - d, cy - d), (cx + d, cy + d)], fill=fg, width=int(radius*0.25))
    draw.line([(cx - d, cy + d), (cx + d, cy - d)], fill=fg, width=int(radius*0.25))

def draw_bullet_dot(draw, cx, cy, radius=5, color=(155, 114, 207)):
    draw.ellipse([cx - radius, cy - radius, cx + radius, cy + radius], fill=color)

def draw_star_badge(draw, x, y, count=5):
    for i in range(count):
        sx = x + i * 22
        # Simple gold star diamond
        draw.polygon([(sx + 10, y), (sx + 14, y + 7), (sx + 20, y + 8),
                      (sx + 15, y + 13), (sx + 17, y + 20), (sx + 10, y + 16),
                      (sx + 3, y + 20), (sx + 5, y + 13), (sx, y + 8), (sx + 6, y + 7)],
                     fill=(245, 158, 11))

def draw_pill_badge(draw, x, y, text, bg, fg=C_WHITE):
    w = len(text) * 11 + 24
    draw.rounded_rectangle([x, y, x + w, y + 32], radius=16, fill=bg)
    draw.text((x + 12, y + 5), text, font=FONT_SMALL_BOLD, fill=fg)
    return w

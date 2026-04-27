"""Create Bank Leumi logo as PNG using PIL."""
from PIL import Image, ImageDraw, ImageFont
import math

W, H = 400, 400

img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
draw = ImageDraw.Draw(img)

# Bank Leumi blue
BLUE = (0, 75, 155)
DARK_BLUE = (0, 40, 100)
WHITE = (255, 255, 255)

# Rounded rectangle background
radius = 80
draw.rounded_rectangle([0, 0, W - 1, H - 1], radius=radius, fill=BLUE)

# White wave in upper portion
wave_y = 180
points = []
steps = 200
for i in range(steps + 1):
    x = i * W / steps
    # Two humps wave
    y = wave_y - 40 * math.sin(math.pi * x / (W / 2))
    points.append((x, y))

# Create wave polygon (fill from wave down to mid-point)
poly = [(0, H)] + [(0, wave_y)] + points + [(W, wave_y)] + [(W, H)]
draw.polygon(poly, fill=WHITE)

# Dark blue bottom band
draw.rounded_rectangle([0, H // 2 + 20, W - 1, H - 1], radius=radius, fill=DARK_BLUE)
# Fix corners of the band
draw.rectangle([0, H // 2 + 20, W - 1, H // 2 + 50], fill=DARK_BLUE)

# Hebrew text "לאומי" - use a basic font
try:
    font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 80)
except Exception:
    font = ImageFont.load_default()

text = "לאומי"
bbox = draw.textbbox((0, 0), text, font=font)
tw = bbox[2] - bbox[0]
th = bbox[3] - bbox[1]
tx = (W - tw) // 2
ty = H - th - 40

draw.text((tx, ty), text, font=font, fill=WHITE)

img.save("/home/user/text/leumi_logo.png", "PNG")
print("Logo saved.")

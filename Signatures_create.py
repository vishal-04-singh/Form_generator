import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os

# Load Excel
df = pd.read_excel("p1.xlsx")  # your participant sheet
output_folder = "signatures"
os.makedirs(output_folder, exist_ok=True)

# Load a handwriting-style font
font_path = "Dancing-Script.ttf"  # download & place in same folder
font_size = 48

# Load font
try:
    font = ImageFont.truetype(font_path, font_size)
except:
    raise Exception("Font not found. Make sure 'Dancing-Script.ttf' exists.")

# Generate signature images
for _, row in df.iterrows():
    name = str(row['Name']).strip()

    img = Image.new('RGBA', (400, 100), (255, 255, 255, 0))  # transparent background
    draw = ImageDraw.Draw(img)

    # Modern way to calculate text size
    bbox = draw.textbbox((0, 0), name, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]

    x = (400 - text_width) // 2
    y = (100 - text_height) // 2

    draw.text((x, y), name, fill=(0, 0, 0), font=font)

    output_path = os.path.join(output_folder, f"{name}.png")
    img.save(output_path, "PNG")

print("✅ All signature images generated and saved to 'signatures/' folder.")

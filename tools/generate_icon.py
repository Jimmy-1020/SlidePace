"""Draw the product icon using geometric shapes (development helper only)."""
from pathlib import Path
from PIL import Image, ImageDraw

asset_dir = Path(__file__).resolve().parent.parent / 'assets'
asset_dir.mkdir(exist_ok=True)
image = Image.new('RGBA', (256, 256), (0, 0, 0, 0))
draw = ImageDraw.Draw(image)
draw.rounded_rectangle((8, 8, 248, 248), radius=52, fill='#161D27')
draw.ellipse((46, 46, 210, 210), outline='#44B5A0', width=12)
draw.line((128, 70, 128, 128, 166, 151), fill='#E5EDF7', width=13)
draw.ellipse((119, 119, 137, 137), fill='#E5EDF7')
draw.line((109, 30, 147, 30), fill='#44B5A0', width=12)
draw.line((128, 30, 128, 43), fill='#44B5A0', width=10)
image.save(asset_dir / 'SlidePace.ico', sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
image.save(asset_dir / 'SlidePace.png')
print('Created assets/SlidePace.ico')

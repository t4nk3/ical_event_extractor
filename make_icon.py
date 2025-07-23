from PIL import Image, ImageDraw, ImageFont

# Create a blank white image
img = Image.new('RGBA', (512, 512), 'white')
draw = ImageDraw.Draw(img)

# Draw calendar outline
draw.rounded_rectangle([32, 64, 480, 480], radius=48, outline='#2d3a4a', width=12, fill='#f7faff')

# Draw blue header
draw.rectangle([32, 64, 480, 150], fill='#3b82f6')

# Draw grid (5 rows x 7 columns)
grid_top = 170
grid_left = 52
grid_right = 460
grid_bottom = 470
cell_w = (grid_right - grid_left) // 7
cell_h = (grid_bottom - grid_top) // 5
for i in range(1, 7):
    x = grid_left + i * cell_w
    draw.line([(x, grid_top), (x, grid_bottom)], fill='#bcd0e5', width=4)
for i in range(1, 5):
    y = grid_top + i * cell_h
    draw.line([(grid_left, y), (grid_right, y)], fill='#bcd0e5', width=4)

# Draw app initials
try:
    font = ImageFont.truetype("arial.ttf", 90)
except:
    font = ImageFont.load_default()
draw.text((256, 100), "iCal", anchor="mm", font=font, fill='white')

img.save('icon.png')
print("Icon saved as icon.png")
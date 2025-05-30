from PIL import Image

# Load the uploaded image
image_path = "Amkor-logo---CMYK.png"
img = Image.open(image_path)

# Convert to square by adding padding if necessary
width, height = img.size
if width != height:
    new_size = max(width, height)
    new_img = Image.new("RGB", (new_size, new_size), (255, 255, 255))
    new_img.paste(img, ((new_size - width) // 2, (new_size - height) // 2))
else:
    new_img = img

# Resize to 32x32
new_img = new_img.resize((32, 32), Image.Resampling.LANCZOS)

# Save as PNG for preview
png_path = "icon_preview.png"
new_img.save(png_path, format="PNG")

print(f"The icon has been successfully converted and saved as {png_path}.")


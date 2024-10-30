from pdf2image import convert_from_path

# Specify the path to your PDF file
pdf_path = "/Users/kosiew/Downloads/WIX2002_1_2024.pdf"

# Convert PDF pages to images
images = convert_from_path(pdf_path)

# Save each page as a separate image file
for i, image in enumerate(images):
    image_path = f"/Users/kosiew/Downloads/page_{i + 1}.png"
    image.save(image_path, "PNG")
    print(f"Saved {image_path}")

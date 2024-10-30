import os
import sys
from pdf2image import convert_from_path


def extract_images_from_pdf(pdf_path, output_folder):
    """Extracts pages from a PDF as images and saves them to an output folder."""
    # Convert PDF pages to images
    images = convert_from_path(pdf_path)

    # Save each page as a separate image file
    for i, image in enumerate(images):
        image_name = f"page_{i + 1}.png"
        image_path = os.path.join(output_folder, image_name)
        image.save(image_path, "PNG")
        print(f"Saved: {image_path}")


def process_pdf_in_folder(folder_path):
    """Processes all PDFs found in the given folder and its subdirectories."""
    # Walk through the folder and subfolders
    for root, _, files in os.walk(folder_path):
        for file_name in files:
            if file_name.lower().endswith(".pdf"):
                # Construct full file path
                pdf_path = os.path.join(root, file_name)

                # Create an output folder named after the PDF (without the .pdf extension)
                output_folder = os.path.join(root, os.path.splitext(file_name)[0])
                os.makedirs(output_folder, exist_ok=True)

                # Extract images from the PDF and save them in the output folder
                print(f"Processing PDF: {pdf_path}")
                extract_images_from_pdf(pdf_path, output_folder)


def main():
    """Process each folder passed as a command-line argument."""
    # Check if folders were provided as arguments
    if len(sys.argv) < 2:
        print("Usage: python script_name.py <folder1> <folder2> ...")
        sys.exit(1)

    # Get folders from command-line arguments
    input_folders = sys.argv[1:]

    for folder in input_folders:
        if os.path.isdir(folder):
            print(f"Scanning folder: {folder}")
            process_pdf_in_folder(folder)
        else:
            print(f"Warning: {folder} is not a valid directory.")


if __name__ == "__main__":
    main()

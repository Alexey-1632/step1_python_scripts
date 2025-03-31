import fitz
import os
import argparse
import logging
from PIL import Image  # Import Pillow for image handling

# It describe which color conversion algorithm will be applied to the particular PDF file
# The digits are the file number in the "PDF_links.xlsx" file
color_groups = {
    "Gayscale-128": ["072", "073", "113", "124", "161", "211"],
    "Gayscale-232": ["053", "054", "055", "062", "094", "128", "129", "140", "145", "146", "184", "186", "223", "233"],
    "BW-128": ["047", "048", "049", "050", "123", "151", "164", "183", "185", "203", "207"],
    "BW-196": ["001", "002", "003", "004", "005", "006", "007", "008", "009", "010", "011", "012", "014", "015", "016", "017", "019", "020", "021", "022", "023", "024", "025", "026", "027", "028", "029", "030", "031", "032", "033", "034", "035", "036", "037", "038", "039", "040", "041", "042", "043", "044", "045", "046", "051", "052", "056", "057", "058", "059", "060", "061", "063", "064", "065", "066", "067", "068", "069", "071", "075", "076", "077", "078", "079", "080", "081", "082", "083", "084", "086", "087", "088", "089", "090", "091", "092", "093", "095", "096", "097", "098", "099", "100", "101", "102", "103", "104", "105", "106", "107", "108", "109", "110", "111", "112", "114", "115", "116", "117", "118", "119", "120", "121", "122", "125", "126", "127", "130", "131", "132", "133", "134", "135", "136", "137", "138", "139", "141", "142", "143", "144", "147", "148", "149", "150", "152", "153", "154", "155", "156", "157", "158", "159", "160", "162", "163", "165", "166", "167", "168", "169", "170", "171", "172", "173", "174", "175", "176", "177", "178", "179", "180", "181", "182", "187", "188", "189", "190", "191", "192", "193", "194", "195", "196", "197", "198", "199", "200", "201", "202", "204", "205", "206", "208", "209", "210", "212", "213", "214", "215", "216", "217", "218", "219", "220", "221", "222", "224", "225", "226", "227", "228", "231", "232", "235", "236"]
}

# Configure logging
logging.basicConfig(
    format="%(asctime)s - %(levelname)s - %(message)s",
    level=logging.INFO
)

def save_pdf_pages_as_images(pdf_path, output_folder, base_name, jpeg_quality, dpi=400):
    pdf_document = fitz.open(pdf_path)
    logging.info(f"Processing PDF: {pdf_path}")  # Log the PDF being processed

    # Calculate scaling factors for resolution adjustment
    scale = dpi / 72  # Default DPI in PyMuPDF is 72
    matrix = fitz.Matrix(scale, scale)

    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        pix = page.get_pixmap(matrix=matrix)  # Apply scaling for higher resolution

        # Convert Pixmap to a PIL Image
        image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # Convert to grayscale
        image = image.convert("L")

        prefix = base_name[:3]  # Extract the first 3 characters: 001, 002, ..., 236

        for group, prefixes in color_groups.items():
            if prefix in prefixes:
                match group:
                    case "Gayscale-128":
                        threshold = 128  # Value range: 0-255

                        # keep some grayscale info 
                        image = image.point(lambda x: 255 if x > threshold else x)

                        logging.info("Gayscale-128")
                    case "Gayscale-232":
                        threshold = 232  # Value range: 0-255
                        
                        # keep some grayscale info 
                        image = image.point(lambda x: 255 if x > threshold else x)

                        logging.info("Gayscale-232")
                    case "BW-128":
                        threshold = 128  # Value range: 0-255

                        # convert to B/W image
                        image = image.point(lambda x: 255 if x > threshold else 0, "1")

                        logging.info("BW-128")
                    case "BW-196":
                        threshold = 196  # Value range: 0-255

                        # convert to B/W image
                        image = image.point(lambda x: 255 if x > threshold else 0, "1")
                        
                        logging.info("BW-196")

        # Set output image path
        image_path = os.path.join(
            output_folder, f"{base_name}_page{page_num + 1}"
        )

        # Save the image with the specified JPEG quality
        # Commented out because png give less file size
        #image.save(image_path + ".jpg", "JPEG", quality=jpeg_quality, optimize=True)
        image.save(image_path + ".png", "PNG")

        logging.info(f"Saved page {page_num + 1} as {image_path}")  # Log each page saved

    pdf_document.close()

def pdf_to_images(input_folder, output_folder, jpeg_quality=75):
    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)
    logging.info(f"Input folder: {input_folder}, Output folder: {output_folder}")  # Log folders

    # Iterate through all files in the input folder
    for filename in os.listdir(input_folder):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(input_folder, filename)
            base_name = os.path.splitext(filename)[0]
            save_pdf_pages_as_images(pdf_path, output_folder, base_name, jpeg_quality)
            logging.info(f"Completed processing: {filename}")  # Log completion of PDF processing

if __name__ == "__main__":
    # Set up argument parsing
    parser = argparse.ArgumentParser(description="Convert PDF files to images.")
    parser.add_argument(
        "input_folder",
        type=str,
        help="Path to the input folder containing PDF files."
    )
    parser.add_argument(
        "output_folder",
        type=str,
        help="Path to the output folder where images will be saved."
    )
    parser.add_argument(
        "--jpeg_quality",
        type=int,
        default=75,
        help="JPEG quality for output images (default is 75)."
    )

    # Parse the command-line arguments
    args = parser.parse_args()

    # Call the main function with parsed arguments
    pdf_to_images(args.input_folder, args.output_folder, args.jpeg_quality)

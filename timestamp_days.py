import os
import sys
import json
import random
import cairo
from PIL import Image
from datetime import datetime
import logging

# Setup logging configuration
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
logger = logging.getLogger(__name__)

# Load configuration from config.json
CONFIG_FILE = "config.json"
try:
    with open(CONFIG_FILE, "r", encoding="utf-8") as config_file:
        config = json.load(config_file)
except Exception as e:
    logger.error("Failed to load configuration file: %s", e)
    sys.exit(1)

folder_path = config.get("folder_path")
START_TIME = config.get("START_TIME")  # Expected format "HH:MM:SS"
DURATION = config.get("DURATION")      # Expected format "HH:MM:SS"
LOCATIONS = config.get("LOCATIONS")

if not folder_path or not START_TIME or not DURATION or not LOCATIONS:
    logger.error("Missing one or more required configuration parameters.")
    sys.exit(1)


def is_leaf_directory(path):
    """
    Check if the given directory is a leaf directory (contains no subdirectories).
    """
    for _, dirs, _ in os.walk(path):
        return len(dirs) == 0
    return True


def remove_stamped_photos(root):
    """
    Remove previously stamped photos from the given folder.
    """
    for file in os.listdir(root):
        if "_stamped" in file.lower():
            file_path = os.path.join(root, file)
            try:
                os.remove(file_path)
                logger.info("Removed stamped file: %s", file_path)
            except Exception as e:
                logger.error("Error removing file %s: %s", file_path, e)


def rotate_landscape_photos(root):
    """
    Rotate landscape images (width > height) by -90 degrees.
    """
    for file in os.listdir(root):
        file_path = os.path.join(root, file)
        if file.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".tiff")):
            try:
                with Image.open(file_path) as img:
                    if img.width > img.height:
                        rotated_img = img.rotate(-90, expand=True)
                        rotated_img.save(file_path)
                        logger.info("Rotated image: %s", file_path)
            except Exception as e:
                logger.error("Error processing file %s: %s", file_path, e)


def time_to_seconds(time_str):
    """
    Convert a time string in HH:MM:SS format to seconds.
    """
    h, m, s = map(int, time_str.split(":"))
    return h * 3600 + m * 60 + s


start_seconds = time_to_seconds(START_TIME)
duration_seconds = time_to_seconds(DURATION)


def generate_random_time():
    """
    Generate a random time (in seconds) within the specified duration.
    """
    offset = random.randint(0, duration_seconds)
    return start_seconds + offset


def format_date(date_obj):
    """
    Format the date as 'DD Mon. YYYY г.' with selected month names in Russian.
    """
    # Replace only 'Jan' and 'Dec' as per requirements
    months = {"Jan": "янв", "Dec": "дек"}
    formatted_date = date_obj.strftime("%d %b. %Y г.")
    for eng, rus in months.items():
        formatted_date = formatted_date.replace(eng, rus)
    return formatted_date


def process_image(image_path, date_obj):
    """
    Stamp the image with date/time and location text, and save with a '_stamped' suffix.
    """
    random_time = generate_random_time()
    location = random.choice(LOCATIONS)

    try:
        with Image.open(image_path) as img:
            # Resize image (maintain aspect ratio, maximum size 2000x2000)
            img.thumbnail((2000, 2000))
            width, height = img.width, img.height

            base_name, ext = os.path.splitext(image_path)
            output_image_path = f"{base_name}_stamped{ext}"

            # Calculate padding and font size relative to image dimensions
            padding_x = width * 0.004
            padding_y = height * 0.004
            font_size = int(height * 0.031)

            # Create a cairo surface to stamp text on the image
            surface = cairo.ImageSurface(cairo.FORMAT_ARGB32, width, height)
            context = cairo.Context(surface)

            # Save the current PIL image temporarily as PNG for cairo processing
            temp_png = output_image_path  # Using the same path for temporary PNG
            img.save(temp_png, format="PNG")

            # Load the saved PNG image into a cairo surface
            input_surface = cairo.ImageSurface.create_from_png(temp_png)
            context.set_source_surface(input_surface, 0, 0)
            context.paint()

            # Set font properties for stamping text
            context.select_font_face("Noto Sans", cairo.FONT_SLANT_NORMAL, cairo.FONT_WEIGHT_BOLD)
            context.set_font_size(font_size)
            context.set_source_rgb(1, 1, 1)  # White text

            # Format time string
            hh = random_time // 3600
            mm = (random_time % 3600) // 60
            ss = random_time % 60
            time_str = f"{format_date(date_obj)} {hh:02}:{mm:02}:{ss:02}"

            # Combine time string with location text (split into lines if needed)
            lines = [time_str] + location.split("\n")

            # Draw each line aligned to bottom-right
            text_y = padding_y + font_size
            for line in lines:
                text_extents = context.text_extents(line)
                text_x = width - text_extents.width - padding_x
                context.move_to(text_x, text_y)
                context.show_text(line)
                text_y += text_extents.height + font_size * 0.28

            # Write final stamped image to file
            surface.write_to_png(output_image_path)
            logger.info("Saved stamped image: %s", output_image_path)
    except Exception as e:
        logger.error("Error processing image %s: %s", image_path, e)


def main():
    """
    Main function to process images in leaf directories.
    """
    # Get leaf directories (directories without subdirectories)
    leaf_folders = [root for root, dirs, _ in os.walk(folder_path) if is_leaf_directory(root)]
    if not leaf_folders:
        logger.error("No leaf directories found.")
        sys.exit(1)

    for leaf_dir in leaf_folders:
        folder_name = os.path.basename(leaf_dir)
        try:
            # Expect folder name to start with date in format 'DD.MM.YYYY'
            date_str = folder_name[:10]
            date_obj = datetime.strptime(date_str, "%d.%m.%Y")
        except ValueError:
            logger.warning("Folder '%s' does not contain a valid date. Skipping.", folder_name)
            continue

        remove_stamped_photos(leaf_dir)
        rotate_landscape_photos(leaf_dir)

        # Filter image files that are not already stamped
        images = [
            f for f in os.listdir(leaf_dir)
            if f.lower().endswith((".jpg", ".jpeg", ".png"))
            and "_stamped" not in f.lower()
        ]
        if not images:
            logger.info("No suitable images found in folder: %s", leaf_dir)
            continue

        # Choose the image with the longest filename
        longest_name_image = max(images, key=len)
        image_path = os.path.join(leaf_dir, longest_name_image)
        process_image(image_path, date_obj)

    logger.info("Processing completed.")


if __name__ == "__main__":
    main()

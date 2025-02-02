import os
import json
import random
import logging
import re
from datetime import datetime, timedelta

import cairo
from PIL import Image, ImageChops

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)

# Load configuration from JSON file
CONFIG_PATH = "config.json"
try:
    with open(CONFIG_PATH, "r", encoding="utf-8") as config_file:
        config = json.load(config_file)
except Exception as e:
    logging.error("Failed to load config: %s", e)
    raise

FOLDER_PATH = config["folder_path"]
START_TIME = config["START_TIME"]
DURATION_BEFORE_WORKS = config["DURATION_BEFORE_WORKS"]
DURATION = config["DURATION"]
LOCATIONS = config["LOCATIONS"]


def remove_stamped_images(folder_path: str) -> None:
    """
    Recursively remove images with '_stamped' in their filename.
    """
    for root, _, files in os.walk(folder_path):
        for file in files:
            if "_stamped" in file.lower():
                file_path = os.path.join(root, file)
                try:
                    os.remove(file_path)
                    logging.info("Removed stamped image: %s", file_path)
                except Exception as e:
                    logging.error("Error removing file %s: %s", file_path, e)


def rotate_landscape_photos(folder_path: str) -> None:
    """
    Rotate landscape images (in directories without subdirectories) if width > height.
    """
    for root, _, files in os.walk(folder_path):
        # Check if folder is a leaf (has no subdirectories)
        if not [d for d in os.listdir(root) if os.path.isdir(os.path.join(root, d))]:
            for file in files:
                if file.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".tiff")):
                    file_path = os.path.join(root, file)
                    try:
                        with Image.open(file_path) as img:
                            if img.width > img.height:
                                rotated_img = img.rotate(-90, expand=True)
                                rotated_img.save(file_path)
                                logging.info("Rotated image: %s", file_path)
                    except Exception as e:
                        logging.error("Error processing file %s: %s", file_path, e)


def time_to_seconds(time_str: str) -> int:
    """
    Convert a time string 'HH:MM:SS' to seconds.
    """
    h, m, s = map(int, time_str.split(":"))
    return h * 3600 + m * 60 + s


def calculate_visual_difference(img1_path: str, img2_path: str) -> float:
    """
    Calculate the normalized visual difference between two images.
    """
    try:
        with Image.open(img1_path) as img1, Image.open(img2_path) as img2:
            img1 = img1.convert("RGB")
            img2 = img2.convert("RGB")
            diff = ImageChops.difference(img1, img2)
            diff_sum = sum(sum(pixel) for pixel in diff.getdata())
            max_diff = img1.width * img1.height * 255 * 3
            return diff_sum / max_diff
    except Exception as e:
        logging.error("Error comparing images %s and %s: %s", img1_path, img2_path, e)
        return 0.0


def generate_duration_timestamps(num_duration: int, t_before: int, duration: int, duration_images: list) -> list:
    """
    Generate timestamps for images in the DURATION interval.
    Time intervals are proportional to the visual difference between consecutive images.
    If there is only one image, simply return [t_before + duration].
    """
    t_end = t_before + duration
    if num_duration == 1:
        return [t_end]
    diffs = []
    for i in range(num_duration - 1):
        diff = calculate_visual_difference(duration_images[i], duration_images[i + 1])
        diffs.append(diff)
    total_diff = sum(diffs) if sum(diffs) != 0 else 1  # prevent division by zero
    intervals = [(d / total_diff) * duration for d in diffs]
    timestamps = []
    current = t_before
    for gap in intervals:
        current += gap
        timestamps.append(int(current))
    return timestamps


def generate_timestamps(num_images: int, start_time: int, duration_before: int, duration: int, image_files: list) -> list:
    """
    Generate a sorted list of timestamps for num_images.
    The first image gets time START_TIME ±30 min.
    The second image gets time = first image + DURATION_BEFORE_WORKS ±30 min.
    The remaining images are distributed in the remaining DURATION period with intervals
    proportional to the visual difference between adjacent images.
    """
    timestamps = []
    t1 = start_time + random.randint(-1800, 1800)
    timestamps.append(t1)
    t2 = t1 + duration_before + random.randint(-1800, 1800)
    timestamps.append(t2)
    if num_images == 2:
        return sorted(timestamps)
    diffs = []
    for i in range(1, num_images - 1):
        diff = calculate_visual_difference(image_files[i], image_files[i + 1])
        diffs.append(diff)
    total_diff = sum(diffs) or 1
    if total_diff == 0:
        intervals = [duration / (num_images - 2)] * (num_images - 2)
    else:
        intervals = [(d / total_diff) * duration for d in diffs]
    current_time = t2
    for gap in intervals:
        current_time += gap
        timestamps.append(int(current_time))
    timestamps.sort()
    return timestamps


def process_image(image_path: str, timestamp: int, date_obj: datetime, location: str) -> None:
    """
    Stamp an image with the date, time, and location information using cairo.
    """
    try:
        with Image.open(image_path) as img:
            img.thumbnail((2000, 2000))
            width, height = img.width, img.height
            output_image_path = f"{os.path.splitext(image_path)[0]}_stamped.png"

            # Create a cairo surface and context
            surface = cairo.ImageSurface(cairo.FORMAT_ARGB32, width, height)
            context = cairo.Context(surface)

            # Save the PIL image temporarily as PNG and load into cairo
            temp_path = output_image_path
            img.save(temp_path, format="PNG")
            input_surface = cairo.ImageSurface.create_from_png(temp_path)
            context.set_source_surface(input_surface, 0, 0)
            context.paint()

            # Set font and calculate positions
            font_size = int(height * 0.031)
            padding_x = width * 0.004
            padding_y = height * 0.004
            context.select_font_face("Noto Sans", cairo.FONT_SLANT_NORMAL, cairo.FONT_WEIGHT_BOLD)
            context.set_font_size(font_size)
            context.set_source_rgb(1, 1, 1)

            # Format date and time
            months = {
                "Jan": "янв", "Feb": "фев", "Mar": "мар", "Apr": "апр",
                "May": "мая", "Jun": "июн", "Jul": "июл", "Aug": "авг",
                "Sep": "сен", "Oct": "окт", "Nov": "ноя", "Dec": "дек"
            }
            formatted_date = date_obj.strftime("%d %b. %Y")
            for eng, rus in months.items():
                formatted_date = formatted_date.replace(eng, rus)
            formatted_date += " г."
            time_str = f"{formatted_date} {timestamp // 3600:02}:{(timestamp % 3600) // 60:02}:{timestamp % 60:02}"

            # Split location info into lines if needed
            lines = [time_str] + location.split("\n")
            text_y = padding_y + font_size

            # Draw each line on the image
            for line in lines:
                extents = context.text_extents(line)
                text_x = width - extents.width - padding_x
                context.move_to(text_x, text_y)
                context.show_text(line)
                text_y += extents.height + font_size * 0.28

            surface.write_to_png(output_image_path)
            logging.info("Saved stamped image: %s", output_image_path)
    except Exception as e:
        logging.error("Error processing image %s: %s", image_path, e)


def group_by_prefix(candidates: list) -> list:
    """
    Group file paths by prefix.
    The prefix is defined as the sequence of digits and underscores at the start of the filename.
    Trailing underscores are removed. From each group, the file with the longest name is selected.
    """
    groups = {}
    for path in candidates:
        filename = os.path.basename(path)
        match = re.match(r'^([\d_]+)', filename)
        if match:
            prefix = match.group(1).rstrip('_')
        else:
            prefix = filename
        if prefix in groups:
            if len(filename) > len(os.path.basename(groups[prefix])):
                groups[prefix] = path
        else:
            groups[prefix] = path
    return list(groups.values())


def process_incident(incident_folder: str) -> None:
    """
    Process an incident folder (e.g. "1", "2") with subfolders.
    Distribution logic:
      - The first subfolder: start photo = START_TIME ±30 min.
      - The second subfolder: photo with prefix "1_" gets time = start + DURATION_BEFORE_WORKS ±30 min,
        photos with prefix "2_" are assigned to DURATION group.
      - All subsequent subfolders: all photos are in the DURATION group.
    For each subfolder, group files by prefix and keep only one file per prefix.
    """
    subfolders = [
        os.path.join(incident_folder, d)
        for d in os.listdir(incident_folder)
        if os.path.isdir(os.path.join(incident_folder, d))
    ]
    if not subfolders:
        logging.info("No subfolders in incident %s", incident_folder)
        return

    subfolders.sort(key=lambda d: int(re.match(r'(\d+)', os.path.basename(d)).group(1)))
    start_images = []
    before_images = []
    duration_images = []

    for idx, folder in enumerate(subfolders):
        candidates = [
            os.path.join(folder, f)
            for f in os.listdir(folder)
            if f.lower().endswith((".jpg", ".jpeg", ".png")) and "_stamped" not in f.lower()
        ]
        candidates = group_by_prefix(candidates)
        if not candidates:
            continue
        if idx == 0:
            start_images.append(candidates[0])
        elif idx == 1:
            for file in candidates:
                basename = os.path.basename(file)
                if re.match(r'^1_', basename):
                    before_images.append(file)
                elif re.match(r'^2_', basename):
                    duration_images.append(file)
                else:
                    if not before_images:
                        before_images.append(file)
                    else:
                        duration_images.append(file)
        else:
            duration_images.extend(candidates)

    if not start_images or not before_images:
        logging.error("Incident %s: Missing start photo or before-work photo", incident_folder)
        return

    folder_date_match = re.search(r"\d{2}\.\d{2}\.\d{4}", os.path.basename(subfolders[0]))
    if not folder_date_match:
        logging.error("Incident %s: Could not extract date from folder name", incident_folder)
        return
    incident_date = datetime.strptime(folder_date_match.group(), "%d.%m.%Y")

    start_seconds = time_to_seconds(START_TIME)
    before_seconds = time_to_seconds(DURATION_BEFORE_WORKS)
    duration_seconds = time_to_seconds(DURATION)

    t_start = start_seconds + random.randint(-1800, 1800)
    t_before = t_start + before_seconds + random.randint(-1800, 1800)
    duration_timestamps = generate_duration_timestamps(len(duration_images), t_before, duration_seconds, duration_images)

    # Ensure the number of timestamps matches the number of images
    if len(duration_timestamps) < len(duration_images):
        duration_timestamps.append(t_before + duration_seconds)

    def adjust_timestamp(t: int, base_date: datetime) -> (datetime, int):
        current_date = base_date
        while t >= 86400:
            current_date += timedelta(days=1)
            t -= 86400
        return current_date, t

    current_date, t = adjust_timestamp(t_start, incident_date)
    process_image(start_images[0], t, current_date, random.choice(LOCATIONS))

    current_date, t = adjust_timestamp(t_before, incident_date)
    process_image(before_images[0], t, current_date, random.choice(LOCATIONS))

    for img_path, ts in zip(duration_images, duration_timestamps):
        current_date, t = adjust_timestamp(ts, incident_date)
        process_image(img_path, t, current_date, random.choice(LOCATIONS))


def process_folder(folder: str) -> None:
    """
    Process a leaf folder (without subdirectories):
      - Group candidate files by prefix.
      - Extract the date from the folder name.
      - Generate timestamps for all photos.
      - Stamp each photo.
    """
    candidates = [
        f for f in os.listdir(folder)
        if f.lower().endswith((".jpg", ".jpeg", ".png")) and "_stamped" not in f.lower()
    ]
    if not candidates:
        logging.info("No candidate images in folder: %s", folder)
        return

    groups = {}
    for file in candidates:
        match = re.match(r'^([\d_]+)', file)
        if match:
            prefix = match.group(1).rstrip('_')
        else:
            prefix = file
        if prefix in groups:
            if len(file) > len(groups[prefix]):
                groups[prefix] = file
        else:
            groups[prefix] = file
    grouped_files = [os.path.join(folder, groups[p]) for p in groups]
    num_files = len(grouped_files)
    logging.info("Folder '%s' - selected %d images", folder, num_files)

    folder_name = os.path.basename(folder)
    date_match = re.search(r"\d{2}\.\d{2}\.\d{4}", folder_name)
    if not date_match:
        logging.error("Could not extract date from folder name: %s", folder_name)
        return
    folder_date = datetime.strptime(date_match.group(), "%d.%m.%Y")

    start_seconds = time_to_seconds(START_TIME)
    duration_before_seconds = time_to_seconds(DURATION_BEFORE_WORKS)
    duration_seconds = time_to_seconds(DURATION)
    timestamps_seconds = generate_timestamps(num_files, start_seconds, duration_before_seconds, duration_seconds, grouped_files)

    timestamps = []
    for t in timestamps_seconds:
        current_date = folder_date
        while t >= 86400:
            current_date += timedelta(days=1)
            t -= 86400
        timestamps.append((current_date, t))

    for image_path, (date_obj, t) in zip(grouped_files, timestamps):
        process_image(image_path, t, date_obj, random.choice(LOCATIONS))


def main() -> None:
    """
    Main function:
      - Removes already stamped photos.
      - Rotates landscape photos.
      - Processes each incident (top-level folders like "1", "2", etc.).
        If the incident has subfolders, process_incident is used;
        if it's a leaf folder, process_folder is used.
    """
    remove_stamped_images(FOLDER_PATH)
    rotate_landscape_photos(FOLDER_PATH)
    for incident in os.listdir(FOLDER_PATH):
        incident_path = os.path.join(FOLDER_PATH, incident)
        if os.path.isdir(incident_path):
            subfolders = [
                d for d in os.listdir(incident_path)
                if os.path.isdir(os.path.join(incident_path, d))
            ]
            if subfolders:
                logging.info("Processing incident: %s", incident_path)
                process_incident(incident_path)
            else:
                logging.info("Processing leaf folder: %s", incident_path)
                process_folder(incident_path)


if __name__ == "__main__":
    main()

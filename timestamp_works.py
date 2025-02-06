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
    Генерация уникальных временных меток для изображений в интервале DURATION.
    Метки времени назначаются в алфавитном порядке имен файлов.
    """
    # Сортировка по имени файла
    duration_images.sort(key=lambda x: os.path.basename(x).lower())

    # Определение начала и конца интервала с разбросом ±30 минут
    t_start = t_before + random.randint(-1800, 1800)
    t_end = t_before + duration + random.randint(-1800, 1800)

    # Проверка, чтобы t_start был меньше t_end
    t_start, t_end = min(t_start, t_end), max(t_start, t_end)

    # Расчёт визуальных различий
    diffs = [calculate_visual_difference(duration_images[i], duration_images[i + 1]) for i in range(num_duration - 1)]
    total_diff = sum(diffs) or 1  # Предотвращение деления на ноль

    # Интервалы времени
    intervals = [(diff / total_diff) * (t_end - t_start) for diff in diffs]

    # Генерация временных меток
    timestamps = [t_start]
    for gap in intervals:
        next_time = max(timestamps[-1] + 1, timestamps[-1] + int(gap))  # Уникальность меток
        timestamps.append(next_time)

    # Ограничение последней метки
    timestamps[-1] = min(timestamps[-1], t_end)

    return timestamps


def generate_timestamps(num_images: int, start_time: int, duration_before: int, duration: int,
                        image_files: list) -> list:
    """
    Генерация временных меток на основе визуальной разницы между изображениями.
    Чем больше разница между фото, тем больше интервал времени между ними.
    """
    timestamps = []

    # Время для первого изображения (с небольшим рандомным сдвигом)
    t1 = start_time + random.randint(-1800, 1800)
    timestamps.append(t1)

    # Время для второго изображения
    t2 = t1 + duration_before + random.randint(-1800, 1800)
    timestamps.append(t2)

    if num_images == 2:
        return sorted(timestamps)

    # Вычисляем визуальную разницу между соседними изображениями
    diffs = []
    for i in range(1, num_images - 1):
        diff = calculate_visual_difference(image_files[i], image_files[i + 1])
        diffs.append(diff)

    # Сумма всех различий
    total_diff = sum(diffs) or 1  # Предотвращение деления на ноль

    # Расчёт интервалов на основе различий
    if total_diff == 0:
        # Если различий нет, распределяем равномерно
        intervals = [duration / (num_images - 2)] * (num_images - 2)
    else:
        # Пропорциональное распределение времени
        intervals = [(d / total_diff) * duration for d in diffs]

    current_time = t2
    for gap in intervals:
        # Генерация следующего временного интервала с учётом минимального шага
        current_time += max(1, int(gap))
        timestamps.append(current_time)

    return sorted(timestamps)


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
    subfolders = sorted(  # Сортировка папок по алфавиту
        [
            os.path.join(incident_folder, d)
            for d in os.listdir(incident_folder)
            if os.path.isdir(os.path.join(incident_folder, d))
        ],
        key=lambda d: os.path.basename(d).lower()
    )

    all_images = []  # Список для всех изображений в нужном порядке

    for folder in subfolders:
        candidates = sorted(  # Сортировка файлов по алфавиту
            [
                os.path.join(folder, f)
                for f in os.listdir(folder)
                if f.lower().endswith((".jpg", ".jpeg", ".png")) and "_stamped" not in f.lower()
            ],
            key=lambda x: os.path.basename(x).lower()
        )

        # Группировка файлов по префиксу
        grouped_files = group_by_prefix(candidates)

        # Сортировка сгруппированных файлов по имени
        grouped_files.sort(key=lambda x: os.path.basename(x).lower())

        # Добавляем в общий список
        all_images.extend(grouped_files)

    # Проверка даты для первой папки
    folder_date_match = re.search(r"\d{2}\.\d{2}\.\d{4}", os.path.basename(subfolders[0]))
    if not folder_date_match:
        logging.error("Incident %s: Could not extract date from folder name", incident_folder)
        return
    incident_date = datetime.strptime(folder_date_match.group(), "%d.%m.%Y")

    # Генерация временных меток
    start_seconds = time_to_seconds(START_TIME)
    before_seconds = time_to_seconds(DURATION_BEFORE_WORKS)
    duration_seconds = time_to_seconds(DURATION)

    timestamps = generate_timestamps(len(all_images), start_seconds, before_seconds, duration_seconds, all_images)

    def adjust_timestamp(t: int, base_date: datetime) -> (datetime, int):
        current_date = base_date
        while t >= 86400:
            current_date += timedelta(days=1)
            t -= 86400
        return current_date, t

    # Штамповка файлов в нужном порядке
    for img_path, ts in zip(all_images, timestamps):
        current_date, t = adjust_timestamp(ts, incident_date)
        process_image(img_path, t, current_date, random.choice(LOCATIONS))


def process_folder(folder: str) -> None:
    candidates = sorted(  # Сортировка файлов по алфавиту
        [
            os.path.join(folder, f)
            for f in os.listdir(folder)
            if f.lower().endswith((".jpg", ".jpeg", ".png")) and "_stamped" not in f.lower()
        ],
        key=lambda x: os.path.basename(x).lower()
    )

    if not candidates:
        logging.info("No candidate images in folder: %s", folder)
        return

    # Группировка файлов по префиксу
    grouped_files = group_by_prefix(candidates)

    # Сортировка сгруппированных файлов по имени
    grouped_files.sort(key=lambda x: os.path.basename(x).lower())

    folder_name = os.path.basename(folder)
    date_match = re.search(r"\d{2}\.\d{2}\.\d{4}", folder_name)
    if not date_match:
        logging.error("Could not extract date from folder name: %s", folder_name)
        return
    folder_date = datetime.strptime(date_match.group(), "%d.%m.%Y")

    start_seconds = time_to_seconds(START_TIME)
    duration_before_seconds = time_to_seconds(DURATION_BEFORE_WORKS)
    duration_seconds = time_to_seconds(DURATION)

    timestamps = generate_timestamps(len(grouped_files), start_seconds, duration_before_seconds, duration_seconds, grouped_files)

    def adjust_timestamp(t: int, base_date: datetime) -> (datetime, int):
        current_date = base_date
        while t >= 86400:
            current_date += timedelta(days=1)
            t -= 86400
        return current_date, t

    # Штамповка файлов
    for img_path, ts in zip(grouped_files, timestamps):
        current_date, t = adjust_timestamp(ts, folder_date)
        process_image(img_path, t, current_date, random.choice(LOCATIONS))


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

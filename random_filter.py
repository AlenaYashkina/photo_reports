import os
import sys
import json
import random
import logging
from typing import Any, Dict
import numpy as np
from PIL import Image, ImageEnhance, ImageFilter

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
logger = logging.getLogger(__name__)

CONFIG_FILE = "config.json"


def load_config(config_file: str) -> Dict[str, Any]:
    """Load configuration from a JSON file."""
    try:
        with open(config_file, "r", encoding="utf-8") as f:
            config = json.load(f)
        return config
    except Exception as e:
        logger.error("Failed to load configuration file: %s", e)
        sys.exit(1)


# Load configuration and get absolute folder path
config = load_config(CONFIG_FILE)
folder_path = os.path.abspath(config.get("folder_path", ""))
if not folder_path:
    logger.error("Folder path is not specified in config.")
    sys.exit(1)


def remove_old_filtered_images(folder: str) -> None:
    """
    Remove all files containing '_filtered' in their name from the specified folder and its subfolders.
    """
    for root, _, files in os.walk(folder):
        for file in files:
            if "_filtered" in file:
                file_path = os.path.join(root, file)
                try:
                    os.remove(file_path)
                    logger.info("Removed old file: %s", file_path)
                except Exception as e:
                    logger.error("Error removing file %s: %s", file_path, e)


def adjust_white_balance(img: Image.Image, red_factor: float, blue_factor: float) -> Image.Image:
    """
    Adjust white balance by modifying red and blue channels.
    Clamps the values to a maximum of 255.

    :param img: PIL.Image in RGB mode.
    :param red_factor: Multiplier for red channel.
    :param blue_factor: Multiplier for blue channel.
    :return: Image with adjusted white balance.
    """
    r, g, b = img.split()
    r = r.point(lambda i: min(255, i * red_factor))
    b = b.point(lambda i: min(255, i * blue_factor))
    return Image.merge("RGB", (r, g, b))


def adjust_hue(img: Image.Image, hue_factor: float) -> Image.Image:
    """
    Adjust the hue of the image by shifting the H channel.

    :param img: PIL.Image in RGB mode.
    :param hue_factor: Float in range [-0.05, 0.05] indicating the hue shift fraction.
    :return: Image with adjusted hue.
    """
    hsv = img.convert("HSV")
    h, s, v = hsv.split()
    np_h = np.array(h, dtype=np.uint8)
    offset = int(hue_factor * 255)
    np_h = (np_h.astype(int) + offset) % 256
    h = Image.fromarray(np_h.astype("uint8"), "L")
    hsv = Image.merge("HSV", (h, s, v))
    return hsv.convert("RGB")


def apply_random_effects(img: Image.Image) -> Image.Image:
    """
    Apply additional random effects to simulate different cameras and conditions:
    - Slight random rotation (-2 to 2 degrees)
    - Random Gaussian blur with 50% chance
    - Addition of Gaussian noise
    """
    # Apply slight random rotation
    angle = random.uniform(-2, 2)
    img = img.rotate(angle, resample=Image.BICUBIC, expand=False)

    # Randomly apply Gaussian blur
    if random.random() < 0.5:
        radius = random.uniform(0.2, 1.0)
        img = img.filter(ImageFilter.GaussianBlur(radius=radius))

    # Add Gaussian noise
    np_img = np.array(img).astype("float32")
    sigma = random.uniform(5, 15)
    noise = np.random.normal(0, sigma, np_img.shape)
    np_img = np.clip(np_img + noise, 0, 255).astype("uint8")
    return Image.fromarray(np_img)


def apply_filters(image: Image.Image) -> Image.Image:
    """
    Apply a series of filters to the image simulating varied shooting conditions:
    - White balance adjustment (red/blue channels)
    - Hue shift
    - Adjustments for brightness, contrast, color, and sharpness
    - Additional random effects (rotation, blur, noise)
    """
    img = image.convert("RGB")

    # Adjust white balance with random factors
    red_factor = random.uniform(0.98, 1.02)
    blue_factor = random.uniform(0.98, 1.02)
    img = adjust_white_balance(img, red_factor, blue_factor)

    # Apply a slight random hue shift
    hue_shift = random.uniform(-0.05, 0.05)
    img = adjust_hue(img, hue_shift)

    # Enhance brightness, contrast, color, and sharpness
    brightness = random.uniform(0.8, 0.9)
    contrast = random.uniform(1.05, 1.15)
    color = random.uniform(0.9, 1.15)
    sharpness = random.uniform(0.95, 1.2)
    img = ImageEnhance.Brightness(img).enhance(brightness)
    img = ImageEnhance.Contrast(img).enhance(contrast)
    img = ImageEnhance.Color(img).enhance(color)
    img = ImageEnhance.Sharpness(img).enhance(sharpness)

    # Apply additional random effects
    img = apply_random_effects(img)
    return img


def process_images(folder: str) -> None:
    """
    Process images in the given folder:
    - Remove previously filtered images.
    - For each subfolder, select the image with the longest filename.
    - Apply filters and save the modified image with a '_filtered' suffix.
    """
    if not os.path.exists(folder):
        logger.error("Folder '%s' does not exist.", folder)
        return

    remove_old_filtered_images(folder)

    supported_extensions = (".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".webp")
    for root, _, files in os.walk(folder):
        image_files = [f for f in files if f.lower().endswith(supported_extensions)]
        if not image_files:
            continue

        # Select image with the longest filename
        longest_name_file = max(image_files, key=len)
        image_path = os.path.join(root, longest_name_file)

        try:
            with Image.open(image_path) as img:
                filtered_img = apply_filters(img)
                base, ext = os.path.splitext(image_path)
                new_image_path = f"{base}_filtered{ext}"
                filtered_img.save(new_image_path)
                logger.info("Processed and saved: %s", new_image_path)
        except Exception as e:
            logger.error("Error processing image %s: %s", image_path, e)


if __name__ == "__main__":
    process_images(folder_path)

# Photo Processing & Report Generation Pipeline

This project provides a complete workflow for downloading, processing, and reporting on photos from **Telegram channels**. The pipeline:

- Downloads images from specified **Telegram topics**
- Applies **random filters** to simulate varied shooting conditions
- Stamps images with **timestamps** and **location data**
- Creates **PowerPoint presentations** (with PDF conversion) as reports

Two variants of the report and timestamping modules are provided:
- **Daily Reports ("days")**
- **Work Reports ("works")**

---

## File Structure

### `config.json` — *Configuration file*
- Contains API credentials, phone number, chat/channel IDs, topic IDs, folder paths, output report paths, titles, time settings, and locations.
- All scripts read from this file — adjust its parameters to match your environment.

---

### `downloader.py` — *Telegram Photo Downloader*
- Connects to Telegram using **Telethon**.
- Downloads photos from specified topics (via reply-to message IDs).
- Saves images into subfolders named after the chat/topic.
- Maintains progress to avoid duplicate downloads.

---

### `random_filter.py` — *Random Image Filtering*
- Applies random filters using **Pillow** & **NumPy**:
  - White balance, hue shift, brightness, contrast, saturation, sharpness
  - Extra effects: rotations, Gaussian blur, noise
- Processes images based on the paths in `config.json`.

---

### `report_creator_days.py` — *Report Creator for Daily Reports*
- Generates PowerPoint reports using **python-pptx**.
- Organizes photos in day-based folders.
- Adds a title slide and arranges photos in a **3-column layout**.
- Converts PPTX to PDF (requires **PowerPoint** installed on Windows).

---

### `report_creator_works.py` — *Report Creator for Work Reports*
- Similar to `report_creator_days.py`, but with:
  - Different slide layouts & photo placements
  - Dynamic font sizing for group labels
- Outputs both **PPTX** & **PDF** formats.

---

### `timestamp_days.py` — *Image Timestamp Stamping (Daily Reports)*
- Adds date, time, and location stamps using **cairo** & **Pillow**.
- Formats include:
  - Russian month abbreviations
  - Randomized time offsets
  - Random locations from a provided list

---

### `timestamp_works.py` — *Advanced Timestamp Stamping (Work Reports)*
- More advanced logic:
  - Groups images by filename prefixes
  - Handles multi-phase work folders
  - Proportional timestamp intervals based on visual differences
  - Supports date rollover if timestamps exceed 24 hours

---

## Dependencies

Ensure you have **Python 3.7+**. Install required packages:

```bash
pip install -r requirements.txt
```

### Notes:
- PDF conversion requires **Windows** with **Microsoft PowerPoint** installed (`win32com.client` dependency).
- For **cairo**, you may need to install system-level libraries (check `pycairo` installation instructions).

---

## Setup & Configuration

1. **Configure `config.json`:**
   - Set **Telegram API credentials** (`api_id`, `api_hash`)
   - Add your **phone number** and **chat/channel IDs**
   - Specify **topics** (message IDs) for photo downloads
   - Define folder paths:
     - `folder_path` (for images)
     - `output_path` (for reports)
   - Set **title content** for reports
   - Adjust **time parameters**:
     - `START_TIME`: e.g., `18:40:09`
     - `DURATION_BEFORE_WORKS`, `DURATION`: e.g., `01:30:00`
   - Add **locations** for timestamping

2. **Environment Setup:**
   - Windows OS recommended (for PowerPoint-to-PDF conversion)
   - Ensure all external dependencies are installed correctly

---

## Workflow

1. **Download Photos:**
   ```bash
   python downloader.py
   ```

2. **Apply Random Filters:**
   ```bash
   python random_filter.py
   ```

3. **Stamp Images:**
   - For daily reports:
     ```bash
     python timestamp_days.py
     ```
   - For work reports:
     ```bash
     python timestamp_works.py
     ```

4. **Generate Reports:**
   - Daily reports:
     ```bash
     python report_creator_days.py
     ```
   - Work reports:
     ```bash
     python report_creator_works.py
     ```

> **Ensure scripts run in the correct order for the workflow to function properly.**

---

## Customization & Troubleshooting

- **Layout Customization:**
  - Adjust dimensions, margins, and fonts directly in `report_creator_days.py` or `report_creator_works.py`.

- **Error Logging:**
  - All scripts include logging. Check the console for debugging information.

- **Platform Limitations:**
  - PDF conversion via `win32com.client` works **only on Windows**.

- **Supported Formats:**
  - JPG, PNG, BMP, TIFF, and other common formats.

---

## Conclusion

This pipeline **automates** the process of:

- Gathering images from Telegram
- Applying creative filters & stamps
- Compiling professional reports in **PPTX** & **PDF** formats

> **Adjust configurations as needed to fit your workflow.**  
> **Feel free to modify and extend the project.**

**Happy processing!**
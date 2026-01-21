# 🎬 Excel-Driven Video Generator (Tkinter + MoviePy)

A desktop GUI application built with Python and Tkinter that automatically generates vertical marketing videos by combining a base video, Excel product data, a brand logo, and custom promotional text.

This tool is especially useful for jewellery brands, catalog teams, and social media marketing workflows.

---

## ✨ Features

- Simple and clean Tkinter-based GUI
- Select base video, Excel file, and logo image
- Reads Excel headers from Row 4
- Automatically matches product data using stock code
- Custom text input (maximum 50 characters)
- Beautiful bottom information bar with product details
- Vertical video output (1080 × 1920)
- Live progress percentage while generating video
- Output videos saved automatically in an `output` folder
- No audio for clean marketing visuals

---

## 🧰 Tech Stack

- Python 3.9+
- Tkinter (GUI)
- Pillow (Image & text rendering)
- MoviePy (Video processing)
- Pandas (Excel handling)
- NumPy (Image conversion)

---

## 📂 Project Structure
video-generator/
│
├── main.py
├── output/
├── README.md
└── requirements.txt

---

## 📑 Excel File Requirements

- Excel headers must be on **Row 4**
- Required column names (case-insensitive):

| Column Name |
|------------|
| stock code |
| purity |
| gross wt |
| total dia wt |
| total ps wt |
| total clr stn wt |

The stock code in Excel must match the video filename.

Example:
Video file: YS18001.mp4
Excel code: YS18001

Spaces, dashes, dots, and special characters are automatically normalized.

---

## 🎥 Video Output Specifications

- Canvas Size: 1080 × 1920
- Video Size: 900 × 1700
- Bottom Bar Height: 220 px
- Frame Rate: 24 FPS
- Output Format: MP4
- Audio: Disabled

---

## 🖱️ How to Use

1. Run the application
```bash
python main.py
Select:

Video file (.mp4)

Excel file (.xlsx)

Logo image (.png / .jpg)

Enter custom text (maximum 50 characters)

Click Generate Video

Generated video will be saved in:
output/STOCKCODE_YS18.mp4
python -m venv venv
source venv/bin/activate
# Windows: venv\Scripts\activate
pip install -r requirements.txt
pillow
pandas
moviepy
numpy
proglog
openpyxl

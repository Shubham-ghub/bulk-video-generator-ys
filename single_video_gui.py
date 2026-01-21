import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import moviepy.editor as mp
import numpy as np
import os
import textwrap
import re
from proglog import ProgressBarLogger

# ================= CONFIG =================
CANVAS_W, CANVAS_H = 1080, 1920
VIDEO_W, VIDEO_H = 900, 1700
BG_COLOR = (239, 235, 226)

BOTTOM_BAR_H = 220
FONT_COLOR = "#545454"

LOGO_TOP = 150
LOGO_RIGHT = 150
LOGO_WIDTH = 150

MAX_TEXT = 50
# =========================================


class TkPercentLogger(ProgressBarLogger):
    def __init__(self, label):
        super().__init__()
        self.label = label

    def callback(self, **changes):
        if "progress" in changes:
            pct = int(changes["progress"] * 100)
            self.label.config(text=f"Processing: {pct}%")
            self.label.update_idletasks()


class VideoGeneratorApp:
    def __init__(self, root):
        self.root = root
        root.title("Video Generator")
        root.geometry("380x480")

        self.video_path = None
        self.logo_path = None
        self.excel_path = None
        self.df = None

        self.build_ui()

    # ---------------- UI ----------------
    def build_ui(self):
        f = tk.Frame(self.root, padx=10, pady=10)
        f.pack()

        tk.Button(f, text="Select Video", width=30, command=self.pick_video).pack(pady=4)
        self.video_lbl = tk.Label(f, text="No video selected", fg="gray", wraplength=300)
        self.video_lbl.pack()

        tk.Button(f, text="Select Excel", width=30, command=self.pick_excel).pack(pady=4)
        self.excel_lbl = tk.Label(f, text="No excel selected", fg="gray", wraplength=300)
        self.excel_lbl.pack()

        tk.Button(f, text="Select Logo", width=30, command=self.pick_logo).pack(pady=4)
        self.logo_lbl = tk.Label(f, text="No logo selected", fg="gray", wraplength=300)
        self.logo_lbl.pack()

        tk.Label(f, text="Custom Text (max 50 chars)").pack(pady=(12, 4))
        self.custom_text = tk.StringVar()
        self.custom_text.trace_add("write", self.limit_text)
        tk.Entry(f, textvariable=self.custom_text, width=38).pack()

        self.progress_lbl = tk.Label(f, text="Processing: 0%", fg="blue")
        self.progress_lbl.pack(pady=10)

        tk.Button(
            f,
            text="Generate Video",
            bg="#4CAF50",
            fg="white",
            command=self.generate
        ).pack(pady=10)

    # ---------------- HELPERS ----------------
    def limit_text(self, *a):
        if len(self.custom_text.get()) > MAX_TEXT:
            self.custom_text.set(self.custom_text.get()[:MAX_TEXT])

    def normalize_code(self, s):
        """Remove spaces, dashes, dots, slashes – keep A-Z 0-9 only"""
        return re.sub(r"[^A-Z0-9]", "", str(s).upper())

    def fmt(self, val, dec=False):
        try:
            val = float(val)
            return f"{val:.2f}" if dec else f"{int(val)}"
        except:
            return ""

    # ---------------- FILE PICKERS ----------------
    def pick_video(self):
        self.video_path = filedialog.askopenfilename(filetypes=[("Video", "*.mp4")])
        if self.video_path:
            self.video_lbl.config(text=os.path.basename(self.video_path), fg="green")

    def pick_excel(self):
        self.excel_path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not self.excel_path:
            return

        # 🔹 HEADERS IN ROW 4
        df = pd.read_excel(self.excel_path, header=3)

        # 🔹 NORMALIZE COLUMN NAMES
        df.columns = (
            df.columns.astype(str)
            .str.strip()
            .str.lower()
            .str.replace(".", "", regex=False)
        )

        required = [
            "stock code",
            "purity",
            "gross wt",
            "total dia wt",
            "total ps wt",
            "total clr stn wt"
        ]

        missing = [c for c in required if c not in df.columns]
        if missing:
            messagebox.showerror("Excel Error", f"Missing columns:\n{', '.join(missing)}")
            return

        # 🔹 NORMALIZE STOCK CODES
        df["stock_code_norm"] = df["stock code"].apply(self.normalize_code)

        self.df = df
        self.excel_lbl.config(text=os.path.basename(self.excel_path), fg="green")

    def pick_logo(self):
        self.logo_path = filedialog.askopenfilename(
            filetypes=[("Image", "*.png *.jpg *.jpeg")]
        )
        if self.logo_path:
            self.logo_lbl.config(text=os.path.basename(self.logo_path), fg="green")

    # ---------------- EXCEL MATCH ----------------
    def get_excel_row(self, video_name):
        if self.df is None:
            return None

        code = self.normalize_code(video_name)

        row = self.df[self.df["stock_code_norm"] == code]
        if row.empty:
            return None

        r = row.iloc[0]
        return [
            ("Item", r["stock code"]),
            ("Kt", r["purity"]),
            ("Gold", self.fmt(r["gross wt"], True)),
            ("Dia", self.fmt(r["total dia wt"], True)),
            ("PS", self.fmt(r["total ps wt"], True)),
            ("CS", self.fmt(r["total clr stn wt"], True)),
        ]

    # ---------------- BOTTOM BAR ----------------
    def build_bottom_bar_clip(self, data, duration):
        img = Image.new("RGBA", (CANVAS_W, BOTTOM_BAR_H), BG_COLOR)
        d = ImageDraw.Draw(img)

        font = ImageFont.truetype("arial.ttf", 20)
        text_font = ImageFont.truetype("arial.ttf", 30)

        wrapped = "\n".join(textwrap.wrap(self.custom_text.get(), width=22)[:2])
        d.multiline_text(
            (CANVAS_W // 4, BOTTOM_BAR_H // 2),
            wrapped,
            fill=FONT_COLOR,
            anchor="mm",
            font=text_font,
            align="center",
            spacing=10
        )

        d.line(
            (CANVAS_W // 2, 20, CANVAS_W // 2, BOTTOM_BAR_H - 20),
            fill=FONT_COLOR,
            width=2
        )

        col_widths = [100, 60, 80, 80, 80, 100]
        table_w = sum(col_widths)
        start_x = CANVAS_W // 2 + ((CANVAS_W // 2 - table_w) // 2)
        start_y = (BOTTOM_BAR_H - 60) // 2

        x = start_x
        for (k, v), w in zip(data, col_widths):
            d.rectangle((x, start_y, x + w, start_y + 60),
                        outline=FONT_COLOR, width=2)
            d.text(
                (x + w // 2, start_y + 30),
                f"{k}\n{v}",
                fill=FONT_COLOR,
                anchor="mm",
                font=font,
                align="center"
            )
            x += w

        return mp.ImageClip(np.array(img)).set_duration(duration)

    # ---------------- GENERATE ----------------
    def generate(self):
        if not all([self.video_path, self.excel_path, self.logo_path]):
            messagebox.showerror("Missing", "All inputs required")
            return

        stock_code = os.path.splitext(os.path.basename(self.video_path))[0]
        excel_data = self.get_excel_row(stock_code)

        if not excel_data:
            messagebox.showerror("Excel Error", f"No matching stock code:\n{stock_code}")
            return

        output_dir = os.path.join(os.path.dirname(self.video_path), "output")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, f"{stock_code}_YS18.mp4")

        self.progress_lbl.config(text="Processing: 0%")
        logger = TkPercentLogger(self.progress_lbl)

        video = mp.VideoFileClip(self.video_path).resize((VIDEO_W, VIDEO_H))

        logo = (mp.ImageClip(self.logo_path)
                .resize(width=LOGO_WIDTH)
                .set_position((CANVAS_W - LOGO_WIDTH - LOGO_RIGHT, LOGO_TOP))
                .set_duration(video.duration))

        bar = (self.build_bottom_bar_clip(excel_data, video.duration)
               .set_position(("center", CANVAS_H - BOTTOM_BAR_H)))

        final = mp.CompositeVideoClip(
            [video.set_position(("center", 100)), logo, bar],
            size=(CANVAS_W, CANVAS_H),
            bg_color=BG_COLOR
        )

        final.write_videofile(
            output_path,
            fps=24,
            logger=logger,
            audio=False
        )

        self.progress_lbl.config(text="Processing: 100%")
        messagebox.showinfo("Done", f"Saved in:\n{output_path}")


if __name__ == "__main__":
    root = tk.Tk()
    VideoGeneratorApp(root)
    root.mainloop()

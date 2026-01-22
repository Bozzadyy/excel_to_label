import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

# ==============================
# PDF EXPORT SETTINGS
# ==============================

FONT_SIZE = 12

LEFT_MARGIN = 36 * mm
RIGHT_MARGIN = 29 * mm
TOP_MARGIN = 18 * mm
BOTTOM_MARGIN = 16 * mm

ROWS = 12
COLS = 2
COL_GAP = 3 * mm
ROW_GAP = 3 * mm

# ==============================
# GUI APPLICATION
# ==============================

class ExcelToPDFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to PDF Exporter")
        self.root.geometry("620x420")

        self.excel_file_path = None

        # Title
        tk.Label(root, text="Excel → PDF Export Tool", font=("Segoe UI", 16, "bold")).pack(pady=10)

        # File selection button
        tk.Button(root, text="Choose Excel File", command=self.choose_excel, width=20).pack()

        # Display selected file
        self.file_label = tk.Label(root, text="No file selected", fg="gray")
        self.file_label.pack(pady=5)

        # Export button
        tk.Button(root, text="Export to PDF", command=self.export_to_pdf, width=20).pack(pady=10)

        # Log area
        tk.Label(root, text="Log Output:").pack()
        self.log_area = scrolledtext.ScrolledText(root, width=70, height=12, font=("Consolas", 10))
        self.log_area.pack(padx=10, pady=5)

        # Reset button
        tk.Button(root, text="Reset", command=self.reset_app, width=15).pack(pady=10)

    def log(self, msg):
        self.log_area.insert(tk.END, msg + "\n")
        self.log_area.see(tk.END)

    def choose_excel(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
        )
        if file_path:
            self.excel_file_path = file_path
            self.file_label.config(text=file_path, fg="black")
            self.log(f"Selected file: {file_path}")

    def export_to_pdf(self):
        if not self.excel_file_path:
            messagebox.showerror("Error", "Please choose an Excel file first.")
            return

        try:
            save_path = filedialog.asksaveasfilename(
                title="Save PDF As",
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")]
            )

            if not save_path:
                self.log("Export cancelled by user.")
                return

            self.log("Loading Excel...")
            df = pd.read_excel(self.excel_file_path)
            df.columns = df.columns.str.strip()

            rows = [f"{r['Name']} {r['With']}" for _, r in df.iterrows()]

            self.log("Registering fonts...")
            # BASE_DIR = os.path.dirname(os.path.abspath(__file__))

            if getattr(sys, 'frozen', False):
                BASE_DIR = os.path.dirname(sys.executable)
            else:
                BASE_DIR = os.path.dirname(os.path.abspath(__file__))

            REGULAR_FONT = os.path.join(BASE_DIR, "Saysettha-Regular.ttf")
            BOLD_FONT = os.path.join(BASE_DIR, "Saysettha-Bold.ttf")

            pdfmetrics.registerFont(TTFont("SaysetthaReg", REGULAR_FONT))
            pdfmetrics.registerFont(TTFont("SaysetthaBold", BOLD_FONT))

            PAGE_WIDTH, PAGE_HEIGHT = landscape(A4)

            usable_width = PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN
            usable_height = PAGE_HEIGHT - TOP_MARGIN - BOTTOM_MARGIN

            col_width = (usable_width - COL_GAP) / COLS
            row_height = (usable_height - (ROWS - 1) * ROW_GAP) / ROWS

            def draw_page(c, page_data):
                c.setFont("SaysetthaBold", FONT_SIZE)

                for col in range(COLS):
                    for row in range(ROWS):
                        idx = col * ROWS + row
                        if idx >= len(page_data):
                            return

                        text = page_data[idx]

                        x = LEFT_MARGIN + col * (col_width + COL_GAP)
                        y_top = PAGE_HEIGHT - TOP_MARGIN - row * (row_height + ROW_GAP)

                        text_width = c.stringWidth(text, "SaysetthaBold", FONT_SIZE)
                        center_x = x + (col_width - text_width) / 2
                        baseline_y = y_top - row_height + (row_height * 0.60)

                        c.drawString(center_x, baseline_y, text)

            self.log("Generating PDF...")
            c = canvas.Canvas(save_path, pagesize=landscape(A4))

            for i in range(0, len(rows), ROWS * COLS):
                page_data = rows[i:i + ROWS * COLS]
                draw_page(c, page_data)
                c.showPage()

            c.save()

            self.log(f"PDF saved to: {save_path}")
            messagebox.showinfo("Success", f"PDF successfully created!\nSaved at:\n{save_path}")

        except Exception as e:
            self.log(f"[ERROR] {str(e)}")
            messagebox.showerror("Export Failed", str(e))

    def reset_app(self):
        self.excel_file_path = None
        self.file_label.config(text="No file selected", fg="gray")
        self.log_area.delete(1.0, tk.END)
        self.log("Application reset.")

# ==============================
# RUN APPLICATION
# ==============================

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToPDFApp(root)
    root.mainloop()

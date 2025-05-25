import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import os

# === Helper Functions (unchanged logic) ===
def is_number(text):
    return text.replace(".", "").replace(",", "").isdigit()

def merge_and_reverse(text):
    if isinstance(text, list):
        text = " ".join(text)
    if isinstance(text, str):
        if is_number(text):
            return text
        lines = text.split("\n")
        reversed_lines = " ".join(lines[::-1])
        if len(reversed_lines.split()) > 0:
            return reversed_lines[::-1]
        return reversed_lines
    return text

char_map = {
    "\uFE8D": "\u0627",  # ا
    "\uFE8E": "\u0627",  # ا
    "\uFE90": "\u0628",  # ب
    "\uFE8F": "\u0628",  # ب
    "\uFE91": "\u0628",  # ب
    "\uFE92": "\u0628",  # ب
    "\uFE95": "\u062A",  # ت
    "\uFE96": "\u062A",  # ت
    "\uFE97": "\u062A",  # ت
    "\uFE98": "\u062A",  # ت
    "\uFE99": "\u062B",  # ث
    "\uFE9A": "\u062B",  # ث
    "\uFE9B": "\u062B",  # ث
    "\uFE9C": "\u062B",  # ث
    "\uFE9D": "\u062C",  # ج
    "\uFE9E": "\u062C",  # ج
    "\uFE9F": "\u062C",  # ج
    "\uFEA0": "\u062C",  # ج
    "\uFEA1": "\u062D",  # ح
    "\uFEA2": "\u062D",  # ح
    "\uFEA3": "\u062D",  # ح
    "\uFEA4": "\u062D",  # ح  
    "\uFEA5": "\u062E",  # خ
    "\uFEA6": "\u062E",  # خ
    "\uFEA7": "\u062E",  # خ
    "\uFEA8": "\u062E",  # خ
    "\uFEA9": "\u062F",  # د
    "\uFEAA": "\u062F",  # د
    "\uFEAB": "\u0630",  # ذ
    "\uFEAC": "\u0630",  # ذ
    "\uFEB1": "\u0631",  # ر
    "\uFEAE": "\u0631",  # ر
    "\uFEAD": "\u0631",  # ر
    "\uFEAF": "\u0632",  # ز
    "\uFEB0": "\u0632",  # ز
    "\uFEB1": "\u0633",  # س
    "\uFEB2": "\u0633",  # س
    "\uFEB3": "\u0633",  # س
    "\uFEB4": "\u0633",  # س
    "\uFEB5": "\u0634",  # ش
    "\uFEB6": "\u0634",  # ش
    "\uFEB7": "\u0634",  # ش
    "\uFEB8": "\u0634",  # ش
    "\uFEB9": "\u0635",  # ص
    "\uFEBA": "\u0635",  # ص
    "\uFEBB": "\u0635",  # ص
    "\uFEBC": "\u0635",  # ص
    "\uFEBD": "\u0636",  # ض
    "\uFEBE": "\u0636",  # ض
    "\uFEBF": "\u0636",  # ض
    "\uFEC0": "\u0636",  # ض
    "\uFEC1": "\u0637",  # ط
    "\uFEC2": "\u0637",  # ط
    "\uFEC3": "\u0637",  # ط
    "\uFEC4": "\u0637",  # ط
    "\uFEC5": "\u0638",  # ظ
    "\uFEC6": "\u0638",  # ظ
    "\uFEC7": "\u0638",  # ظ
    "\uFEC8": "\u0638",  # ظ
    "\uFEC9": "\u0639",  # ع
    "\uFECA": "\u0639",  # ع
    "\uFECB": "\u0639",  # ع
    "\uFECC": "\u0639",  # ع
    "\uFECD": "\u063A",  # غ
    "\uFECE": "\u063A",  # غ
    "\uFECF": "\u063A",  # غ
    "\uFED0": "\u063A",  # غ
    "\uFED1": "\u0641",  # ف
    "\uFED2": "\u0641",  # ف
    "\uFED3": "\u0641",  # ف
    "\uFED4": "\u0641",  # ف
    "\uFED5": "\u0642",  # ق
    "\uFED6": "\u0642",  # ق
    "\uFED7": "\u0642",  # ق
    "\uFED8": "\u0642",  # ق
    "\uFED9": "\u0643",  # ك
    "\uFEDA": "\u0643",  # ك
    "\uFEDB": "\u0643",  # ك
    "\uFEDC": "\u0643",  # ك
    "\uFEDD": "\u0644",  # ل
    "\uFEDE": "\u0644",  # ل
    "\uFEDF": "\u0644",  # ل
    "\uFEE0": "\u0644",  # ل
    "\uFEE1": "\u0645",  # م
    "\uFEE2": "\u0645",  # م
    "\uFEE3": "\u0645",  # م
    "\uFEE4": "\u0645",  # م
    "\uFEE5": "\u0646",  # ن
    "\uFEE6": "\u0646",  # ن
    "\uFEE7": "\u0646",  # ن
    "\uFEE8": "\u0646",  # ن
    "\uFEE9": "\u0647",  # ه
    "\uFEEA": "\u0647",  # ه
    "\uFEEB": "\u0647",  # ه
    "\uFEEC": "\u0647",  # ه
    "\uFEED": "\u0648",  # و
    "\uFEEE": "\u0648",  # و
    "\uFEF1": "\u064A",  # ي
    "\uFEF2": "\u064A",  # ي
    "\uFEF3": "\u064A",  # ي
    "\uFEF4": "\u064A",  # ي
    "\uFE93": "\u0629",  # ة
    "\uFE94": "\u0629",  # ة
    "\uFE85": "\u0624",  # ؤ
    "\uFE86": "\u0624",  # ؤ
    "\uFE89": "\u0626",  # ئ
    "\uFE8A": "\u0626",  # ئ
    "\uFE8B": "\u0626",  # ئ
    "\uFE8C": "\u0626",  # ئ
    "\uFE80": "\u0621",  # ء
    "\uFE83": "\u0623",  # أ
    "\uFE84": "\u0623",  # أ
    "\uFEEF": "\u0649",  # ى
    "\uFEF0": "\u0649",  # ى
    "\uFEFC": "\u0644" + "\u0627",  # لا
    "\uFEFB": "\u0644" + "\u0627",  # لا
}


def convert_text(text, char_map):
    if not isinstance(text, str):
        return text
    return ''.join(char_map.get(char, char) for char in text)

def convert_excel_file(input_file, output_file, char_map):
    df = pd.read_excel(input_file)
    df_cleaned = df.applymap(lambda cell: convert_text(cell, char_map))
    df_cleaned.to_excel(output_file, index=False)

def fix_words_in_cell(cell, words_to_fix):
    cell = str(cell)
    words_in_cell = cell.split()
    updated_words = []
    for word in words_in_cell:
        updated_word = word
        if word in words_to_fix:
            updated_word = updated_word[:-1] + "ء"
        updated_words.append(updated_word)
    return " ".join(updated_words)

words_to_fix = [
    "علار", "شيمار", "اسمار", "نجلار", "ولار", "الار", "وفار",
    "ضيار", "صفار", "هنار", "دعار", "رجار", "نجلار", "لميار",
    "حسنار", "سنار", "ثنار"
]

# === GUI Action ===
def select_pdfs():
    files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
    pdf_paths_var.set("; ".join(files))

def select_output_folder():
    folder = filedialog.askdirectory()
    output_folder_var.set(folder)

def run_batch_conversion():
    pdf_paths = pdf_paths_var.get().split("; ")
    output_folder = output_folder_var.get()

    if not pdf_paths or not os.path.isfile(pdf_paths[0]):
        messagebox.showerror("Error", "Please select at least one valid PDF file.")
        return
    if not output_folder:
        messagebox.showerror("Error", "Please select an output folder.")
        return

    for pdf_path in pdf_paths:
        try:
            status_var.set(f"Processing: {os.path.basename(pdf_path)}")
            root.update()

            all_tables = []
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if table:
                        df = pd.DataFrame(table)
                        df = df.applymap(merge_and_reverse)
                        all_tables.append(df)

            if not all_tables:
                continue

            tmp_excel = pdf_path[:-4] + "_tmp.xlsx"
            final_df = pd.concat(all_tables, ignore_index=True)
            final_df.to_excel(tmp_excel, index=False, header=False)

            cleaned_excel = tmp_excel[:-9] + "_cleaned.xlsx"
            convert_excel_file(tmp_excel, cleaned_excel, char_map)

            df = pd.read_excel(cleaned_excel, header=None)
            for col in df.columns:
                for row in df.index:
                    df.at[row, col] = fix_words_in_cell(df.at[row, col], words_to_fix)

            final_filename = os.path.splitext(os.path.basename(pdf_path))[0] + "_Final.xlsx"
            final_path = os.path.join(output_folder, final_filename)
            df.to_excel(final_path, index=False, header=False)

            os.remove(tmp_excel)
            os.remove(cleaned_excel)
        except Exception as e:
            messagebox.showerror("Error", f"Error processing {pdf_path}:\n{str(e)}")

    status_var.set("✅ All PDFs processed.")
    messagebox.showinfo("Done", "All files converted and cleaned successfully.")

# === GUI Layout ===
root = tk.Tk()
root.title("Batch Arabic PDF to Excel Converter")

pdf_paths_var = tk.StringVar()
output_folder_var = tk.StringVar()
status_var = tk.StringVar(value="Select PDFs and output folder to begin...")

tk.Label(root, text="PDF Files:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
tk.Entry(root, textvariable=pdf_paths_var, width=60).grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="Browse PDFs", command=select_pdfs).grid(row=0, column=2, padx=5, pady=5)

tk.Label(root, text="Output Folder:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
tk.Entry(root, textvariable=output_folder_var, width=60).grid(row=1, column=1, padx=5, pady=5)
tk.Button(root, text="Select Folder", command=select_output_folder).grid(row=1, column=2, padx=5, pady=5)

tk.Button(root, text="Run Conversion", command=run_batch_conversion).grid(row=2, column=1, pady=10)
tk.Label(root, textvariable=status_var, fg="blue").grid(row=3, column=0, columnspan=3, pady=5)

root.mainloop()

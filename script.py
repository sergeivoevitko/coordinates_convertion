import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pyproj
from tkinter import ttk


class SuperDuperConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Super Duper Converter")
        self.root.geometry("500x400")
        self.file_path = ""
        self.df = None

        # Title Label
        self.title_label = tk.Label(root, text="Super Duper Converter", font=("Arial", 16, "bold"))
        self.title_label.pack(pady=10)

        # Buttons for conversion selection
        self.btn_wgs_to_tm = tk.Button(root, text="WGS84 to Israel TM", command=self.wgs_to_tm_interface)
        self.btn_wgs_to_tm.pack(pady=10)

        self.btn_tm_to_wgs = tk.Button(root, text="Israel TM to WGS84", command=self.tm_to_wgs_interface)
        self.btn_tm_to_wgs.pack(pady=10)

    def wgs_to_tm_interface(self):
        self.clear_interface()
        self.load_excel()
        self.create_conversion_interface("Select Latitude Column:", "Select Longitude Column:", self.convert_wgs_to_tm)

    def tm_to_wgs_interface(self):
        self.clear_interface()
        self.load_excel()
        self.create_conversion_interface("Select X Column:", "Select Y Column:", self.convert_tm_to_wgs)

    def clear_interface(self):
        for widget in self.root.winfo_children():
            widget.destroy()
        self.title_label = tk.Label(self.root, text="Super Duper Converter", font=("Arial", 16, "bold"))
        self.title_label.pack(pady=10)

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if file_path:
            self.file_path = file_path
            self.df = pd.read_excel(file_path)

            if self.df.empty:
                messagebox.showerror("Error", "The selected file is empty.")
                return

            messagebox.showinfo("Success", "File loaded successfully! Select columns.")

    def create_conversion_interface(self, label1, label2, conversion_function):
        """Creates the dropdown interface for conversion."""
        self.col_label1 = tk.Label(self.root, text=label1)
        self.col_label1.pack()
        self.col_dropdown1 = ttk.Combobox(self.root, state="readonly")
        self.col_dropdown1.pack()

        self.col_label2 = tk.Label(self.root, text=label2)
        self.col_label2.pack()
        self.col_dropdown2 = ttk.Combobox(self.root, state="readonly")
        self.col_dropdown2.pack()

        if self.df is not None:
            columns = self.df.columns.tolist()
            self.col_dropdown1["values"] = columns
            self.col_dropdown2["values"] = columns

        self.btn_convert = tk.Button(self.root, text="Convert", command=conversion_function)
        self.btn_convert.pack(pady=10)

    def convert_wgs_to_tm(self):
        if self.df is None:
            messagebox.showerror("Error", "Please upload an Excel file first.")
            return

        lat_col = self.col_dropdown1.get()
        lon_col = self.col_dropdown2.get()

        if not lat_col or not lon_col:
            messagebox.showerror("Error", "Please select latitude and longitude columns.")
            return

        try:
            transformer = pyproj.Transformer.from_crs("EPSG:4326", "EPSG:2039", always_xy=True, force_over=True)

            self.df['X'], self.df['Y'] = transformer.transform(
                self.df[lon_col].astype(float).round(8),  # Maintain high precision
                self.df[lat_col].astype(float).round(8)
            )

            output_file = self.file_path.replace(".xlsx", "_converted.xlsx")
            self.df.to_excel(output_file, index=False)
            messagebox.showinfo("Success", f"Converted coordinates saved to {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")

    def convert_tm_to_wgs(self):
        if self.df is None:
            messagebox.showerror("Error", "Please upload an Excel file first.")
            return

        x_col = self.col_dropdown1.get()
        y_col = self.col_dropdown2.get()

        if not x_col or not y_col:
            messagebox.showerror("Error", "Please select X and Y columns.")
            return

        try:
            transformer = pyproj.Transformer.from_crs("EPSG:2039", "EPSG:4326", always_xy=True, force_over=True)

            self.df['Longitude'], self.df['Latitude'] = transformer.transform(
                self.df[x_col].astype(float).round(3),  # Maintain GIS precision
                self.df[y_col].astype(float).round(3)
            )

            output_file = self.file_path.replace(".xlsx", "_reverted.xlsx")
            self.df.to_excel(output_file, index=False)
            messagebox.showinfo("Success", f"Converted coordinates saved to {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = SuperDuperConverter(root)
    root.mainloop()
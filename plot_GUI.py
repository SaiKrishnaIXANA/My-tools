import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import ttk

class PlottingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV/Excel Plotter")
        self.root.geometry("800x600")  # Main window

        self.filepath = None
        self.df = None
        self.excel_file = None  # To store the loaded Excel file

        self.create_widgets()

    def create_widgets(self):
        # Widgets frame
        self.widgets_frame = tk.Frame(self.root)
        self.widgets_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # Button to import file
        self.import_button = tk.Button(self.widgets_frame, text="Import CSV/Excel", command=self.import_file)
        self.import_button.grid(row=0, column=0, padx=10, pady=10)

        # Label to show the file path
        self.file_label = tk.Label(self.widgets_frame, text="No file selected")
        self.file_label.grid(row=1, column=0, padx=10, pady=10)

        # Dropdown for sheet name (Excel only)
        self.sheet_label = tk.Label(self.widgets_frame, text="Select Sheet (for Excel):")
        self.sheet_label.grid(row=2, column=0, padx=10, pady=5)
        self.sheet_menu = ttk.Combobox(self.widgets_frame, state='readonly')
        self.sheet_menu.grid(row=2, column=1, padx=10, pady=5)
        self.sheet_menu.bind("<<ComboboxSelected>>", self.on_sheet_selected)  # Bind selection event

        # Dropdowns for axes
        self.x_axis_label = tk.Label(self.widgets_frame, text="Select X-axis:")
        self.x_axis_label.grid(row=3, column=0, padx=10, pady=5)
        self.x_axis_menu = ttk.Combobox(self.widgets_frame, state='readonly')
        self.x_axis_menu.grid(row=3, column=1, padx=10, pady=5)

        self.y_axis_label = tk.Label(self.widgets_frame, text="Select Y-axis:")
        self.y_axis_label.grid(row=4, column=0, padx=10, pady=5)
        self.y_axis_menu = ttk.Combobox(self.widgets_frame, state='readonly')
        self.y_axis_menu.grid(row=4, column=1, padx=10, pady=5)

        self.legend_axis_label = tk.Label(self.widgets_frame, text="Select Legend Axis (Optional, use Ctrl for multiple):")
        self.legend_axis_label.grid(row=5, column=0, padx=10, pady=5)
        self.legend_axis_menu = tk.Listbox(self.widgets_frame, selectmode=tk.MULTIPLE, height=5)
        self.legend_axis_menu.grid(row=5, column=1, padx=10, pady=5, sticky="ew")

        # Y-axis min/max
        self.y_min_label = tk.Label(self.widgets_frame, text="Y-axis Min:")
        self.y_min_label.grid(row=6, column=0, padx=10, pady=5)
        self.y_min_entry = tk.Entry(self.widgets_frame)
        self.y_min_entry.grid(row=6, column=1, padx=10, pady=5)

        self.y_max_label = tk.Label(self.widgets_frame, text="Y-axis Max:")
        self.y_max_label.grid(row=7, column=0, padx=10, pady=5)
        self.y_max_entry = tk.Entry(self.widgets_frame)
        self.y_max_entry.grid(row=7, column=1, padx=10, pady=5)

        # Buttons for Plot and Save next to each other
        self.plot_button = tk.Button(self.widgets_frame, text="Plot", command=self.plot)
        self.plot_button.grid(row=8, column=0, padx=10, pady=10)

        self.save_button = tk.Button(self.widgets_frame, text="Save Plot", command=self.save_plot)
        self.save_button.grid(row=8, column=1, padx=10, pady=10)

    def import_file(self):
        self.filepath = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
        
        try:
            if self.filepath.endswith('.csv'):
                self.df = pd.read_csv(self.filepath)
                self.excel_file = None
                self.sheet_menu['values'] = []  # No sheet for CSV
                messagebox.showinfo("File Loaded", f"Loaded file: {self.filepath}")
            elif self.filepath.endswith('.xlsx'):
                self.excel_file = pd.ExcelFile(self.filepath)
                self.sheet_menu['values'] = self.excel_file.sheet_names
                self.sheet_menu.current(0)
                self.load_excel_sheet(self.excel_file.sheet_names[0])  # Load the first sheet by default
                messagebox.showinfo("File Loaded", f"Loaded Excel file: {self.filepath}")
            else:
                messagebox.showerror("Error", "Unsupported file type")
                return

            # Update label with file path
            self.file_label.config(text=f"Loaded file: {self.filepath}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {str(e)}")

    def on_sheet_selected(self, event):
        """Event handler when a sheet is selected from the dropdown."""
        selected_sheet = self.sheet_menu.get()
        self.load_excel_sheet(selected_sheet)

    def load_excel_sheet(self, sheet_name):
        try:
            if self.excel_file:
                self.df = pd.read_excel(self.excel_file, sheet_name=sheet_name)
                self.update_dropdowns()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sheet: {str(e)}")

    def update_dropdowns(self):
        if self.df is not None:
            columns = list(self.df.columns)
            self.x_axis_menu['values'] = columns
            self.y_axis_menu['values'] = columns
            
            # Clear and update legend listbox for multiple selections
            self.legend_axis_menu.delete(0, tk.END)
            for col in columns:
                self.legend_axis_menu.insert(tk.END, col)
            
            # Set defaults for X and Y axis
            self.x_axis_menu.current(0)
            self.y_axis_menu.current(1)

    def plot(self):
        x_col = self.x_axis_menu.get()
        y_col = self.y_axis_menu.get()
        
        # Get selected legend columns
        legend_indices = self.legend_axis_menu.curselection()
        legend_cols = [self.legend_axis_menu.get(i) for i in legend_indices]
        
        y_min = self.y_min_entry.get()
        y_max = self.y_max_entry.get()

        try:
            fig, ax = plt.subplots()

            # Group by multiple columns if legends are selected
            if legend_cols:
                grouped = self.df.groupby(legend_cols)
                for key, grp in grouped:
                    label = ', '.join([str(k) for k in key]) if isinstance(key, tuple) else str(key)
                    ax.plot(grp[x_col], grp[y_col], label=label, linestyle='-', marker='o')


                ax.legend()
            else:
                ax.plot(self.df[x_col], self.df[y_col])
            
            ax.set_xlabel(x_col, fontsize=12, fontweight='bold')
            ax.set_ylabel(y_col, fontsize=12, fontweight='bold')
            
            # Apply gridlines to both X and Y axes
            ax.grid(True)

            if y_min:
                ax.set_ylim(bottom=float(y_min))
            if y_max:
                ax.set_ylim(top=float(y_max))
                
            # Set tick properties for bold font
            plt.xticks(fontsize=10, rotation=45, fontweight='bold')
            plt.yticks(fontsize=10, fontweight='bold')

            # Open a new dialog for the plot
            plot_window = tk.Toplevel(self.root)
            plot_window.title("Plot")
            plot_window.geometry("800x600")  # Set the size for the plot window

            canvas = FigureCanvasTkAgg(fig, master=plot_window)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def save_plot(self):
        if not self.df:
            messagebox.showerror("Error", "No plot to save")
            return
        
        file = filedialog.asksaveasfilename(defaultextension=".png",
                                            filetypes=[("PNG files", "*.png"), ("JPEG files", "*.jpg")])
        if file:
            plt.savefig(file)
            messagebox.showinfo("Success", "Plot saved successfully")

if __name__ == "__main__":
    root = tk.Tk()
    app = PlottingApp(root)
    root.mainloop()

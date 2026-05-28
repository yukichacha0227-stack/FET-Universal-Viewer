import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import numpy as np
import os
import datetime

from src.data_loader import merge_measurement_files, write_merged_excel

# Plotting settings
import matplotlib
matplotlib.use('TkAgg') 
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.cm as cm
import matplotlib.ticker as ticker

# Windows Excel Automation
try:
    import win32com.client as win32
    # Excel Constants
    xlXYScatterSmooth = 72
    xlTickMarkInside = 4
    xlCategory = 1
    xlValue = 2
    xlLegendPositionTop = -4160
    xlMarkerStyleCircle = 8
    
except ImportError:
    win32 = None

# Font setup for Matplotlib (English/Standard)
plt.rcParams['font.family'] = 'Arial' 

print("FET Universal Viewer Starting...")


def axis_bounds(values):
    finite_values = np.asarray(values, dtype=float)
    finite_values = finite_values[np.isfinite(finite_values)]
    if finite_values.size == 0:
        return 0.0, 1.0

    min_value = float(np.min(finite_values))
    max_value = float(np.max(finite_values))
    if min_value == max_value:
        padding = abs(min_value) * 0.05 or 1.0
        return min_value - padding, max_value + padding

    return min_value, max_value


class FET_Bilingual_Grapher(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('FET Universal Viewer - Bilingual Edition')
        self.geometry('1280x900')
        self.configure(bg='#f5f5f5')

        self.df = None
        self.last_saved_path = ""
        
        self.create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        paned = tk.PanedWindow(self, orient='horizontal', sashrelief='raised')
        paned.pack(fill='both', expand=True)

        # --- Left Control Panel ---
        control = tk.Frame(paned, bg='#f5f5f5', width=380, padx=10, pady=10)
        paned.add(control, minsize=380)

        # 1. Batch Load
        self.add_lbl(control, "1. Import & Merge")
        tk.Button(control, text="Select Files & Merge", command=self.load_batch_files,
                  bg='white', relief='groove', font=("Arial", 10, "bold")).pack(fill='x', pady=5)
        
        self.lbl_status = tk.Label(control, text="Ready", bg='#f5f5f5', fg='blue', wraplength=350)
        self.lbl_status.pack(fill='x')

        # 2. Mapping
        self.add_lbl(control, "2. Column Mapping")
        grid = tk.Frame(control, bg='#f5f5f5')
        grid.pack(fill='x')
        self.ent_vsd = self.add_inp(grid, "V_SD (Col Name):", "Vsd", 0)
        self.ent_vbg = self.add_inp(grid, "V_BG (Col Name):", "Vbg", 1)
        self.ent_isd = self.add_inp(grid, "I_SD (Col Name):", "Isd", 2)

        # 3. Plot Mode
        self.add_lbl(control, "3. Plot Mode")
        self.plot_type = tk.StringVar(value="output")
        mode_frame = tk.Frame(control, bg='white', relief='sunken', bd=1)
        mode_frame.pack(fill='x', pady=5, ipadx=5, ipady=5)
        tk.Radiobutton(mode_frame, text="Output Char. (Isd - Vsd)\n[Group by Vbg]", 
                       variable=self.plot_type, value="output", bg='white', justify='left').pack(anchor='w', padx=5, pady=2)
        tk.Radiobutton(mode_frame, text="Transfer Char. (Isd - Vbg)\n[Group by Vsd]", 
                       variable=self.plot_type, value="transfer", bg='white', justify='left').pack(anchor='w', padx=5, pady=2)

        # 4. Settings
        self.add_lbl(control, "4. Settings")
        self.is_log = tk.BooleanVar(value=False)
        tk.Checkbutton(control, text="Log Scale (|Y-axis|)", variable=self.is_log, bg='#f5f5f5').pack(anchor='w')

        # Run Buttons
        tk.Button(control, text="Preview (Python)", command=self.plot_graph,
                  bg='#e6f2ff', font=("Arial", 11, "bold")).pack(fill='x', pady=10)

        tk.Button(control, text="Create Excel Graphs (Japanese Legend)", command=self.create_native_excel_charts,
                  bg='#ffebd0', fg='#d35400', font=("Arial", 11, "bold")).pack(fill='x', pady=5)

        self.txt_info = tk.Text(control, height=12, width=40)
        self.txt_info.pack(fill='both', expand=True)

        # --- Right Graph Area ---
        g_frame = tk.Frame(paned, bg='white')
        paned.add(g_frame)
        self.fig, self.ax = plt.subplots(figsize=(6, 5))
        self.fig.subplots_adjust(left=0.18, right=0.95, top=0.9, bottom=0.15)
        self.canvas = FigureCanvasTkAgg(self.fig, master=g_frame)
        self.canvas.get_tk_widget().pack(side='top', fill='both', expand=True)
        NavigationToolbar2Tk(self.canvas, g_frame).update()

    def add_lbl(self, p, t):
        tk.Label(p, text=t, font=("Arial", 11, "bold"), bg='#f5f5f5').pack(anchor='w', pady=(15, 2))

    def add_inp(self, p, l, d, r):
        tk.Label(p, text=l, bg='#f5f5f5').grid(row=r, column=0, sticky='w', pady=2)
        e = tk.Entry(p, width=10)
        e.insert(0, d)
        e.grid(row=r, column=1, sticky='e', pady=2)
        return e

    # --- 1. File Load & Merge ---
    def load_batch_files(self):
        ft = [("Data Files", "*.Dat *.csv *.txt"), ("All Files", "*.*")]
        file_paths = filedialog.askopenfilenames(filetypes=ft)
        if not file_paths: return

        dir_name = os.path.dirname(file_paths[0])
        now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.last_saved_path = os.path.join(dir_name, f"Merged_{now_str}.xlsx").replace('/', '\\')

        self.lbl_status.config(text="Processing...", fg="blue")
        self.update()
        
        try:
            result = merge_measurement_files(file_paths)
            self.lbl_status.config(text="Writing Excel workbook...", fg="orange")
            self.update()
            write_merged_excel(result, self.last_saved_path)

            self.df = result.dataframe
            self.lbl_status.config(text="Success!", fg="green")

            self.ent_vsd.delete(0, tk.END)
            self.ent_vsd.insert(0, "Vsd")
            self.ent_vbg.delete(0, tk.END)
            self.ent_vbg.insert(0, "Vbg")
            self.ent_isd.delete(0, tk.END)
            self.ent_isd.insert(0, "Isd")
            
            info = f"Merged: {os.path.basename(self.last_saved_path)}\n\n"
            info += f"Files: {result.processed_count}\nTotal Points: {len(self.df)}"
            self.txt_info.delete(1.0, tk.END)
            self.txt_info.insert(tk.END, info)
            messagebox.showinfo("Success", f"Merged Excel created!\nNext: Click 'Create Excel Graphs'")

        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.lbl_status.config(text="Error")

    # --- 2. Create Native Excel Charts (Japanese Legend) ---
    def create_native_excel_charts(self):
        if not self.last_saved_path or not os.path.exists(self.last_saved_path):
            messagebox.showwarning("Warning", "Please merge files first.")
            return
        if self.df is None: return

        if win32 is None:
            messagebox.showerror("Error", "pywin32 library is missing.")
            return

        c_vsd = self.ent_vsd.get().strip()
        c_vbg = self.ent_vbg.get().strip()
        c_isd = self.ent_isd.get().strip()
        ptype = self.plot_type.get()

        if ptype == "output":
            x_col, y_col, grp_col_R = c_vsd, c_isd, 'Vbg_R'
            legend_prefix = "Vbg"
        else:
            x_col, y_col, grp_col_R = c_vbg, c_isd, 'Vsd_R'
            legend_prefix = "Vsd"

        if grp_col_R not in self.df.columns:
            orig = c_vbg if ptype == "output" else c_vsd
            self.df[grp_col_R] = self.df[orig].round(2)
        
        groups = sorted(self.df[grp_col_R].dropna().unique())

        self.lbl_status.config(text="Opening Excel...", fg="blue")
        self.update()

        excel = None
        wb = None
        
        # Excel Colors
        COLOR_BLUE_XL = 12611584 # RGB(0, 112, 192)
        COLOR_ORANGE_XL = 3243501 # RGB(237, 125, 49)

        try:
            excel = win32.Dispatch('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False

            wb = excel.Workbooks.Open(self.last_saved_path)
            
            sheet_name = "Graphs"
            try: wb.Sheets(sheet_name).Delete()
            except: pass
            ws = wb.Sheets.Add()
            ws.Name = sheet_name

            current_row = 1
            chart_top = 10

            self.lbl_status.config(text="Creating Final Graphs...", fg="blue")
            self.update()

            for val in groups:
                grp_df = self.df[self.df[grp_col_R] == val].sort_values(by='_Sort_ID')
                if len(grp_df) == 0: continue

                x_vals = grp_df[x_col].values
                y_vals = grp_df[y_col].values
                
                split_idx = np.argmax(x_vals)
                if split_idx == 0 or split_idx == len(x_vals)-1:
                    split_idx = np.argmin(x_vals)

                start_row = current_row
                ws.Cells(start_row, 1).Value = f"{legend_prefix}={val}"
                ws.Cells(start_row+1, 1).Value = "X"
                ws.Cells(start_row+1, 2).Value = "Y"
                
                data_len = len(x_vals)
                data_chunk = np.vstack((x_vals, y_vals)).T.tolist()
                ws.Range(f"A{start_row+2}:B{start_row+1+data_len}").Value = data_chunk
                
                data_start = start_row + 2
                
                # Chart
                chart_obj = ws.Shapes.AddChart2(-1, xlXYScatterSmooth)
                chart = chart_obj.Chart
                chart.ChartTitle.Text = f"{legend_prefix}={val}"
                
                chart_obj.Left = 200 
                chart_obj.Top = chart_top
                chart_obj.Width = 350
                chart_obj.Height = 250
                
                for _ in range(chart.SeriesCollection().Count):
                    chart.SeriesCollection(1).Delete()

                # Series
                len1 = split_idx + 1
                range_x1 = ws.Range(f"A{data_start}:A{data_start + len1 - 1}")
                range_y1 = ws.Range(f"B{data_start}:B{data_start + len1 - 1}")
                range_x2 = ws.Range(f"A{data_start + split_idx}:A{data_start + data_len - 1}")
                range_y2 = ws.Range(f"B{data_start + split_idx}:B{data_start + data_len - 1}")

                start_val = x_vals[0]
                mid_val = x_vals[split_idx]
                
                def add_series_xl(name, rx, ry, color):
                    s = chart.SeriesCollection().NewSeries()
                    s.Name = name # Japanese
                    s.XValues = rx
                    s.Values = ry
                    s.Format.Line.ForeColor.RGB = color
                    s.Format.Line.Weight = 1.5
                    s.MarkerStyle = xlMarkerStyleCircle
                    s.MarkerBackgroundColor = color
                    s.MarkerForegroundColor = color
                    s.MarkerSize = 5

                # Japanese Labels for Excel
                if start_val < mid_val:
                    if len1 > 1: add_series_xl("順掃引", range_x1, range_y1, COLOR_BLUE_XL)
                    if (data_len - split_idx) > 1: add_series_xl("逆掃引", range_x2, range_y2, COLOR_ORANGE_XL)
                else:
                    if len1 > 1: add_series_xl("逆掃引", range_x1, range_y1, COLOR_ORANGE_XL)
                    if (data_len - split_idx) > 1: add_series_xl("順掃引", range_x2, range_y2, COLOR_BLUE_XL)

                # Styling
                try:
                    chart.Axes(xlCategory).HasMajorGridlines = False
                    chart.Axes(xlValue).HasMajorGridlines = False
                except: pass

                chart.Axes(xlCategory).MajorTickMark = xlTickMarkInside
                chart.Axes(xlValue).MajorTickMark = xlTickMarkInside
                chart.Axes(xlCategory).Format.Line.ForeColor.RGB = 0
                chart.Axes(xlValue).Format.Line.ForeColor.RGB = 0 

                x_min, x_max = axis_bounds(x_vals)
                y_min, y_max = axis_bounds(y_vals)
                chart.Axes(xlCategory).MinimumScale = x_min
                chart.Axes(xlCategory).MaximumScale = x_max
                chart.Axes(xlValue).MinimumScale = y_min
                chart.Axes(xlValue).MaximumScale = y_max
                
                # Keep axes on the lower-left chart frame instead of a hard-coded value.
                chart.Axes(xlCategory).CrossesAt = y_min
                chart.Axes(xlValue).CrossesAt = x_min

                chart.Axes(xlValue).TickLabels.NumberFormat = "0.00E+00"
                chart.Axes(xlCategory).TickLabels.NumberFormat = "General"

                chart.PlotArea.Format.Line.Visible = True
                chart.PlotArea.Format.Line.ForeColor.RGB = 0 
                chart.PlotArea.Format.Line.Weight = 1.0
                chart.ChartArea.Format.Line.Visible = False
                chart.HasLegend = True
                chart.Legend.Position = xlLegendPositionTop

                current_row = data_start + data_len + 5
                chart_top += 260

            wb.Save()
            self.lbl_status.config(text="Complete!", fg="green")
            messagebox.showinfo("Success", f"Graphs Created!\nFile: {os.path.basename(self.last_saved_path)}")

        except Exception as e:
            self.lbl_status.config(text="Error", fg="red")
            messagebox.showerror("Win32 Error", str(e))
        finally:
            if excel: excel.Quit()

    # --- 3. Python Preview (English, Colored Groups) ---
    def plot_graph(self):
        if self.df is None: return
        c_vsd = self.ent_vsd.get().strip()
        c_vbg = self.ent_vbg.get().strip()
        c_isd = self.ent_isd.get().strip()
        
        if not {c_vsd, c_vbg, c_isd}.issubset(self.df.columns):
            messagebox.showerror("Error", "Check columns.")
            return

        self.ax.clear()
        ptype = self.plot_type.get()
        log_mode = self.is_log.get()
        
        if ptype == "output":
            self.ax.set_title("Output Characteristics ($I_{sd} - V_{sd}$)")
            self.ax.set_xlabel(f"{c_vsd} (V)")
            self.ax.set_ylabel(f"{c_isd} (A)")
            x_col, y_col, group_col_R = c_vsd, c_isd, 'Vbg_R'
            legend_title = "Vbg"
        else:
            self.ax.set_title("Transfer Characteristics ($I_{sd} - V_{bg}$)")
            self.ax.set_xlabel(f"{c_vbg} (V)")
            self.ax.set_ylabel(f"{c_isd} (A)")
            x_col, y_col, group_col_R = c_vbg, c_isd, 'Vsd_R'
            legend_title = "Vsd"

        if log_mode:
            self.ax.set_ylabel(f"|{c_isd}| (A)")
            self.ax.set_yscale('log')
        else:
            self.ax.yaxis.set_major_formatter(ticker.FormatStrFormatter('%.2E'))
            self.ax.yaxis.set_major_locator(ticker.MaxNLocator(nbins=10, prune=None))

        self.ax.grid(True, which="both", ls="--", alpha=0.5)

        if group_col_R not in self.df.columns:
            orig = c_vbg if ptype == "output" else c_vsd
            self.df[group_col_R] = self.df[orig].round(2)

        groups = sorted(self.df[group_col_R].dropna().unique())
        
        # Colormap for different groups.
        colors = cm.viridis(np.linspace(0, 1, len(groups)))

        for i, val in enumerate(groups):
            subset = self.df[self.df[group_col_R] == val].sort_values(by='_Sort_ID')
            if len(subset) == 0: continue

            x_vals = subset[x_col].values
            y_vals = subset[y_col].values
            if log_mode:
                y_vals = np.abs(y_vals.astype(float))
                y_vals[y_vals <= 0] = np.nan

            split_idx = np.argmax(x_vals)
            if split_idx == 0 or split_idx == len(x_vals)-1:
                split_idx = np.argmin(x_vals)

            x1, y1 = x_vals[:split_idx+1], y_vals[:split_idx+1]
            x2, y2 = x_vals[split_idx:], y_vals[split_idx:]

            label_base = f"{legend_title}={val}V"
            start_val = x_vals[0]
            mid_val = x_vals[split_idx]
            
            # Using Group Color, distinguishing direction by Line Style
            grp_color = colors[i]

            # English Legend
            if start_val < mid_val:
                if len(x1) > 1: self.ax.plot(x1, y1, marker='.', markersize=3, color=grp_color, linestyle='-', label=f"{label_base} (Fwd)")
                if len(x2) > 1: self.ax.plot(x2, y2, marker='.', markersize=3, color=grp_color, linestyle='--', label=f"{label_base} (Rev)")
            else:
                if len(x1) > 1: self.ax.plot(x1, y1, marker='.', markersize=3, color=grp_color, linestyle='--', label=f"{label_base} (Rev)")
                if len(x2) > 1: self.ax.plot(x2, y2, marker='.', markersize=3, color=grp_color, linestyle='-', label=f"{label_base} (Fwd)")

        self.ax.legend(fontsize='small', bbox_to_anchor=(1.05, 1), loc='upper left')
        self.fig.tight_layout()
        self.canvas.draw()

    def on_closing(self):
        self.destroy()
        sys.exit()

if __name__ == "__main__":
    app = FET_Bilingual_Grapher()
    app.mainloop()

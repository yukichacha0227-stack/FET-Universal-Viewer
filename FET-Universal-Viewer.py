import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np
import os
import datetime
import tempfile

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
    xlSolid = 1
    xlCategory = 1
    xlValue = 2
    xlLegendPositionTop = -4160
    xlMarkerStyleCircle = 8
    
except ImportError:
    win32 = None

# Font setup for Matplotlib (English/Standard)
plt.rcParams['font.family'] = 'Arial' 

print("FET Universal Viewer (Bilingual Final) Starting...")

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
        tk.Button(control, text="📂 Select Files & Merge", command=self.load_batch_files, 
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
        tk.Checkbutton(control, text="Log Scale (Y-axis)", variable=self.is_log, bg='#f5f5f5').pack(anchor='w')

        # Run Buttons
        tk.Button(control, text="📊 Preview (Python)", command=self.plot_graph, 
                  bg='#e6f2ff', font=("Arial", 11, "bold")).pack(fill='x', pady=10)

        tk.Button(control, text="📉 Create Excel Graphs (Jpn Legend)", command=self.create_native_excel_charts, 
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
        if win32 is None:
            messagebox.showerror("Error", "pywin32 library is missing.")
            return

        ft = [("Data Files", "*.Dat *.csv *.txt"), ("All Files", "*.*")]
        file_paths = filedialog.askopenfilenames(filetypes=ft)
        if not file_paths: return

        dir_name = os.path.dirname(file_paths[0])
        now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.last_saved_path = os.path.join(dir_name, f"Merged_{now_str}.xlsx").replace('/', '\\')

        self.lbl_status.config(text="Processing...", fg="blue")
        self.update()

        excel = None
        all_dfs = []
        sys_temp_dir = tempfile.gettempdir()
        
        try:
            excel = win32.Dispatch('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False

            c_vsd = self.ent_vsd.get().strip()
            c_vbg = self.ent_vbg.get().strip()
            c_isd = self.ent_isd.get().strip()

            processed_count = 0
            
            for i, path in enumerate(file_paths):
                self.lbl_status.config(text=f"Reading ({i+1}/{len(file_paths)})", fg="blue")
                self.update()
                
                abs_path = os.path.abspath(path).replace('/', '\\')
                wb = excel.Workbooks.Open(abs_path)
                ws = wb.Sheets(1)

                row1_vals = ""
                try:
                    vals = [str(ws.Cells(1, c).Value) for c in range(1, 11)]
                    row1_vals = " ".join(vals).lower()
                except: pass
                
                keywords = ["isd", "vsd", "vbg", "no.", "id", "vg", "vd"]
                if not any(k in row1_vals for k in keywords):
                    ws.Rows(1).Insert()
                    # Headers (English for App compatibility)
                    headers = ["No.", "Temp", "Mag", "Isd", "Vsd", "Vbg"]
                    for c, h in enumerate(headers, 1):
                        ws.Cells(1, c).Value = h
                
                temp_name = f"fet_tmp_{now_str}_{i}.xlsx"
                temp_path = os.path.join(sys_temp_dir, temp_name).replace('/', '\\')
                wb.SaveAs(temp_path, FileFormat=51)
                wb.Close(SaveChanges=False)

                df_temp = pd.read_excel(temp_path)
                df_temp.columns = df_temp.columns.str.strip()
                
                for col in [c_vsd, c_vbg, c_isd]:
                    if col in df_temp.columns:
                        df_temp[col] = pd.to_numeric(df_temp[col], errors='coerce')
                
                if {c_vsd, c_vbg, c_isd}.issubset(df_temp.columns):
                    df_temp = df_temp.dropna(subset=[c_vsd, c_vbg, c_isd])
                    df_temp['_Sort_ID'] = i * 1000000 + np.arange(len(df_temp))
                    all_dfs.append(df_temp)
                    processed_count += 1
                
                try: os.remove(temp_path)
                except: pass

            if not all_dfs:
                messagebox.showerror("Error", "No valid data found.")
                return

            self.lbl_status.config(text="Merging Excel...", fg="orange")
            self.update()

            with pd.ExcelWriter(self.last_saved_path, engine='openpyxl') as writer:
                current_col = 0
                for df in all_dfs:
                    export_df = df.drop(columns=['_Sort_ID'], errors='ignore')
                    export_df.to_excel(writer, sheet_name='Raw_Data', startcol=current_col, index=False)
                    current_col += len(export_df.columns) + 1

            self.df = pd.concat(all_dfs, ignore_index=True)
            self.lbl_status.config(text="Success!", fg="green")
            
            self.df['Vsd_R'] = self.df[c_vsd].round(2)
            self.df['Vbg_R'] = self.df[c_vbg].round(2)
            
            info = f"Merged: {os.path.basename(self.last_saved_path)}\n\n"
            info += f"Files: {processed_count}\nTotal Points: {len(self.df)}"
            self.txt_info.delete(1.0, tk.END)
            self.txt_info.insert(tk.END, info)
            messagebox.showinfo("Success", f"Merged Excel created!\nNext: Click 'Create Excel Graphs'")

        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.lbl_status.config(text="Error")
        finally:
            if excel: excel.Quit()

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
        
        groups = sorted(self.df[grp_col_R].unique())

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

                # Styling (Axes at -200)
                try:
                    chart.Axes(xlCategory).HasMajorGridlines = False
                    chart.Axes(xlValue).HasMajorGridlines = False
                except: pass

                chart.Axes(xlCategory).MajorTickMark = xlTickMarkInside
                chart.Axes(xlValue).MajorTickMark = xlTickMarkInside
                chart.Axes(xlCategory).Format.Line.ForeColor.RGB = 0
                chart.Axes(xlValue).Format.Line.ForeColor.RGB = 0 

                x_min = np.min(x_vals)
                x_max = np.max(x_vals)
                chart.Axes(xlCategory).MinimumScale = x_min
                chart.Axes(xlCategory).MaximumScale = x_max
                
                # Axis Crosses at -200
                chart.Axes(xlCategory).CrossesAt = -200 
                chart.Axes(xlValue).CrossesAt = -200

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

        if self.is_log.get(): self.ax.set_yscale('log')
        else:
            self.ax.yaxis.set_major_formatter(ticker.FormatStrFormatter('%.2E'))
            self.ax.yaxis.set_major_locator(ticker.MaxNLocator(nbins=10, prune=None))

        self.ax.grid(True, which="both", ls="--", alpha=0.5)

        if group_col_R not in self.df.columns:
            orig = c_vbg if ptype == "output" else c_vsd
            self.df[group_col_R] = self.df[orig].round(2)

        groups = sorted(self.df[group_col_R].unique())
        
        # Colormap for different groups (Multi-Color)
        colors = cm.jet(np.linspace(0, 1, len(groups)))

        for i, val in enumerate(groups):
            subset = self.df[self.df[group_col_R] == val].sort_values(by='_Sort_ID')
            if len(subset) == 0: continue

            x_vals = subset[x_col].values
            y_vals = subset[y_col].values

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
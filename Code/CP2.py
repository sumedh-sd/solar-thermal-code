import math
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np

class SolarDesignApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Solar Site Analysis & System Design")
        self.root.geometry("1400x850")
        self.root.configure(bg="#f4f4f4")

        # Tech 1 variables
        self.file_path_tech1 = tk.StringVar()
        self.current_df = None  # Store loaded dataframe for graphing
        self.current_threshold = 650  # Store current threshold
        self.current_modal = 700  # Store current modal cluster
        
        # Tech 2 variables
        self.file_path_tech2 = tk.StringVar()
        self.optn1 = tk.StringVar(value="A")
        self.optn2 = tk.StringVar(value="C")

        self.create_widgets()

    def create_widgets(self):
        # Title
        tk.Label(self.root, text="SOLAR SITE ANALYSIS & SYSTEM DESIGN", 
                 font=('Arial', 14, 'bold'), bg="#f4f4f4", fg="#1976D2").pack(pady=10)

        # TWO-COLUMN LAYOUT
        main_container = tk.Frame(self.root, bg="#f4f4f4")
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # ========================================================================
        # LEFT COLUMN: TECHNIQUE 1 - Direct File Extraction
        # ========================================================================
        left_col = tk.Frame(main_container, bg="#f4f4f4", relief="solid", borderwidth=2)
        left_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        # Header
        tk.Label(left_col, text="TECHNIQUE ① - Weather Data File", 
                 font=('Arial', 12, 'bold'), fg="white", bg="#1976D2").pack(fill=tk.X, pady=5)
        
        tk.Label(left_col, text="Direct Extraction from NSRDB", 
                 font=('Arial', 9), bg="#f4f4f4").pack(pady=5)

        # File selection for Tech 1
        tk.Label(left_col, text="Select NSRDB CSV File:", font=('Arial', 9, 'bold'), 
                 bg="#f4f4f4").pack(pady=(10,5))
        f1_frame = tk.Frame(left_col, bg="#f4f4f4")
        f1_frame.pack()
        tk.Entry(f1_frame, textvariable=self.file_path_tech1, width=45).pack(side=tk.LEFT, padx=5)
        tk.Button(f1_frame, text="Browse", command=self.browse_file_tech1, 
                 bg="#1976D2", fg="white").pack(side=tk.LEFT)

        # Run button for Tech 1
        self.btn_tech1 = tk.Button(left_col, text="RUN EXTRACTION & CALC", 
                                     bg="#1976D2", fg="white", height=2, font=('Arial', 10, 'bold'),
                                     command=lambda: self.start_thread("tech1"))
        self.btn_tech1.pack(pady=15, padx=20, fill=tk.X)

        # Generate Graph button for Tech 1
        self.btn_graph = tk.Button(left_col, text="📊 GENERATE DNI DISTRIBUTION GRAPH", 
                                    bg="#FF9800", fg="white", height=2, font=('Arial', 9, 'bold'),
                                    command=self.generate_dni_graph, state="disabled")
        self.btn_graph.pack(pady=5, padx=20, fill=tk.X)

        # Result Display for Technique 1
        tk.Label(left_col, text="Result Display:", font=('Arial', 9, 'bold'), 
                 bg="#f4f4f4").pack(pady=(10,5))
        scroll1 = tk.Scrollbar(left_col)
        scroll1.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 5))
        self.txt_output1 = tk.Text(left_col, height=32, width=55, font=('Courier', 8), 
                                   bg="white", yscrollcommand=scroll1.set)
        self.txt_output1.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        scroll1.config(command=self.txt_output1.yview)

        # ========================================================================
        # RIGHT COLUMN: TECHNIQUE 2 - Mathematical Model
        # ========================================================================
        right_col = tk.Frame(main_container, bg="#f4f4f4", relief="solid", borderwidth=2)
        right_col.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

        # Header
        tk.Label(right_col, text="TECHNIQUE ② - Mathematical Model", 
                 font=('Arial', 12, 'bold'), fg="white", bg="#388E3C").pack(fill=tk.X, pady=5)

        # File selection for Tech 2
        tk.Label(right_col, text="Select Excel or CSV File:", font=('Arial', 9, 'bold'), 
                 bg="#f4f4f4").pack(pady=(10,5))
        f2_frame = tk.Frame(right_col, bg="#f4f4f4")
        f2_frame.pack()
        tk.Entry(f2_frame, textvariable=self.file_path_tech2, width=45).pack(side=tk.LEFT, padx=5)
        tk.Button(f2_frame, text="Browse", command=self.browse_file_tech2, 
                 bg="#388E3C", fg="white").pack(side=tk.LEFT)

        # Irradiance Type Options
        tk.Label(right_col, text="Input Irradiance Type:", font=('Arial', 9, 'bold'), 
                 bg="#f4f4f4").pack(pady=(10,5))
        opt_frame = tk.Frame(right_col, bg="#f4f4f4")
        opt_frame.pack()
        tk.Radiobutton(opt_frame, text="Ig & Id", variable=self.optn1, value="A", 
                      bg="#f4f4f4").pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(opt_frame, text="Ig & Ibn", variable=self.optn1, value="B", 
                      bg="#f4f4f4").pack(side=tk.LEFT, padx=5)

        # Model Selection
        tk.Label(right_col, text="Estimation Model:", font=('Arial', 9, 'bold'), 
                 bg="#f4f4f4").pack(pady=(10,5))
        model_frame = tk.Frame(right_col, bg="#f4f4f4")
        model_frame.pack()
        tk.Radiobutton(model_frame, text="Liu & Jordan", variable=self.optn2, value="C", 
                      bg="#f4f4f4").pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(model_frame, text="HDKR Model", variable=self.optn2, value="D", 
                      bg="#f4f4f4").pack(side=tk.LEFT, padx=5)

        # Site Parameters
        tk.Label(right_col, text="Site Parameters:", font=('Arial', 9, 'bold'), 
                 bg="#f4f4f4").pack(pady=(10,5))
        
        p_container = tk.Frame(right_col, bg="white", relief="groove", borderwidth=1, padx=10, pady=10)
        p_container.pack(pady=5, padx=10, fill=tk.X)

        self.inputs = {}
        fields = [
            ("Latitude (deg):", "lat"),
            ("Longitude (deg):", "lon"),
            ("Tilt angle β (deg):", "beta"),
            ("Surface azimuth γ (deg):", "gamma"),
            ("Ground reflectivity ρ:", "rho")
        ]

        for text, key in fields:
            frame = tk.Frame(p_container, bg="white")
            frame.pack(fill="x", pady=2)
            tk.Label(frame, text=text, width=25, anchor='w', bg="white", 
                    font=('Arial', 9)).pack(side=tk.LEFT)
            ent = tk.Entry(frame, width=15)
            ent.pack(side=tk.RIGHT)
            self.inputs[key] = ent

        # Buttons for Tech 2
        btn_frame = tk.Frame(right_col, bg="#f4f4f4")
        btn_frame.pack(pady=15)

        self.btn_tech2 = tk.Button(btn_frame, text="RUN MODEL & SAVE", bg="#388E3C", 
                                   fg="white", width=20, height=2, font=('Arial', 10, 'bold'),
                                   command=lambda: self.start_thread("tech2"))
        self.btn_tech2.pack(side=tk.LEFT, padx=5)

        self.btn_clear = tk.Button(btn_frame, text="CLEAR ALL", bg="#d32f2f", 
                                   fg="white", width=15, height=2, font=('Arial', 10, 'bold'),
                                   command=self.manual_clear)
        self.btn_clear.pack(side=tk.LEFT, padx=5)

        # Result Display for Technique 2
        tk.Label(right_col, text="Status/Results:", font=('Arial', 9, 'bold'), 
                 bg="#f4f4f4").pack(pady=(10,5))
        scroll2 = tk.Scrollbar(right_col)
        scroll2.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 5))
        self.txt_output2 = tk.Text(right_col, height=15, width=55, font=('Courier', 8), 
                                   bg="white", yscrollcommand=scroll2.set)
        self.txt_output2.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        scroll2.config(command=self.txt_output2.yview)

    def manual_clear(self):
        """Clears all displays, file paths, and input boxes"""
        self.txt_output1.delete('1.0', tk.END)
        self.txt_output2.delete('1.0', tk.END)
        self.file_path_tech1.set("")
        self.file_path_tech2.set("")
        for ent in self.inputs.values():
            ent.delete(0, tk.END)
        self.current_df = None
        self.btn_graph.config(state="disabled")
        self.txt_output1.insert(tk.END, "✓ System Reset. All fields cleared.\n")
        self.txt_output2.insert(tk.END, "✓ System Reset. All fields cleared.\n")

    def browse_file_tech1(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
        if filename: 
            self.file_path_tech1.set(filename)

    def browse_file_tech2(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv"),
                      ("Excel files", "*.xlsx *.xls"),
                      ("CSV files", "*.csv"),
                      ("All files", "*.*")])
        if filename:
            self.file_path_tech2.set(filename)

    def start_thread(self, mode):
        if mode == "tech1":
            self.txt_output1.delete('1.0', tk.END)
            self.txt_output1.insert(tk.END, "⚙ Starting TECHNIQUE 1 analysis...\n")
            self.btn_tech1.config(state="disabled")
            target = self.run_tech1
        else:
            self.txt_output2.delete('1.0', tk.END)
            self.txt_output2.insert(tk.END, "⚙ Starting TECHNIQUE 2 processing...\n")
            self.btn_tech2.config(state="disabled")
            target = self.run_tech2
        
        threading.Thread(target=target, daemon=True).start()

    # ========================================================================
    # TECHNIQUE 1: Direct File Extraction
    # ========================================================================
    
    def load_clean_df(self, path):
        """Load and clean NSRDB CSV file, also extract lat/lon"""
        # Read first few rows to detect header structure and extract metadata
        df_preview = pd.read_csv(path, nrows=5)
        
        # Extract lat/lon from header if NSRDB format
        lat, lon = None, None
        if "Source" in str(df_preview.columns):
            # NSRDB format - lat/lon in row 2
            try:
                header_row = pd.read_csv(path, nrows=1, skiprows=1)
                if 'Latitude' in header_row.columns and 'Longitude' in header_row.columns:
                    lat = float(header_row['Latitude'].iloc[0])
                    lon = float(header_row['Longitude'].iloc[0])
            except:
                pass
            skip = 2
        else:
            skip = 0
        
        df = pd.read_csv(path, skiprows=skip)
            
        # Clean column names and convert to numeric
        df.columns = [c.strip() for c in df.columns]
        for col in ['GHI', 'DNI', 'DHI', 'Month', 'Day', 'Hour', 'Minute']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # CRITICAL FIX: Filter out invalid rows (Month < 1 or Day < 1)
        df = df[(df['Month'] >= 1) & (df['Month'] <= 12) & (df['Day'] >= 1) & (df['Day'] <= 31)]
        
        return df, lat, lon

    def calculate_dynamic_threshold(self, df, dni_col='DNI'):
        """Calculate dynamic threshold using modal DNI cluster"""
        dni_filtered = df[df[dni_col] > 500][dni_col]
        
        if len(dni_filtered) == 0:
            return 650, 700
        
        dni_rounded = (dni_filtered / 50).round() * 50
        value_counts = dni_rounded.value_counts()
        modal_cluster = int(value_counts.index[0])
        
        # Use 93% of modal cluster
        threshold = int(modal_cluster * 0.93)
        threshold = round(threshold / 10) * 10
        
        return threshold, modal_cluster

    def run_tech1(self):
        try:
            path = self.file_path_tech1.get()
            if not path:
                raise ValueError("Please select a CSV file first.")
            
            df, lat, lon = self.load_clean_df(path)
            
            # Store dataframe for graphing
            self.current_df = df
            
            # Auto-fill lat/lon in Tech 2 if available
            if lat is not None and lon is not None:
                self.inputs['lat'].delete(0, tk.END)
                self.inputs['lat'].insert(0, str(lat))
                self.inputs['lon'].delete(0, tk.END)
                self.inputs['lon'].insert(0, str(lon))
            
            # Calculate threshold
            threshold, modal_cluster = self.calculate_dynamic_threshold(df, 'DNI')
            
            # Store for graphing
            self.current_threshold = threshold
            self.current_modal = modal_cluster
            
            # Calculate daily totals
            daily_data = df.groupby(['Month', 'Day'])[['DNI', 'GHI']].sum() / 1000
            
            # Monthly averages and totals
            monthly_avg_daily = daily_data.groupby('Month').mean()
            monthly_totals = daily_data.groupby('Month').sum()
            annual_dni_total = monthly_totals['DNI'].sum()
            annual_ghi_total = monthly_totals['GHI'].sum()
            
            # Design DNI
            productive_hours = df[df['DNI'] >= threshold]
            design_dni = productive_hours['DNI'].mean() if not productive_hours.empty else 0
            productive_count = len(productive_hours)

            # Build report
            self.root.after(0, lambda: self.display_tech1_report(
                df, monthly_avg_daily, monthly_totals, annual_dni_total, annual_ghi_total,
                threshold, modal_cluster, design_dni, productive_count, lat, lon
            ))
            
            # Enable graph button
            self.root.after(0, lambda: self.btn_graph.config(state="normal"))
            
        except Exception as error:
            error_msg = str(error)
            self.root.after(0, lambda msg=error_msg: self.show_error_tech1(msg))

    def display_tech1_report(self, df, monthly_avg_daily, monthly_totals, 
                            annual_dni_total, annual_ghi_total, threshold, 
                            modal_cluster, design_dni, productive_count, lat=None, lon=None):
        
        self.txt_output1.delete('1.0', tk.END)
        
        # Configure bold tag
        self.txt_output1.tag_configure("bold", font=('Courier', 9, 'bold'))
        self.txt_output1.tag_configure("title", font=('Courier', 10, 'bold'))
        
        self.txt_output1.insert(tk.END, f"{'='*60}\n")
        self.txt_output1.insert(tk.END, f"{'TECHNIQUE 1: DIRECT FILE EXTRACTION'.center(60)}\n", "title")
        self.txt_output1.insert(tk.END, f"{'='*60}\n\n")
        
        # Show location if available
        if lat is not None and lon is not None:
            self.txt_output1.insert(tk.END, f"LOCATION INFORMATION\n")
            self.txt_output1.insert(tk.END, f"{'-'*60}\n")
            self.txt_output1.insert(tk.END, f"Latitude:  {lat:>8.2f}°  (Auto-filled in Tech 2 ✓)\n")
            self.txt_output1.insert(tk.END, f"Longitude: {lon:>8.2f}°  (Auto-filled in Tech 2 ✓)\n")
            self.txt_output1.insert(tk.END, f"\n{'='*60}\n\n")
        
        self.txt_output1.insert(tk.END, f"MONTHLY SUMMARY (Day-wise Averages)\n")
        self.txt_output1.insert(tk.END, f"{'-'*60}\n")
        self.txt_output1.insert(tk.END, f"{'Month':<8} {'Avg DNI':>12} {'Avg GHI':>12} {'Total DNI':>14}\n")
        self.txt_output1.insert(tk.END, f"{'':8} {'(kWh/m²/d)':>12} {'(kWh/m²/d)':>12} {'(kWh/m²)':>14}\n")
        self.txt_output1.insert(tk.END, f"{'-'*60}\n")
        
        mo_names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", 
                    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        
        for m in range(1, 13):
            if m in monthly_avg_daily.index:
                avg_dni = monthly_avg_daily.loc[m, 'DNI']
                avg_ghi = monthly_avg_daily.loc[m, 'GHI']
                tot_dni = monthly_totals.loc[m, 'DNI']
                self.txt_output1.insert(tk.END, 
                    f"{mo_names[m-1]:<8} {avg_dni:>12.2f} {avg_ghi:>12.2f} {tot_dni:>14.2f}\n")

        # Add verification sum
        sum_monthly = sum(monthly_totals.loc[m, 'DNI'] for m in range(1,13) if m in monthly_totals.index)
        self.txt_output1.insert(tk.END, f"{'-'*60}\n")
        self.txt_output1.insert(tk.END, f"{'Sum (Jan-Dec):':<8} {'':>12} {'':>12} {sum_monthly:>14.2f}\n")

        self.txt_output1.insert(tk.END, f"\n{'='*60}\n")
        self.txt_output1.insert(tk.END, f"ANNUAL SUMMARY\n")
        self.txt_output1.insert(tk.END, f"{'-'*60}\n")
        self.txt_output1.insert(tk.END, f"Total DNI (Annual):       {annual_dni_total:>12.2f} kWh/m²/yr\n")
        self.txt_output1.insert(tk.END, f"Total GHI (Annual):       {annual_ghi_total:>12.2f} kWh/m²/yr\n")
        
        self.txt_output1.insert(tk.END, f"\n{'='*60}\n")
        self.txt_output1.insert(tk.END, f"DESIGN VALUE (for System Sizing)\n")
        self.txt_output1.insert(tk.END, f"{'-'*60}\n")
        self.txt_output1.insert(tk.END, f"Modal DNI Cluster:        {modal_cluster:>12.0f} W/m²\n")
        self.txt_output1.insert(tk.END, f"  (Most frequent DNI in dataset)\n")
        self.txt_output1.insert(tk.END, f"Dynamic Threshold:        {threshold:>12.0f} W/m²\n")
        self.txt_output1.insert(tk.END, f"  (93% of modal cluster)\n")
        
        # BOLD Design DNI - the most important value
        self.txt_output1.insert(tk.END, f"Design DNI:               ")
        self.txt_output1.insert(tk.END, f"{design_dni:>12.2f} W/m²", "bold")
        self.txt_output1.insert(tk.END, f"\n")
        
        self.txt_output1.insert(tk.END, f"  (Average during hours ≥{threshold} W/m²)\n")
        self.txt_output1.insert(tk.END, f"Productive Hours/Year:    {productive_count:>12d} hours\n")
        self.txt_output1.insert(tk.END, f"{'='*60}\n\n")
        self.txt_output1.insert(tk.END, f"✓ Analysis Complete!\n")
        
        self.btn_tech1.config(state="normal")

    def show_error_tech1(self, msg):
        messagebox.showerror("Tech 1 Error", msg)
        self.btn_tech1.config(state="normal")

    def generate_dni_graph(self):
        """Generate DNI distribution histogram to validate modal cluster"""
        if self.current_df is None:
            messagebox.showwarning("No Data", "Please run Tech 1 analysis first!")
            return
        
        try:
            # Create new window for graph
            graph_window = tk.Toplevel(self.root)
            graph_window.title("DNI Distribution Analysis")
            graph_window.geometry("1000x700")
            
            # Create figure with subplots
            fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 8))
            fig.suptitle('DNI Distribution Analysis - Modal Cluster Validation', 
                        fontsize=14, fontweight='bold')
            
            # Get DNI data
            dni_data = self.current_df['DNI'].values
            dni_nonzero = dni_data[dni_data > 0]  # Exclude nighttime
            
            # GRAPH 1: Histogram with bins 0-100, 100-200, etc.
            bins = range(100, 1000, 50)  # 100-150, 150-200, ..., 950-1000
            counts, edges, patches = ax1.hist(dni_nonzero, bins=bins, edgecolor='black',
                                             color='skyblue', alpha=0.7)
            
            # Highlight the modal cluster bin
            for i in range(len(edges) - 1):
                if edges[i] <= self.current_modal < edges[i + 1]:
                    patches[i-1].set_facecolor('orange')
                    patches[i-1].set_label(f'Modal Cluster: {self.current_modal} W/m²')
                    break
            #modal_bin_index = self.current_modal // 100
            #if modal_bin_index < len(patches):
                #patches[modal_bin_index].set_facecolor('orange')
                #patches[modal_bin_index].set_label(f'Modal Cluster: {self.current_modal} W/m²')
            
            # Add threshold line
            #ax1.axvline(self.current_threshold, color='red', linestyle='--',
                       #linewidth=2, label=f'Threshold: {self.current_threshold} W/m²')
            
            ax1.set_xlabel('DNI (W/m²)', fontsize=11, fontweight='bold')
            ax1.set_ylabel('Frequency (Hours)', fontsize=11, fontweight='bold')
            ax1.set_title('DNI Frequency Distribution (50 W/m² bins)', fontsize=12)
            ax1.legend(loc='upper right')
            ax1.grid(True, alpha=0.3)
            
            # Add value labels on bars
            for i, (count, edge) in enumerate(zip(counts, edges[:-1])):
                if count > 0:
                    ax1.text(edge + 25, count, f'{int(count)}',
                           ha='center', va='bottom', fontsize=8)
            
            # GRAPH 2: Fine-grained histogram for productive hours (DNI > 500)
            dni_productive = dni_data[dni_data > 650]
            bins_fine = range(500, 1100, 50)  # 500-550, 550-600, etc.
            counts2, edges2, patches2 = ax2.hist(dni_productive, bins=bins_fine, 
                                                edgecolor='black', color='lightgreen', alpha=0.7)
            
            # Highlight modal cluster
            for i, patch in enumerate(patches2):
                if edges2[i] <= self.current_modal < edges2[i+1]:
                    patch.set_facecolor('orange')
                    patch.set_label(f'Modal: {self.current_modal} W/m²')
            
            # Add threshold line
            ax2.axvline(self.current_threshold, color='red', linestyle='--', 
                       linewidth=2, label=f'Threshold: {self.current_threshold} W/m²')
            
            ax2.set_xlabel('DNI (W/m²)', fontsize=11, fontweight='bold')
            ax2.set_ylabel('Frequency (Hours)', fontsize=11, fontweight='bold')
            ax2.set_title('Productive Hours Distribution (50 W/m² bins, DNI > 500)', fontsize=12)
            ax2.legend(loc='upper right')
            ax2.grid(True, alpha=0.3)
            
            # Add statistics text
            stats_text = f"""
Statistics:
• Total hours: {len(dni_data)}
• Hours with DNI > 0: {len(dni_nonzero)} ({len(dni_nonzero)/len(dni_data)*100:.1f}%)
• Hours with DNI > 500: {len(dni_productive)} ({len(dni_productive)/len(dni_data)*100:.1f}%)
• Modal Cluster: {self.current_modal} W/m²
• Dynamic Threshold: {self.current_threshold} W/m²
• Peak DNI: {dni_data.max():.0f} W/m²
            """
            fig.text(0.02, 0.02, stats_text, fontsize=9, family='monospace',
                    bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
            
            plt.tight_layout(rect=[0, 0.12, 1, 0.96])
            
            # Embed in tkinter window
            canvas = FigureCanvasTkAgg(fig, master=graph_window)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            # Add save button
            save_btn = tk.Button(graph_window, text="💾 Save Graph as PNG", 
                               command=lambda: self.save_graph(fig),
                               bg="#4CAF50", fg="white", font=('Arial', 10, 'bold'))
            save_btn.pack(pady=10)
            
            messagebox.showinfo("Graph Generated", 
                              "DNI distribution graph generated successfully!\n\n"
                              "The orange bars show the modal cluster.\n"
                              "The red line shows the threshold (650 W/m²).")
            
        except Exception as e:
            messagebox.showerror("Graph Error", f"Error generating graph: {str(e)}")
    
    def save_graph(self, fig):
        """Save the matplotlib figure as PNG"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG files", "*.png"), ("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            fig.savefig(file_path, dpi=300, bbox_inches='tight')
            messagebox.showinfo("Success", f"Graph saved to:\n{file_path}")

    # ========================================================================
    # TECHNIQUE 2: Mathematical Model
    # ========================================================================

    def calculate_irradiance(self, row, lat_deg, lon_deg, beta_deg, gamma_deg, rho):
        """Calculate irradiance using mathematical models"""
        try:
            # Constants
            days_in_months = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

            # Coordinate Conversion
            lat = lat_deg * math.pi / 180
            beta = beta_deg * math.pi / 180
            gamma = gamma_deg * math.pi / 180

            # Date Parsing
            day = int(str(row['Day']))
            month = int(str(row['Month']))
            n = day + sum(days_in_months[:month])

            # Time Parsing
            hour = int(str(row['Hour']))
            minute = int(str(row['Minute']))
            tm = ((hour + minute / 60) - 12) * 60

            # Solar Geometry
            B = (n - 1) * (360 / 365) * math.pi / 180
            E = 229.18 * (0.000075 + 0.001868 * math.cos(B) - 0.032077 * math.sin(B) - 
                         0.014615 * math.cos(2 * B) - 0.04089 * math.sin(2 * B))
            LAT = tm - 4 * (82.5 - lon_deg) + E
            omega = (LAT / 4) * math.pi / 180
            delta = (23.45 * math.sin(((360 / 365) * (284 + n)) * math.pi / 180)) * math.pi / 180

            cos_theta = (math.sin(beta) * math.cos(gamma) * 
                        (math.cos(delta) * math.cos(omega) * math.sin(lat) - 
                         math.sin(delta) * math.cos(lat)) +
                        math.sin(beta) * math.sin(gamma) * math.cos(delta) * math.sin(omega) +
                        math.cos(beta) * (math.cos(delta) * math.cos(omega) * math.cos(lat) + 
                                         math.sin(delta) * math.sin(lat)))

            cos_theta_z = (math.cos(delta) * math.cos(omega) * math.cos(lat) + 
                          math.sin(delta) * math.sin(lat))

            # Prevent division by zero
            cos_theta_z = max(cos_theta_z, 0.0001)

            rb = cos_theta / cos_theta_z
            rd = (1 + math.cos(beta)) / 2
            rr = (rho * (1 - math.cos(beta))) / 2

            # Irradiance Logic
            Ig = float(row['Ig'])
            if self.optn1.get() == 'A':
                Id = float(row['Id'])
                Ib = Ig - Id
                Ibn = Ib / cos_theta_z
            else:
                Ibn = float(row['Ibn'])
                Ib = Ibn * cos_theta_z
                Id = Ig - Ib

            # Models
            if self.optn2.get() == 'C':
                It = (Ib * rb) + (Id * rd) + (Ig * rr)
            else:
                Isc = 1367 * (1 + 0.033 * math.cos(((360 * n) / 365) * math.pi / 180))
                Io = Isc * cos_theta_z
                Ai = Ib / Io if Io > 0 else 0
                f = (Ib / Ig) ** 0.5 if Ig > 0 else 0
                It = ((Ib + Id * Ai) * rb + Id * (1 - Ai) * rd * 
                     (1 + f * (math.sin(0.5 * beta)) ** 3) + Ig * rr)

            if self.optn1.get() == 'A':
                return pd.Series({'Ib': Ib, 'Ibn': Ibn, 'It': It})
            else:
                return pd.Series({'Ib': Ib, 'Id': Id, 'It': It})
                
        except Exception as e:
            return pd.Series({'Ib': 0, 'Ibn': 0, 'It': 0}) if self.optn1.get() == 'A' else pd.Series({'Ib': 0, 'Id': 0, 'It': 0})

    def run_tech2(self):
        try:
            if not self.file_path_tech2.get():
                raise ValueError("Please select an Excel or CSV file first.")

            # Read parameters
            params = {}
            for k, v in self.inputs.items():
                val = v.get().strip()
                if not val:
                    raise ValueError(f"Field {k} cannot be empty.")
                params[k] = float(val)

            # Read file (CSV or Excel)
            file_path = self.file_path_tech2.get()
            if file_path.lower().endswith('.csv'):
                # Try to detect if CSV has header rows (like NSRDB format)
                df_preview = pd.read_csv(file_path, nrows=5)
                skip = 2 if "Source" in str(df_preview.columns) else 0
                df = pd.read_csv(file_path, skiprows=skip)
            else:
                # Excel file
                df = pd.read_excel(file_path)

            # Clean column names
            df.columns = [c.strip() for c in df.columns]
            
            # CRITICAL FIX: Map NSRDB column names to expected names
            column_mapping = {
                'GHI': 'Ig',   # Global Horizontal Irradiance → Ig
                'DHI': 'Id',   # Diffuse Horizontal Irradiance → Id
                'DNI': 'Ibn'   # Direct Normal Irradiance → Ibn
            }
            
            # Apply mapping if NSRDB columns exist
            df.rename(columns=column_mapping, inplace=True)
            
            # Display column info for debugging
            self.root.after(0, lambda: self.txt_output2.insert(tk.END, 
                f"Columns found: {', '.join(df.columns[:10])}\n"))

            # Verify required columns
            required = ['Day', 'Month', 'Hour', 'Minute', 'Ig']
            if self.optn1.get() == 'A':
                required.append('Id')
            else:
                required.append('Ibn')

            missing = [col for col in required if col not in df.columns]
            if missing:
                raise ValueError(f"Missing required columns: {', '.join(missing)}\n"
                               f"Available columns: {', '.join(df.columns)}")

            # Process
            self.root.after(0, lambda: self.txt_output2.insert(tk.END, "Processing calculations...\n"))
            
            results = df.apply(self.calculate_irradiance, axis=1, args=(
                params['lat'], params['lon'], params['beta'], params['gamma'], params['rho']
            ))

            # Combine and save
            final_df = pd.concat([df, results], axis=1)
            
            # Ask for save location
            def save_results():
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx", 
                    filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
                )
                
                if save_path:
                    # Save based on chosen extension
                    if save_path.lower().endswith('.csv'):
                        final_df.to_csv(save_path, index=False)
                    else:
                        final_df.to_excel(save_path, index=False)
                    
                    self.txt_output2.insert(tk.END, f"\n✓ File saved successfully!\n")
                    self.txt_output2.insert(tk.END, f"Location: {save_path}\n")
                    self.txt_output2.insert(tk.END, f"Rows processed: {len(df)}\n")
                    self.txt_output2.insert(tk.END, f"Model used: {'Liu & Jordan' if self.optn2.get() == 'C' else 'HDKR'}\n")
                    self.txt_output2.insert(tk.END, f"Input type: {'Ig & Id' if self.optn1.get() == 'A' else 'Ig & Ibn'}\n")
                    messagebox.showinfo("Success", "Processing complete! File saved.")
                else:
                    self.txt_output2.insert(tk.END, "\n⚠ Save cancelled by user.\n")
                
                self.btn_tech2.config(state="normal")
            
            self.root.after(0, save_results)

        except Exception as error:
            error_msg = str(error)
            self.root.after(0, lambda msg=error_msg: self.show_error_tech2(msg))

    def show_error_tech2(self, msg):
        self.txt_output2.insert(tk.END, f"\n❌ Error: {msg}\n")
        messagebox.showerror("Tech 2 Error", msg)
        self.btn_tech2.config(state="normal")


if __name__ == "__main__":
    root = tk.Tk()
    app = SolarDesignApp(root)
    root.mainloop()
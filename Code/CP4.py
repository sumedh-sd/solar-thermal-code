import math
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


class SolarIrradianceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Solar Irradiance Calculator")
        self.root.geometry("500x650")
        self.file_path = tk.StringVar()
        self.create_widgets()

    def create_widgets(self):
        # File Selection
        tk.Label(self.root, text="Step 1: Select Excel File", font=('Arial', 10, 'bold')).pack(pady=10)
        file_frame = tk.Frame(self.root)
        file_frame.pack()
        tk.Entry(file_frame, textvariable=self.file_path, width=40).pack(side=tk.LEFT, padx=5)
        tk.Button(file_frame, text="Browse", command=self.browse_file).pack(side=tk.LEFT)

        # UPDATED: Use ttk.Separator instead of tk.Separator
        ttk.Separator(self.root, orient='horizontal').pack(fill='x', pady=10)

        # Numerical Inputs
        tk.Label(self.root, text="Step 2: Location Parameters", font=('Arial', 10, 'bold')).pack()

        self.inputs = {}
        field1 = [("Latitude (deg):", "lat_deg"), ("Longitude (deg):", "lon_deg")]

        for text, key in field1:
            frame = tk.Frame(self.root)
            frame.pack(pady=2)
            tk.Label(frame, text=text, width=25, anchor='w').pack(side=tk.LEFT)
            ent = tk.Entry(frame, width=15)
            ent.pack(side=tk.LEFT)
            self.inputs[key] = ent

        ttk.Separator(self.root, orient='horizontal').pack(fill='x', pady=10)

        tk.Label(self.root, text="Step 3: Collector Parameters", font=('Arial', 10, 'bold')).pack()

        field2 = [("Total collector area (m2):", "A_col")]

        for text, key in field2:
            frame = tk.Frame(self.root)
            frame.pack(pady=2)
            tk.Label(frame, text=text, width=40, anchor='w').pack(side=tk.LEFT)
            ent = tk.Entry(frame, width=15)
            ent.pack(side=tk.LEFT)
            self.inputs[key] = ent

        ttk.Separator(self.root, orient='horizontal').pack(fill='x', pady=10)

        tk.Label(self.root, text="Step 4: Plant Parameters", font=('Arial', 10, 'bold')).pack()

        field3 = [("Calorific value of fuel (kcal/kg):", "CV")]

        for text, key in field3:
            frame = tk.Frame(self.root)
            frame.pack(pady=2)
            tk.Label(frame, text=text, width=35, anchor='w').pack(side=tk.LEFT)
            ent = tk.Entry(frame, width=15)
            ent.pack(side=tk.LEFT)
            self.inputs[key] = ent

        ttk.Separator(self.root, orient='horizontal').pack(fill='x', pady=10)

        # Process Button
        tk.Button(self.root, text="Calculate & Save Results", bg="#4CAF50", fg="white",
                  font=('Arial', 11, 'bold'), command=self.process_data).pack(pady=30)

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.file_path.set(filename)

    def calculate_logic(self, row, lat_deg, lon_deg, A_col, CV):
        # Constants
        days_in_months = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

        # 1. Coordinate Conversion
        lat = lat_deg * math.pi / 180

        # 2. Date Parsing
        day = int(str(row['Day']))
        month = int(str(row['Month']))
        n = day + sum(days_in_months[:month])

        # 3. Time Parsing
        hour = int(str(row['Hour']))
        minute = int(str(row['Minute']))
        tm = ((hour + minute / 60) - 12) * 60

        # 4. Solar Geometry
        B = (n - 1) * (360 / 365) * math.pi / 180
        E = 229.18 * (0.000075 + 0.001868 * math.cos(B) - 0.032077 * math.sin(B) - 0.014615 * math.cos(2 * B) - 0.04089 * math.sin(2 * B))
        LAT = tm - 4 * (82.5 - lon_deg) + E
        omega = (LAT / 4) * math.pi / 180
        delta = (23.45 * math.sin(((360 / 365) * (284 + n)) * math.pi / 180)) * math.pi / 180
        cos_theta_z = math.cos(delta) * math.cos(omega) * math.cos(lat) + math.sin(delta) * math.sin(lat)
        if cos_theta_z < 0:
            cos_theta_z = 0
        cos_theta = (cos_theta_z**2 + (math.cos(delta))**2 * (math.sin(omega))**2)**(1/2)
        theta = math.degrees(math.acos(cos_theta))
        tan_theta = math.tan(theta*math.pi/180)

        # 5. Solar Collector
        IAM = 1 + ((0.000884 * theta)/cos_theta) - ((0.00005369 * (theta**2))/cos_theta)
        if IAM > 1:
            IAM = 1
        if IAM < 0:
            IAM = 0
        eta_col = 0.99 * 0.98 * 0.935 * 0.95
        eta_rec = 0.98 * 0.963 * 0.963 * 0.96
        eta_sha = (15 * cos_theta_z)/(5 * cos_theta)
        if eta_sha > 1:
            eta_sha = 1
        f_end = 1 - (1.49 * tan_theta)/49
        eta_sf = eta_col * eta_rec * eta_sha * f_end

        # 6. Fuel calculation
        DNI = float(row['DNI'])
        Ta = float(row['Temperature'])
        Tf = 250
        while True:
            Tm = (Tf + 150) / 2
            Q_sf = ((DNI * cos_theta * eta_sf * IAM) - 0.1 * (Tm - Ta)) * A_col * 0.001 # unit is kW
            if Q_sf < 0:
                Q_sf = 0
            if DNI == 0:
                eta_overall = 0
            else:
                eta_overall = (Q_sf * 1000) / (DNI * A_col)
            Cp_f = (0.000035 * Tf**2) - 0.00845 * Tf + 4.785
            Cp = (Cp_f + 4.3) / 2
            T_out = (Q_sf / (54.2 * Cp)) + 150
            error = abs(T_out - Tf)
            if error < 0.001:
                break
            else:
                Tf = T_out
        h_PTC_out = (Q_sf / 54.2) + 642.45
        h_boiler_inlet = 0.2 * h_PTC_out + 513.96
        Q_boiler = 271 * (3418.9 - h_boiler_inlet)
        mf_coal = Q_boiler / (CV * 4.2)

        return pd.Series({'Q (kW)':Q_sf, 'Fuel (kg/s)':mf_coal, 'Solar Field Efficiency':eta_sf, 'Overall Efficiency':eta_overall})

    def process_data(self):
        try:
            if not self.file_path.get():
                messagebox.showwarning("Input Error", "Please select an Excel file first.")
                return

            # Read parameters from UI
            params = {}
            for k, v in self.inputs.items():
                val = v.get().strip()
                if not val:
                    raise ValueError(f"Field {k} cannot be empty.")
                params[k] = float(val)

            # Read Excel
            df = pd.read_excel(self.file_path.get())

            # Verify required columns exist
            required = ['Month', 'Day', 'Hour', 'Minute', 'DNI']

            for col in required:
                if col not in df.columns:
                    raise ValueError(f"Missing required column: {col}")

            # Process
            results = df.apply(self.calculate_logic, axis=1, args=(params['lat_deg'], params['lon_deg'],
                                                                   params['A_col'],params['CV']))

            # Save
            final_df = pd.concat([df, results], axis=1)
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

            if save_path:
                final_df.to_excel(save_path, index=False)
                messagebox.showinfo("Success", "Processing complete! File saved.")

        except Exception as e:
            messagebox.showerror("System Error", f"Error details: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = SolarIrradianceApp(root)
    root.mainloop()
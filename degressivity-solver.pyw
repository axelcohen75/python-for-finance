### Download module if necessary
import subprocess
import sys

def install_package(package_name):
    try:
        __import__(package_name)
    except ImportError:
        print(f"{package_name} n'est pas installé. Installation en cours...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])

required_packages = ['pandas', 'tkinter', 'numpy']

for package in required_packages:
    install_package(package)

import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
from datetime import datetime, timedelta
import numpy as np

def degressivity_solver():
    try:
        ### Retrieve Input
        maturite_annees = int(entry_maturite.get())
        periodicite = combo_periodicite.get()  # Retrieved using drop-down menu
        target_degressivite = float(entry_target.get())
        periode_non_call = float(entry_non_call.get())
        barriere_initiale = float(entry_barriere.get())
        date_strike = datetime.strptime(entry_strike.get(), '%d/%m/%Y')

        ### Calculate maturity date
        date_maturite = date_strike + timedelta(days=maturite_annees * 365)
        
        ### Calculate non-call period in days
        jours_non_call = int(periode_non_call * 365)

        ### Define frequency
        if periodicite == 'Daily':
            delta = timedelta(days=1)
        elif periodicite == 'Monthly':
            delta = timedelta(days=30)
        elif periodicite == 'Quarterly':
            delta = timedelta(days=91)
        elif periodicite == 'Semiannually':
            delta = timedelta(days=182)
        else:
            delta = timedelta(days=365)

        ### Pandas DataFrame Creation
        dates = []
        barrieres = []
        current_date = date_strike
        nb_periodes = 0

        ### Calculate dates and knock out barriers
        while current_date <= date_maturite:
            if nb_periodes * delta.days < jours_non_call: # Able to starts degressivity after the non-call period
                barriere = barriere_initiale
            else:
                # Calculate new barrier for each date
                barriere = max(target_degressivite, barriere_initiale - (barriere_initiale - target_degressivite) * (nb_periodes * delta.days - jours_non_call) / (maturite_annees * 365 - jours_non_call))

            dates.append(current_date.strftime('%d/%m/%Y'))
            barrieres.append(f"{barriere:.2f}%")

            # Next period
            current_date += delta
            nb_periodes += 1

        # Force last barrier to target value
        dates[-1] = date_maturite.strftime('%d/%m/%Y')
        barrieres[-1] = f"{target_degressivite:.2f}%"

        # Create DataFrame with results
        global df_degressivite
        df_degressivite = pd.DataFrame({'Date': dates, 'Barrière Autocall (%)': barrieres})

        # Display results in text box
        text_result.delete(1.0, tk.END) # Clear text box
        text_result.insert(tk.END, f"Degressivity per period : {(barriere_initiale - target_degressivite) / (nb_periodes - (jours_non_call / delta.days)):.4f}%\n")
        afficher_dataframe(df_degressivite)

    except Exception as e:
        messagebox.showerror("Error", str(e))

# Display the DataFrame in the Treeview
def afficher_dataframe(df):
    # Clear treeview
    for row in tree.get_children():
        tree.delete(row)

    # Insert columns
    tree["column"] = list(df.columns)
    tree["show"] = "headings"
    for col in df.columns:
        tree.heading(col, text=col)

    # Insert row
    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        tree.insert("", "end", values=row)

# Export table to excel
def exporter_excel():
    try:
        df_degressivite.to_excel("result_degressivity.xlsx", index=False)
        messagebox.showinfo("Success", "Data was succesfuly exported to Excel")
    except Exception as e:
        messagebox.showerror(f"Export error : {str(e)}")

### GUI Configuration

# Global Interface
root = tk.Tk()
root.title("Degressivity Solver")
style = ttk.Style()
style.theme_use('clam') 

# Main Frame
frame_main = ttk.Frame(root, padding="10")
frame_main.grid(row=0, column=0, sticky="nsew")

# Create multiple sections
frame_inputs = ttk.LabelFrame(frame_main, text="Inputs", padding="10")
frame_inputs.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

frame_results = ttk.LabelFrame(frame_main, text="Results", padding="10")
frame_results.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

# Parameter entry widgets
ttk.Label(frame_inputs, text="Maturity (in years) :").grid(row=0, column=0, sticky="w")
entry_maturite = ttk.Entry(frame_inputs)
entry_maturite.grid(row=0, column=1)

ttk.Label(frame_inputs, text="Frequency :").grid(row=1, column=0, sticky="w")
combo_periodicite = ttk.Combobox(frame_inputs, values=["Daily", "Monthly", "Quarterly", "Semiannualy", "Annually"])
combo_periodicite.grid(row=1, column=1)
combo_periodicite.current(3)  # Set default value to Quarterly

ttk.Label(frame_inputs, text="Target at maturity (in %) :").grid(row=2, column=0, sticky="w")
entry_target = ttk.Entry(frame_inputs)
entry_target.grid(row=2, column=1)

ttk.Label(frame_inputs, text="Non-call period (in years) :").grid(row=3, column=0, sticky="w")
entry_non_call = ttk.Entry(frame_inputs)
entry_non_call.grid(row=3, column=1)

ttk.Label(frame_inputs, text="Autocall Barrier Start (%) :").grid(row=4, column=0, sticky="w")
entry_barriere = ttk.Entry(frame_inputs)
entry_barriere.grid(row=4, column=1)

ttk.Label(frame_inputs, text="Strike Date (DD/MM/YYYY) :").grid(row=5, column=0, sticky="w")
entry_strike = ttk.Entry(frame_inputs)
entry_strike.grid(row=5, column=1)

# Solve Button
ttk.Button(frame_inputs, text="Solve", command=degressivity_solver).grid(row=6, column=0, columnspan=2, pady=10)

# Text box to display results
text_result = tk.Text(frame_results, height=5, width=50)
text_result.grid(row=0, column=0, padx=10, pady=10)

# Tableau pour afficher le DataFrame
tree = ttk.Treeview(frame_results)
tree.grid(row=1, column=0, padx=10, pady=10)

# Table to display the DataFrame
ttk.Button(frame_results, text="Exporter en Excel", command=exporter_excel).grid(row=2, column=0, pady=10)

# Adjust window size to content
root.update_idletasks()
root.geometry(f"{root.winfo_width()}x{root.winfo_height()}")

# Freeze window size
root.resizable(False, False)

root.mainloop()

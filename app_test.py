# -*- coding: utf-8 -*-
"""
Created on Thu Mar 20 14:55:05 2025

@author: BENJAMIN
"""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors

def detect_separator(file_path):
    with open(file_path, 'r', encoding='ISO-8859-15') as f:
        first_line = f.readline()
    
    sep_candidates = ['|', '\t']
    best_sep = max(sep_candidates, key=lambda sep: first_line.count(sep))
    return best_sep

def select_file():
    file_path = filedialog.askopenfilename(title="Sélectionner un fichier FEC.txt", filetypes=[("Text Files", "*.txt")])
    if file_path:
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, file_path)
        load_compte_nums(file_path)
        btn_convert.config(state="normal")

def load_compte_nums(file_path):
    try:
        separator = detect_separator(file_path)
        df = pd.read_csv(file_path, sep=separator, header=0, skiprows=0, encoding="ISO-8859-15")
        
        df['CompteNum'] = df['CompteNum'].astype(str)
        comptes = sorted(df['CompteNum'].unique())
        
        compte_combobox['values'] = comptes
        compte_combobox.current(0)
        compte_combobox.grid(row=2, column=1, padx=10, pady=10)
    
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors du chargement des comptes : {e}")

def export_to_pdf(df_filtered):
    try:
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not save_path:
            return

        doc = SimpleDocTemplate(save_path, pagesize=letter)
        elements = []
        
        data = [df_filtered.columns.to_list()] + df_filtered.values.tolist()
        
        table = Table(data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        elements.append(table)
        doc.build(elements)
        messagebox.showinfo("Succès", f"Le fichier PDF a été enregistré : {save_path}")
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur est survenue lors de la conversion en PDF : {e}")

def convert_to_pdf():
    file_path = entry_file_path.get()
    compte_selected = compte_combobox.get()
    
    if not file_path:
        messagebox.showerror("Erreur", "Veuillez sélectionner un fichier.")
        return
    
    try:
        separator = detect_separator(file_path)
        df = pd.read_csv(file_path, sep=separator, header=0, skiprows=0, encoding="ISO-8859-15")    
        df['Debit'] = pd.to_numeric(df['Debit'].str.replace(',', '.'), errors='coerce').fillna(0)
        df['Credit'] = pd.to_numeric(df['Credit'].str.replace(',', '.'), errors='coerce').fillna(0)
        df['EcritureDate'] = pd.to_datetime(df['EcritureDate'].astype(str), format='%Y%m%d')
             
        df_filtered = df.copy()
        df_filtered['CompteNum'] = df_filtered['CompteNum'].astype(str)
        df_filtered = df_filtered[df_filtered['CompteNum'] == compte_selected]
        df_filtered = df_filtered[df_filtered['EcritureLet'].isna()]
        # df_filtered = df_filtered.groupby('CompteNum', as_index=False).agg({
        #         'CompteLib': 'first', 
        #         'Debit': 'sum',
        #         'Credit': 'sum'
        #     })
        df_filtered['Solde'] = df_filtered['Debit'] - df_filtered['Credit']
        df_filtered = df_filtered[['CompteNum', 'CompteLib', 'EcritureDate', 'Credit', 'Debit', 'Solde']]
        # Ajouter une ligne de total
        total_row = pd.DataFrame({
            'CompteNum': ['Total'],
            'CompteLib': [''],
            'EcritureDate': [''],
            'Debit': [df_filtered['Debit'].sum()],
            'Credit': [df_filtered['Credit'].sum()],
            'Solde': [df_filtered['Solde'].sum()]
        })
        df_filtered = pd.concat([df_filtered, total_row], ignore_index=True)
        
        export_to_pdf(df_filtered)
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur est survenue lors de la conversion : {e}")

root = tk.Tk()
root.title("Convertisseur FEC en PDF avec Filtrage")

entry_file_path = tk.Entry(root, width=50)
entry_file_path.grid(row=0, column=0, padx=10, pady=10)

btn_select_file = tk.Button(root, text="Sélectionner un fichier FEC", command=select_file)
btn_select_file.grid(row=0, column=1, padx=10, pady=10)

compte_combobox = ttk.Combobox(root, state="readonly")
compte_combobox.grid_forget()

btn_convert = tk.Button(root, text="Exporter en PDF", command=convert_to_pdf, state="disabled")
btn_convert.grid(row=3, column=0, columnspan=2, pady=20)

root.mainloop()
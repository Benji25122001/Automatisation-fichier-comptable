# -*- coding: utf-8 -*-
"""
Created on Wed Apr  9 14:45:41 2025

@author: BENJAMIN
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
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
        df = pd.read_csv(file_path, sep=separator, encoding="ISO-8859-15")
        df['CompteNum'] = df['CompteNum'].astype(str)
        comptes = sorted(df['CompteNum'].unique())

        compte_listbox.delete(0, tk.END)
        for compte in comptes:
            compte_listbox.insert(tk.END, compte)

        compte_listbox.grid(row=2, column=1, padx=10, pady=10, rowspan=5)
        scrollbar.grid(row=2, column=2, sticky="ns", rowspan=5, pady=10)

    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors du chargement des comptes : {e}")

def export_to_pdf(data):
    try:
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not save_path:
            return

        doc = SimpleDocTemplate(save_path, pagesize=letter, leftMargin=30, rightMargin=30)
        elements = []
        styles = getSampleStyleSheet()

        elements.append(Paragraph("Écritures des comptes sélectionnés", styles['Title']))
        elements.append(Spacer(1, 12))
        
        #####
        
        col_widths = [80, 200, 80, 80, 80, 80]  # Augmenter la largeur de la colonne 'CompteLib' (index 1)

        table_data = []
        
        # Ajouter l'en-tête
        table_data.append(['CompteNum', 'CompteLib', 'EcritureDate', 'Credit', 'Debit', 'Solde'])

        for row in data[1:]:  # Ignorer la ligne d'en-tête
            if isinstance(row[1], str) and len(row[1]) > 30:  # Si 'CompteLib' est trop long, convertir en Paragraph
                row[1] = Paragraph(row[1], styles['BodyText'])
            table_data.append(row)

        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        
        #####
        
        # table = Table(data, repeatRows=1, colWidths=[80, 120, 80, 80, 80, 80])

        style = TableStyle([
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ])

        for i, row in enumerate(data[1:], start=1):  # skip header
            if row[0].startswith("➤ Compte"):
                style.add('SPAN', (0, i), (-1, i))
                style.add('BACKGROUND', (0, i), (-1, i), colors.lightgrey)
                style.add('FONTNAME', (0, i), (-1, i), 'Helvetica-Bold')
                style.add('FONTSIZE', (0, i), (-1, i), 9)
                style.add('ALIGN', (0, i), (-1, i), 'LEFT')
            elif row[0].strip().lower() == "total":
                style.add('BACKGROUND', (0, i), (-1, i), colors.lightblue)
                style.add('FONTNAME', (0, i), (-1, i), 'Helvetica-Bold')
            elif all(cell == '' for cell in row):
                style.add('LINEBEFORE', (0, i), (-1, i), 0, colors.white)
                style.add('LINEAFTER', (0, i), (-1, i), 0, colors.white)
                # style.add('LINEABOVE', (0, i), (-1, i), 0, colors.white)
                # style.add('LINEBELOW', (0, i), (-1, i), 0, colors.white)

        table.setStyle(style)
        elements.append(table)
        doc.build(elements)

        messagebox.showinfo("Succès", f"Le fichier PDF a été enregistré : {save_path}")
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur est survenue lors de la création du PDF : {e}")

def convert_to_pdf():
    file_path = entry_file_path.get()
    selected_indices = compte_listbox.curselection()
    if not file_path or not selected_indices:
        messagebox.showerror("Erreur", "Veuillez sélectionner un fichier et au moins un numéro de compte.")
        return

    try:
        separator = detect_separator(file_path)
        df = pd.read_csv(file_path, sep=separator, encoding="ISO-8859-15")

        df['Debit'] = pd.to_numeric(df['Debit'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        df['Credit'] = pd.to_numeric(df['Credit'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        df['EcritureDate'] = pd.to_datetime(df['EcritureDate'].astype(str), format='%Y%m%d', errors='coerce')
        df['CompteNum'] = df['CompteNum'].astype(str)

        selected_comptes = [compte_listbox.get(i) for i in selected_indices]
        data_rows = []

        # Ajouter en-tête
        data_rows.append(['CompteNum', 'CompteLib', 'EcritureDate', 'Credit', 'Debit', 'Solde'])

        for compte in selected_comptes:
            df_filtered = df[(df['CompteNum'] == compte) & (df['EcritureLet'].isna())].copy()
            if df_filtered.empty:
                continue

            df_filtered['Solde'] = df_filtered['Debit'] - df_filtered['Credit']
            df_filtered = df_filtered[['CompteNum', 'CompteLib', 'EcritureDate', 'Credit', 'Debit', 'Solde']]

            compte_lib = df_filtered['CompteLib'].iloc[0]
            titre = f"➤ Compte {compte} – {compte_lib}"

            # Ligne de titre personnalisée
            data_rows.append([titre, '', '', '', '', ''])

            # Ajouter les lignes du compte
            for _, row in df_filtered.iterrows():
                data_rows.append([
                    row['CompteNum'],
                    row['CompteLib'],
                    row['EcritureDate'].strftime("%d/%m/%Y") if pd.notnull(row['EcritureDate']) else '',
                    f"{row['Credit']:.2f}",
                    f"{row['Debit']:.2f}",
                    f"{row['Solde']:.2f}",
                ])

            # Ligne de total
            data_rows.append([
                'Total',
                '',
                '',
                f"{df_filtered['Credit'].sum():.2f}",
                f"{df_filtered['Debit'].sum():.2f}",
                f"{(df_filtered['Debit'].sum() - df_filtered['Credit'].sum()):.2f}"
            ])

            # Ligne vide pour espacer
            data_rows.append(['', '', '', '', '', ''])

        export_to_pdf(data_rows)

    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur est survenue : {e}")

# Interface Tkinter
root = tk.Tk()
root.title("Convertisseur FEC en PDF - Multi Comptes + Scroll + Tableau unique")

entry_file_path = tk.Entry(root, width=50)
entry_file_path.grid(row=0, column=0, padx=10, pady=10)

btn_select_file = tk.Button(root, text="Sélectionner un fichier FEC", command=select_file)
btn_select_file.grid(row=0, column=1, padx=10, pady=10)

tk.Label(root, text="Sélectionner un ou plusieurs comptes :").grid(row=2, column=0, sticky="nw", padx=10)

frame_listbox = tk.Frame(root)
frame_listbox.grid(row=2, column=1, sticky="n")

compte_listbox = tk.Listbox(frame_listbox, selectmode='multiple', exportselection=0, height=15, width=25)
scrollbar = tk.Scrollbar(root, orient="vertical", command=compte_listbox.yview)
compte_listbox.config(yscrollcommand=scrollbar.set)

btn_convert = tk.Button(root, text="Exporter en PDF", command=convert_to_pdf, state="disabled")
btn_convert.grid(row=8, column=0, columnspan=2, pady=20)

root.mainloop()
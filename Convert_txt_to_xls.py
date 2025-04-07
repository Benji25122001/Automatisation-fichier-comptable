# -*- coding: utf-8 -*-
"""
Created on Thu Mar 13 15:04:35 2025

@author: BENJAMIN
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def detect_separator(file_path):
    """
    Détecte le séparateur le plus probable en testant '|' et '\t'
    """
    with open(file_path, 'r', encoding='ISO-8859-15') as f:
        first_line = f.readline()  # Lire la première ligne

    # Tester les séparateurs sur la première ligne
    sep_candidates = ['|', '\t']
    best_sep = max(sep_candidates, key=lambda sep: first_line.count(sep))

    return best_sep

# Fonction pour sélectionner un fichier et, si c'est un fichier Excel, afficher les feuilles disponibles
def select_file():
    file_type = file_type_var.get()  # Récupère la valeur sélectionnée dans le menu déroulant
    
    # Définir les types de fichiers disponibles dans la boîte de dialogue en fonction de la sélection
    if file_type == "Texte (.txt)":
        filetypes = [("Text Files", "*.txt")]
        file_path = filedialog.askopenfilename(title="Sélectionner un fichier", filetypes=filetypes)
        if file_path:
            entry_file_path.delete(0, tk.END)  # Efface le champ de texte
            entry_file_path.insert(0, file_path)  # Affiche le chemin du fichier sélectionné
            sheet_menu.grid_forget()  # Cache le menu déroulant des feuilles (non applicable pour TXT)
    elif file_type == "Excel (.xlsx)":
        filetypes = [("Excel Files", "*.xlsx")]
        file_path = filedialog.askopenfilename(title="Sélectionner un fichier", filetypes=filetypes)
        if file_path:
            entry_file_path.delete(0, tk.END)  # Efface le champ de texte
            entry_file_path.insert(0, file_path)  # Affiche le chemin du fichier sélectionné
            
            # Charger le fichier Excel pour récupérer les feuilles disponibles
            try:
                excel_file = pd.ExcelFile(file_path)
                sheets = excel_file.sheet_names  # Liste des noms de feuilles
                
                # Mettre à jour le menu déroulant pour choisir une feuille
                sheet_var.set(sheets[0])  # Définir la première feuille comme valeur par défaut
                sheet_menu['menu'].delete(0, 'end')  # Supprime les anciennes options
                
                # Ajouter toutes les feuilles au menu déroulant
                for sheet in sheets:
                    sheet_menu['menu'].add_command(label=sheet, command=tk._setit(sheet_var, sheet))
                
                sheet_menu.grid(row=1, column=1, padx=10, pady=10)  # Afficher le menu déroulant des feuilles
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors du chargement des feuilles Excel : {e}")
                return
    else:
        file_path = filedialog.askopenfilename(title="Sélectionner un fichier", filetypes=[("Tous les fichiers", "*.*")])
        if file_path:
            entry_file_path.delete(0, tk.END)  # Efface le champ de texte
            entry_file_path.insert(0, file_path)  # Affiche le chemin du fichier sélectionné
            sheet_menu.grid_forget()  # Cache le menu déroulant des feuilles (non applicable)
    btn_convert.config(state="normal")

# Fonction pour convertir le fichier en Excel
def convert_to_excel():
    file_path = entry_file_path.get()
    if not file_path:
        messagebox.showerror("Erreur", "Veuillez sélectionner un fichier.")
        return
    
    # Vérifie si le fichier existe
    if not os.path.exists(file_path):
        messagebox.showerror("Erreur", "Le fichier sélectionné n'existe pas.")
        return
    
    try:
        # Lire le fichier en fonction de son extension
        file_extension = os.path.splitext(file_path)[1].lower()
        
        if file_extension == ".txt":
            separator = detect_separator(file_path)
            df = pd.read_csv(file_path, sep=separator, header=0, skiprows=0, encoding="ISO-8859-15")
            df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
            df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
            df['EcritureDate'] = pd.to_datetime(df['EcritureDate'].astype(str), format='%Y%m%d')
        elif file_extension == ".xlsx":
            sheet_name = sheet_var.get()  # Récupère le nom de la feuille sélectionnée
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            messagebox.showerror("Erreur", "Format de fichier non pris en charge.")
            return
        
        # Demander à l'utilisateur où enregistrer le fichier Excel
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not save_path:
            return  # Si l'utilisateur annule, on quitte
        
        # Création d'un objet ExcelWriter pour ajouter plusieurs feuilles
        with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
            # Sauvegarder la feuille "FEC Importé"
            df.to_excel(writer, sheet_name='FEC Importé', index=False)
            
            # Exemple de tri pour ajouter une deuxième feuille "FEC Trié"
            df_sorted = df.copy()  # Remplacez 'Date' par la colonne que vous voulez trier
            df_sorted = df[df['EcritureLet'].isna()]
            df_sorted = df_sorted.groupby('CompteNum', as_index=False).agg({
                'CompteLib': 'first', 
                'Debit': 'sum',
                'Credit': 'sum'
            })
            df_sorted['Solde'] = df_sorted['Debit'] - df_sorted['Credit']
            df_sorted.to_excel(writer, sheet_name='Balance Générale', index=False)
        
        btn_convert.config(state="disabled")  # Désactiver le bouton
        
        messagebox.showinfo("Succès", f"Le fichier a été converti et sauvegardé à : {save_path}")
    
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur est survenue lors de la conversion : {e}")

# Création de l'interface Tkinter
root = tk.Tk()
root.title("Convertisseur FEC en Excel")

# Champ pour afficher le chemin du fichier sélectionné
entry_file_path = tk.Entry(root, width=50)
entry_file_path.grid(row=0, column=0, padx=10, pady=10)

# Menu déroulant pour sélectionner le type de fichier (Excel ou Texte)
file_type_var = tk.StringVar(root)
file_type_var.set("Texte (.txt)")  # Valeur par défaut (avant sélection)
file_type_options = ["Texte (.txt)", "Excel (.xlsx)"]

file_type_menu = tk.OptionMenu(root, file_type_var, *file_type_options)
file_type_menu.grid(row=1, column=0, padx=10, pady=10)

# Menu déroulant pour sélectionner la feuille (initialement vide)
sheet_var = tk.StringVar(root)
sheet_menu = tk.OptionMenu(root, sheet_var, [])
sheet_menu.grid_forget()  # Initialement, il est caché

# Bouton pour ouvrir le sélecteur de fichier
btn_select_file = tk.Button(root, text="Sélectionner un fichier", command=select_file)
btn_select_file.grid(row=0, column=1, padx=10, pady=10)

# Bouton pour convertir le fichier
btn_convert = tk.Button(root, text="Générer Balance Générale", command=convert_to_excel)
btn_convert.grid(row=2, column=0, columnspan=2, pady=20)

# Démarrer l'interface
root.mainloop()
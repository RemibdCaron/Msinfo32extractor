import os
import csv
import tkinter as tk
from tkinter import filedialog
import re

exports_folder = ""
save_path = ""

# Fonction pour obtenir le dossier contenant les exports
def get_exports_folder():
    global exports_folder
    exports_folder = filedialog.askdirectory(title="Sélectionnez le dossier contenant les exports MSInfo32")
    exports_folder_label.config(text=exports_folder)

# Fonction pour obtenir le chemin du fichier "Full_Export.csv"
def get_save_path():
    global save_path
    save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Fichiers CSV", "*.csv")], title="Enregistrer le fichier Full_Export.csv")
    save_path_label.config(text=save_path)

# Fonction pour obtenir les noms des fichiers txt à partir du dossier d'exports
def get_files_names_from(folder_path):
    names = []
    if os.path.exists(folder_path) and os.path.isdir(folder_path):
        files = os.listdir(folder_path)
        txt_files = [file for file in files if file.endswith(".txt")]
        for txt_file in txt_files:
            names.append(os.path.join(folder_path, txt_file))
    return names

# Fonction pour extraire des informations à partir d'un fichier txt
def extract_info_from_txt(file_path):
    try:
        with open(file_path, 'r', encoding="utf-16le") as file:
            lines = file.readlines()
        nom_utilisateur_line = next(x for x in lines if x.startswith('Utilisateur'))
        nom_utilisateur = nom_utilisateur_line.strip('Utilisateur').strip()
        # Diviser le nom d'utilisateur en deux parties, en supprimant "\"
        nom, utilisateur = nom_utilisateur.split("\\", 1)
        version_windows = next(x for x in lines if x.startswith('Nom du système d’exploitation')).strip('Nom du système d’exploitation').strip()
        modele = next(x for x in lines if x.startswith('Modèle')).strip('Modèle').strip()
        processeur = next(x for x in lines if x.startswith('Processeur')).strip('Processeur').strip()
        ram = next(x for x in lines if x.startswith('Mémoire physique (RAM) installée')).strip('Mémoire physique (RAM) installée').strip()
        # Utilisation de l'expression régulière pour extraire la date
        bios_date_line = next(x for x in lines if x.startswith('Version du BIOS/Date'))
        date_match = re.search(r'\d{2}/\d{2}/\d{4}', bios_date_line)
        version_bios = date_match.group() if date_match else ""
        espace_disque = next(x for x in lines if x.startswith('Capacité')).strip('Capacité').strip().split('(')[0]
        return {
            'Nom': nom,
            'Utilisateur': utilisateur,
            'Version_Windows': version_windows,
            'Modele': modele,
            'Processeur': processeur,
            'Ram': ram,
            'Date_Bios': version_bios,
            'Espace_Disque': espace_disque
        }
    except FileNotFoundError:
        return {'error': f"Le fichier '{file_path}' n'a pas été trouvé."}
    except Exception as e:
        return {'error': f"Une erreur s'est produite : {str(e)}"}

# Fonction principale pour créer le fichier CSV
def create_csv():
    Exports_Paths_Name = get_files_names_from(exports_folder)

    with open(save_path, 'w', newline='', encoding='utf-16le') as csvfile:
        fieldnames = ['Nom', 'Utilisateur', 'Version_Windows', 'Modele', 'Processeur', 'Ram', 'Date_Bios', 'Espace_Disque']
        writer = csv.DictWriter(csvfile, delimiter=';', fieldnames=fieldnames)
        writer.writeheader()
        for path in Exports_Paths_Name:
            info = extract_info_from_txt(path)
            if 'error' in info:
                error_label.config(text=info['error'])
            else:
                writer.writerow(info)
    
    success_label.config(text="Export terminé.")
    result_path_label.config(text=f"Le fichier Full_Export.csv a été enregistré à l'emplacement : {save_path}")
    error_label.config(text="")  # Réinitialise le texte de l'erreur à vide

# Interface utilisateur avec Tkinter
root = tk.Tk()
root.title("Extracteur d'informations MSInfo32")

# Libellé et bouton pour le dossier d'exports
exports_label = tk.Label(root, text="Étape 1: Sélectionnez le dossier contenant les exports MSInfo32")
exports_label.pack(pady=10)  # Ajoute une marge verticale de 10 pixels
exports_button = tk.Button(root, text="Sélectionner le dossier", command=get_exports_folder)
exports_button.pack(pady=5)  # Ajoute une marge verticale de 5 pixels
exports_folder_label = tk.Label(root, text="", wraplength=400)
exports_folder_label.pack(pady=5)  # Ajoute une marge verticale de 5 pixels

# Libellé et bouton pour l'emplacement de sauvegarde
save_label = tk.Label(root, text="Étape 2: Sélectionnez l'emplacement où vous souhaitez enregistrer le fichier Full_Export.csv")
save_label.pack(pady=10)  # Ajoute une marge verticale de 10 pixels
save_button = tk.Button(root, text="Sélectionner l'emplacement", command=get_save_path)
save_button.pack(pady=5)  # Ajoute une marge verticale de 5 pixels
save_path_label = tk.Label(root, text="", wraplength=400)
save_path_label.pack(pady=5)  # Ajoute une marge verticale de 5 pixels

# Bouton pour exécuter le programme
run_button = tk.Button(root, text="Exécuter", command=create_csv)
run_button.pack(pady=10)  # Ajoute une marge verticale de 10 pixels

# Libellé pour afficher le résultat
success_label = tk.Label(root, text="", fg="green")
success_label.pack(pady=5)  # Ajoute une marge verticale de 5 pixels
error_label = tk.Label(root, text="", fg="red")
error_label.pack(pady=5)  # Ajoute une marge verticale de 5 pixels
result_path_label = tk.Label(root, text="."*170)
result_path_label.pack(pady=20)  # Ajoute une marge verticale de 20 pixels

# Placement des éléments dans la fenêtre
exports_label.pack()
exports_button.pack()
exports_folder_label.pack()
save_label.pack()
save_button.pack()
save_path_label.pack()
run_button.pack()
success_label.pack()
error_label.pack()
result_path_label.pack()

root.mainloop()
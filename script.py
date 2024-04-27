import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re
import csv

sorted_colors = None
file_inserted = False
export_button = None
button_frame = None
hue_columns_visible = False
rgb_columns_visible = False
colors_frame = None
results_displayed = False
hue_button = None
rgb_button = None
rgb_column_text = None
file_path = None
scrollbar = None
saturation_columns_visible = False
export_button = None
code_nom_columns_visible = False
export_data = []

def hex_to_rgb(hex_color):
    # Assurez-vous que hex_color est une chaîne de caractères
    if isinstance(hex_color, str):
        # Retirez le '#' s'il est présent
        hex_color = hex_color.lstrip('#')
        # Convertissez le code hexadécimal en valeurs R, G, B
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return r, g, b
    else:
        return None 

def calculate_required_width():
    global sorted_colors, hue_columns_visible, rgb_columns_visible, saturation_columns_visible

    num_columns = 1

    if hue_columns_visible:
        num_columns += 1

    if rgb_columns_visible:
        num_columns += 1

    if saturation_columns_visible:
        num_columns += 1

    column_width = 200
    required_width = max(num_columns * column_width, 1200)

    return required_width

def calculate_saturation_percentage(color):
    r, g, b = 0, 0, 0  # Initialisation des valeurs par défaut

    if isinstance(color, str):
        r, g, b = hex_to_rgb(color)

    max_val = max(r, g, b)
    min_val = min(r, g, b)
    if max_val == 0:
        return 0
    return ((max_val - min_val) / max_val) * 100

def sort_colors(colors):
    def hue(color):
        r, g, b = hex_to_rgb(color)
        max_val = max(r, g, b)
        min_val = min(r, g, b)
        delta = max_val - min_val

        if max_val == r:
            hue = (g - b) / delta if delta != 0 else 0
        elif max_val == g:
            hue = 2 + (b - r) / delta if delta != 0 else 0
        else:
            hue = 4 + (r - g) / delta if delta != 0 else 0

        hue *= 60

        if hue < 0:
            hue += 360

        return hue

    def calculate_brightness(color):
        r, g, b = hex_to_rgb(color)
        return (r * 299 + g * 587 + b * 114) / 1000

    categories = {
        'Rouge': [],
        'Orange': [],
        'Marron': [],
        'Jaune': [],
        'Vert': [],
        'Bleu': [],
        'Violet': [],
        'Rose': [],
        'Nuances de Gris': [],
    }

    for color in colors:
        h = hue(color)
        sat = calculate_saturation_percentage(color)

        if (0 <= h < 12) or (345 <= h <= 360):
            categories['Rouge'].append(color)
        elif (12 <= h < 26):
            categories['Orange'].append(color)
        elif (26 <= h < 40):
            categories['Marron'].append(color)
        elif (40 <= h < 56):
            categories['Jaune'].append(color)
        elif (56 <= h < 180):
            categories['Vert'].append(color)
        elif (180 <= h < 230):
            categories['Bleu'].append(color)
        elif (230 <= h < 315):
            categories['Violet'].append(color)
        elif (315 <= h < 345):
            categories['Rose'].append(color)

        if 2 <= sat <= 7:
            categories['Nuances de Gris'].append(color)

    for category, color_group in categories.items():
        # Triez d'abord par luminosité, puis par teinte pour les couleurs de même luminosité
        categories[category] = sorted(color_group, key=lambda x: calculate_brightness(x), reverse=True)

    return categories

def open_excel_file():
    global sorted_colors, file_inserted, file_path, code_nom_button, export_data  # Ajoutez export_data ici
    file_path = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx *.xls")])
    if file_path:
        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            messagebox.showerror("Erreur de lecture du fichier Excel", str(e))
            return

        if 'Code' in df.columns and 'Nom' in df.columns and 'Couleur' in df.columns:
            valid_colors = df['Couleur'].apply(validate_color).dropna()
            df['HUE'] = df['Couleur'].apply(calculate_hue)
            df['RGB'] = df['Couleur'].apply(calculate_rgb)
            df['% Saturation'] = df['Couleur'].apply(calculate_saturation_percentage)

            sorted_colors = sort_colors(valid_colors.tolist())

            # Mettez à jour le statut du bouton "Afficher Code+Nom" ici
            code_nom_button.config(state="normal")

            # Ajoutez les données à exporter
            export_data = []
            for category in sorted_colors:
                colors = sorted_colors[category]

                for color in colors:
                    match = df[df['Couleur'] == color]
                    if not match.empty:
                        code = match.iloc[0]['Code']
                        nom = match.iloc[0]['Nom']
                        hue = match.iloc[0]['HUE']
                        saturation = match.iloc[0]['% Saturation']
                        export_data.append([code, nom, color, hue, saturation])  # Ajoutez le pourcentage de saturation aux données exportées

            file_inserted = True
            display_colors()
        else:
            messagebox.showerror("Colonnes manquantes", "Le fichier Excel doit contenir les colonnes 'Code', 'Nom' et 'Couleur'.")

def export_to_csv():
    global export_data
    if not export_data:
        messagebox.showinfo("Aucune donnée à exporter", "Aucune donnée à exporter pour le moment.")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Fichiers CSV", "*.csv")])
    if file_path:
        try:
            with open(file_path, 'w', newline='') as csvfile:
                csvwriter = csv.writer(csvfile)
                # Write the header row
                csvwriter.writerow(["Code", "Nom", "Couleur", "HUE", "% deSaturation"])  # Ajoutez "% deSaturation" au header
                # Write the data
                csvwriter.writerows(export_data)
            messagebox.showinfo("Export terminé", "Les données ont été exportées au format CSV avec succès.")
        except Exception as e:
            messagebox.showerror("Erreur lors de l'exportation", f"Une erreur s'est produite lors de l'exportation : {str(e)}")

def export_to_txt():
    global sorted_colors
    if not file_inserted:
        messagebox.showerror("Aucun fichier inséré", "Veuillez insérer un fichier avant d'exporter les couleurs.")
        return
    file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Fichiers texte", "*.txt")])
    if file_path:
        with open(file_path, 'w') as file:
            for category, color_group in sorted_colors.items():
                file.write(f"{category}:\n")
                for color in color_group:
                    file.write(f"{color}\n")
        messagebox.showinfo("Export terminé", "Les couleurs triées ont été exportées avec succès en fichier texte.")

def create_export_buttons():
    global export_button, export_txt_button  # Assurez-vous que ces variables sont globales
    export_button = tk.Button(button_frame, text="Exporter CSV", command=export_to_csv, state="disabled")
    export_button.grid(row=0, column=6, padx=5)

    export_txt_button = tk.Button(button_frame, text="Exporter TXT", command=export_to_txt, state="disabled")
    export_txt_button.grid(row=0, column=7, padx=5)

    quit_button = tk.Button(button_frame, text="Quitter", command=root.quit)
    quit_button.grid(row=0, column=8, padx=5)  # Assurez-vous que le bouton "Quitter" est bien créé et ajouté

def create_rgb_column_text():
    global rgb_column_text
    rgb_column_text = tk.StringVar()
    if rgb_columns_visible:
        rgb_column_text.set("Cacher RGB")
    else:
        rgb_column_text.set("Afficher RGB")

def calculate_rgb(color):
    if isinstance(color, str):
        r, g, b = hex_to_rgb(color)
    else:
        r, g, b = color, color, color  # Utilisation de la valeur entière pour R, G et B
    return f"RGB({r}, {g}, {b})"

def close_colors_window():
    global colors_frame
    if colors_frame:
        colors_frame.destroy()

def validate_color(color):
    if isinstance(color, str):
        hex_color = color.lstrip('#')
        if re.match(r'^(?:[0-9A-Fa-f]{3}){1,2}$', hex_color) or re.match(r'^#[0-9A-Fa-f]{3}(?:[0-9A-Fa-f]{3})?$', color):
            return '#' + hex_color
    return None

def export_colors(sorted_colors):
    if not file_inserted:
        messagebox.showerror("Aucun fichier inséré", "Veuillez insérer un fichier avant d'exporter les couleurs.")
        return
    file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Fichiers texte", "*.txt")])
    if file_path:
        with open(file_path, 'w') as file:
            for category, color_group in sorted_colors.items():
                file.write(f"{category}:\n")
                for color in color_group:
                    file.write(f"{color}\n")
        messagebox.showinfo("Export terminé", "Les couleurs triées ont été exportées avec succès.")

def calculate_hue(color):
    if isinstance(color, str):
        r, g, b = hex_to_rgb(color)
    else:
        r, g, b = color, color, color  # Utilisation de la valeur entière pour R, G et B
    max_val = max(r, g, b)
    min_val = min(r, g, b)
    delta = max_val - min_val

    if max_val == r:
        hue = (g - b) / delta if delta != 0 else 0
    elif max_val == g:
        hue = 2 + (b - r) / delta if delta != 0 else 0
    else:
        hue = 4 + (r - g) / delta if delta != 0 else 0

    hue *= 60

    if hue < 0:
        hue += 360

    return hue

def set_main_window_size():
    global root, hue_columns_visible, rgb_columns_visible, saturation_columns_visible

    num_columns = 1

    if hue_columns_visible:
        num_columns += 1

    if rgb_columns_visible:
        num_columns += 1

    if saturation_columns_visible:
        num_columns += 1

    column_width = 200
    required_width = max(num_columns * column_width, 1080)
    height = 1000

    root.geometry(f"{required_width}x{height}")

def toggle_hue_columns():
    global hue_columns_visible
    hue_columns_visible = not hue_columns_visible

    display_colors()
    if hue_columns_visible:
        hue_button.config(text="Cacher HUE")
    else:
        hue_button.config(text="Afficher HUE")

def toggle_rgb_columns():
    global rgb_columns_visible
    rgb_columns_visible = not rgb_columns_visible

    display_colors()
    if rgb_columns_visible:
        rgb_button.config(text="Cacher les colonnes RGB")
    else:
        rgb_button.config(text="Afficher RGB")

def toggle_saturation_columns():
    global saturation_columns_visible, colors_frame, sorted_colors

    saturation_columns_visible = not saturation_columns_visible

    if colors_frame:
        colors_frame.destroy()

    display_colors()

    if saturation_columns_visible:
        saturation_button.config(text="Cacher % Saturation")
    else:
        saturation_button.config(text="Afficher % Saturation")

def toggle_code_nom_columns():
    global code_nom_columns_visible

    # Inversez l'état actuel (visible ou non)
    code_nom_columns_visible = not code_nom_columns_visible

    # Mettez à jour l'affichage des colonnes en fonction de l'état
    display_colors()

    # Mettez à jour le texte du bouton en conséquence
    if code_nom_columns_visible:
        code_nom_button.config(text="Cacher Code+Nom")
    else:
        code_nom_button.config(text="Afficher Code+Nom")

def display_colors():
    global colors_frame, canvas, file_inserted, export_button, hue_button
    global hue_columns_visible, rgb_columns_visible, sorted_colors, file_path, rgb_column_text, saturation_button, code_nom_columns_visible

    if colors_frame:
        colors_frame.destroy()

    if not file_inserted:
        return

    create_rgb_column_text()

    # Lire le fichier Excel
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        messagebox.showerror("Erreur de lecture du fichier Excel", f"Une erreur s'est produite lors de la lecture du fichier Excel : {str(e)}")
        return

    valid_colors = []

    # Traitement des données Excel
    if 'Code' in df.columns and 'Nom' in df.columns and 'Couleur' in df.columns:
        valid_colors = df['Couleur'].apply(validate_color).dropna()
        df['HUE'] = df['Couleur'].apply(calculate_hue)
        df['RGB'] = df['Couleur'].apply(calculate_rgb)
        df['% Saturation'] = df['Couleur'].apply(calculate_saturation_percentage)

        sorted_colors = sort_colors(valid_colors.tolist())
    else:
        messagebox.showerror("Colonnes manquantes", "Le fichier Excel doit contenir les colonnes 'Code', 'Nom' et 'Couleur'.")

    colors_frame = tk.Frame(root)
    colors_frame.pack(fill="both", expand=True)

    display_frame = tk.Frame(colors_frame)
    display_frame.pack(fill="both", expand=True)

    # Créez les barres de défilement horizontales et verticales
    x_scrollbar = tk.Scrollbar(display_frame, orient="horizontal")
    y_scrollbar = tk.Scrollbar(display_frame, orient="vertical")

    canvas = tk.Canvas(display_frame, xscrollcommand=x_scrollbar.set, yscrollcommand=y_scrollbar.set)
    frame = tk.Frame(canvas)

    x_scrollbar.config(command=canvas.xview)
    y_scrollbar.config(command=canvas.yview)

    # Placez les barres de défilement
    x_scrollbar.pack(side="bottom", fill="x")
    y_scrollbar.pack(side="right", fill="y")

    canvas.pack(side="left", fill="both", expand=True)
    canvas.create_window((0, 0), window=frame, anchor="nw")

    for i, (category, color_group) in enumerate(sorted_colors.items()):
        code_column_label = tk.Label(frame, text="Code", font=("Helvetica", 12, "bold"), width=10)
        code_column_label.grid(row=0, column=i * 11, padx=5, pady=2, sticky='nsew')
        code_column_label.grid_remove()  # Masquez la colonne par défaut si nécessaire

        nom_column_label = tk.Label(frame, text="Nom", font=("Helvetica", 12, "bold"), width=10)
        nom_column_label.grid(row=0, column=i * 11 + 1, padx=5, pady=2, sticky='nsew')
        nom_column_label.grid_remove()  # Masquez la colonne par défaut si nécessaire

        hue_column_label = tk.Label(frame, text=f"{category}", font=("Helvetica", 12, "bold"), width=8)
        hue_column_label.grid(row=0, column=i * 11 + 2, padx=5, pady=2, sticky='nsew')

        if hue_columns_visible:
            hue_column_label = tk.Label(frame, text=f"HUE", font=("Helvetica", 12, "bold"), width=5)
            hue_column_label.grid(row=0, column=i * 11 + 3, padx=5, pady=2, sticky='nsew')

        if rgb_columns_visible:
            rgb_column_label = tk.Label(frame, text=f"RGB", font=("Helvetica", 12, "bold"), width=15)
            rgb_column_label.grid(row=0, column=i * 11 + 4, padx=5, pady=2, sticky='nsew')

        if saturation_columns_visible:
            saturation_column_label = tk.Label(frame, text="% Saturation", font=("Helvetica", 12, "bold"), width=8)
            saturation_column_label.grid(row=0, column=i * 11 + 5, padx=5, pady=2, sticky='nsew')

        for j, color in enumerate(color_group):
            code_value = df[df['Couleur'] == color]['Code'].values
            nom_value = df[df['Couleur'] == color]['Nom'].values

            code_label_text = code_value[0] if code_value.size > 0 else ''
            code_label = tk.Label(frame, text=code_label_text, width=10)
            code_label.grid(row=j + 1, column=i * 11, padx=5, pady=2, sticky='nsew')
            code_label.grid_remove()  # Masquez la colonne par défaut si nécessaire

            nom_label_text = nom_value[0] if nom_value.size > 0 else ''
            nom_label = tk.Label(frame, text=nom_label_text, width=10)
            nom_label.grid(row=j + 1, column=i * 11 + 1, padx=5, pady=2, sticky='nsew')
            nom_label.grid_remove()  # Masquez la colonne par défaut si nécessaire

            hue_column_label = tk.Label(frame, text=color, bg=color, width=8)
            hue_column_label.grid(row=j + 1, column=i * 11 + 2, padx=5, pady=2, sticky='nsew')

            if hue_columns_visible:
                hue_value = calculate_hue(color)
                hue_label = tk.Label(frame, text=f"{hue_value:.2f}", width=5)
                hue_label.grid(row=j + 1, column=i * 11 + 3, padx=5, pady=2, sticky='nsew')

            if rgb_columns_visible:
                rgb_value = calculate_rgb(color)
                rgb_label = tk.Label(frame, text=rgb_value, width=15)
                rgb_label.grid(row=j + 1, column=i * 11 + 4, padx=5, pady=2, sticky='nsew')

            if saturation_columns_visible:
                saturation_value = calculate_saturation_percentage(color)
                saturation_label = tk.Label(frame, text=f"{saturation_value:.2f}%", width=8)
                saturation_label.grid(row=j + 1, column=i * 11 + 5, padx=5, pady=2, sticky='nsew')

            if code_nom_columns_visible:
                code_column_label.grid()
                nom_column_label.grid()
                code_label.grid()
                nom_label.grid()

    canvas.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))

    hue_button.config(state="normal")
    rgb_button.config(state="normal")
    saturation_button.config(state="normal")

    if sorted_colors:
        export_button.config(state="normal")
        export_txt_button.config(state="normal")

    if rgb_columns_visible:
        rgb_column_text.set("Cacher RGB")
    else:
        rgb_column_text.set("Afficher RGB")

def open_file_dialog():
    global file_inserted, file_path
    file_path = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx *.xls")])
    if not file_path:
        return
    file_inserted = True
    display_colors()

def main():
    global root, file_path, export_button, hue_button, rgb_button, button_frame, saturation_button, code_nom_button

    root = tk.Tk()
    root.title("Tri de couleurs by BreakingTech")

    set_main_window_size() 

    label = tk.Label(root, text="Cliquez sur Parcourir et choisissez un fichier Excel contenant vos données.")
    label.pack(pady=5)

    button_frame = tk.Frame(root)
    button_frame.pack()

    import_excel_button = tk.Button(button_frame, text="Import Excel", command=open_excel_file)
    import_excel_button.grid(row=0, column=1, padx=5)

    hue_button = tk.Button(button_frame, text="Afficher HUE", command=toggle_hue_columns, state="disabled")
    hue_button.grid(row=0, column=2, padx=5)

    rgb_button = tk.Button(button_frame, text="Afficher RGB", command=toggle_rgb_columns, state="disabled")
    rgb_button.grid(row=0, column=3, padx=5)

    saturation_button = tk.Button(button_frame, text="Afficher % Saturation", command=toggle_saturation_columns, state="disabled")
    saturation_button.grid(row=0, column=4, padx=5)

    code_nom_button = tk.Button(button_frame, text="Afficher Code+Nom", command=toggle_code_nom_columns, state="disabled")
    code_nom_button.grid(row=0, column=5, padx=5)

    create_export_buttons()

    quit_button = tk.Button(button_frame, text="Quitter", command=root.quit)
    quit_button.grid(row=0, column=8, padx=5)

    root.mainloop()

if __name__ == "__main__":
    main()
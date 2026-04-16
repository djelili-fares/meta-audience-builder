"""
Script : filter_high_value_clients.py

Objectif :
Filtrer un fichier CSV contenant des clients et garder uniquement ceux
dont la valeur (champ "value") est supérieure à 4000.

Entrée :
- meta_custom_audience_master.csv

Sortie :
- meta_custom_audience_high_value_4000plus.csv

Colonnes attendues :
phone, fn, ln, ct, st, country, value

Auteur : Toi (future machine à cash 😄)
"""

import csv

# ==============================
# CONFIGURATION
# ==============================

INPUT_FILE = "meta_custom_audience_master.csv"
OUTPUT_FILE = "meta_custom_audience_high_value_5000plus.csv"
VALUE_THRESHOLD = 5000  # seuil de filtrage

# ==============================
# TRAITEMENT
# ==============================

count_total = 0
count_kept = 0

with open(INPUT_FILE, mode='r', encoding='utf-8') as infile, \
     open(OUTPUT_FILE, mode='w', newline='', encoding='utf-8') as outfile:
    
    reader = csv.DictReader(infile)
    writer = csv.DictWriter(outfile, fieldnames=reader.fieldnames)
    
    # Écrire l'en-tête dans le fichier de sortie
    writer.writeheader()
    
    # Parcours des lignes
    for row in reader:
        count_total += 1
        
        try:
            # Conversion de la valeur en float
            value = float(row['value'])
            
            # Condition de filtrage
            if value > VALUE_THRESHOLD:
                writer.writerow(row)
                count_kept += 1
                
        except Exception as e:
            # Ignore les lignes problématiques
            print(f"Ligne ignorée (erreur) : {row} | Erreur : {e}")
            continue

# ==============================
# RÉSULTATS
# ==============================

print("===================================")
print("Filtrage terminé")
print("===================================")
print(f"Total clients analysés : {count_total}")
print(f"Clients conservés (> {VALUE_THRESHOLD}) : {count_kept}")
print(f"Pourcentage conservé : {round((count_kept / count_total) * 100, 2)} %")
print(f"Fichier généré : {OUTPUT_FILE}")
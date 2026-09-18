"""
Analyseur de consommation énergie — électricité & gaz
-------------------------------------------------------
Lit un fichier CSV de relevés de consommation, nettoie les lignes
invalides, calcule des statistiques, génère un graphique de l'évolution
mensuelle, puis exporte un rapport complet en fichier Excel.

Auteur : Solofonirina
Utilise uniquement : csv, json (stdlib) + matplotlib, openpyxl
"""

import csv
import json
import os
from datetime import datetime

import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill

# ──────────────────────────────────────────────────────────
# CONFIGURATION
# ──────────────────────────────────────────────────────────
FICHIER_ENTREE = "consommation_exemple.csv"
FICHIER_GRAPHIQUE = "evolution_consommation.png"
FICHIER_RAPPORT_EXCEL = "rapport_consommation.xlsx"
FICHIER_LOG_ERREURS = "lignes_ignorees.json"


# ──────────────────────────────────────────────────────────
# ÉTAPE 1 — LECTURE ET NETTOYAGE DU CSV
# ──────────────────────────────────────────────────────────
def lire_et_nettoyer(chemin_fichier):
    """
    Lit le CSV et ne garde que les lignes valides.
    Une ligne est valide si :
      - la consommation (kWh) est un nombre positif
      - le montant (€) est un nombre positif
    Les lignes invalides sont mises de côté (pas supprimées en silence).
    Retourne (lignes_valides, lignes_ignorees).
    """
    if not os.path.exists(chemin_fichier):
        print(f"❌ Erreur : le fichier '{chemin_fichier}' est introuvable.")
        print("   Vérifie qu'il se trouve dans le même dossier que ce script.")
        return [], []

    lignes_valides = []
    lignes_ignorees = []

    with open(chemin_fichier, mode="r", encoding="utf-8") as fichier:
        lecteur = csv.DictReader(fichier)

        colonnes_attendues = {"date", "type_energie", "consommation_kwh", "montant_eur"}
        if not colonnes_attendues.issubset(set(lecteur.fieldnames or [])):
            print(f"❌ Erreur : colonnes manquantes. Attendu : {colonnes_attendues}")
            print(f"   Trouvé : {lecteur.fieldnames}")
            return [], []

        for i, ligne in enumerate(lecteur, start=2):  # start=2 : ligne 1 = en-tête
            conso_brute = ligne["consommation_kwh"].strip()
            montant_brut = ligne["montant_eur"].strip()

            # Vérification : la consommation doit être un nombre positif
            try:
                conso = float(conso_brute)
                if conso <= 0:
                    raise ValueError("consommation négative ou nulle")
            except ValueError:
                lignes_ignorees.append({"ligne": i, "raison": "consommation invalide", "contenu": ligne})
                continue

            # Vérification : le montant doit être un nombre positif
            try:
                montant = float(montant_brut)
                if montant <= 0:
                    raise ValueError("montant négatif ou nul")
            except ValueError:
                lignes_ignorees.append({"ligne": i, "raison": "montant invalide", "contenu": ligne})
                continue

            lignes_valides.append({
                "date": ligne["date"].strip(),
                "type_energie": ligne["type_energie"].strip().lower(),
                "consommation_kwh": conso,
                "montant_eur": montant,
            })

    return lignes_valides, lignes_ignorees


# ──────────────────────────────────────────────────────────
# ÉTAPE 2 — CALCUL DES STATISTIQUES
# ──────────────────────────────────────────────────────────
def calculer_statistiques(lignes):
    """Calcule les statistiques globales et par type d'énergie."""
    stats = {
        "electricite": {"total_kwh": 0, "total_eur": 0, "nb_releves": 0},
        "gaz": {"total_kwh": 0, "total_eur": 0, "nb_releves": 0},
    }

    for ligne in lignes:
        type_e = ligne["type_energie"]
        if type_e not in stats:
            continue
        stats[type_e]["total_kwh"] += ligne["consommation_kwh"]
        stats[type_e]["total_eur"] += ligne["montant_eur"]
        stats[type_e]["nb_releves"] += 1

    for type_e in stats:
        nb = stats[type_e]["nb_releves"]
        stats[type_e]["moyenne_kwh"] = stats[type_e]["total_kwh"] / nb if nb else 0
        stats[type_e]["moyenne_eur"] = stats[type_e]["total_eur"] / nb if nb else 0

    # Trouve le mois le plus cher (tous types confondus)
    montants_par_mois = {}
    for ligne in lignes:
        mois = ligne["date"][:7]  # "2025-03-05" -> "2025-03"
        montants_par_mois[mois] = montants_par_mois.get(mois, 0) + ligne["montant_eur"]

    mois_le_plus_cher = max(montants_par_mois, key=montants_par_mois.get) if montants_par_mois else None

    return stats, montants_par_mois, mois_le_plus_cher


# ──────────────────────────────────────────────────────────
# ÉTAPE 3 — GÉNÉRATION DU GRAPHIQUE
# ──────────────────────────────────────────────────────────
def generer_graphique(montants_par_mois, chemin_sortie):
    """Génère un graphique en barres de l'évolution mensuelle des dépenses."""
    mois_tries = sorted(montants_par_mois.keys())
    valeurs = [montants_par_mois[m] for m in mois_tries]

    plt.figure(figsize=(9, 4.5))
    barres = plt.bar(mois_tries, valeurs, color="#3ddc84")
    plt.title("Évolution mensuelle des dépenses énergie (électricité + gaz)")
    plt.xlabel("Mois")
    plt.ylabel("Montant total (€)")
    plt.xticks(rotation=45, ha="right")

    # Affiche la valeur au-dessus de chaque barre
    for barre in barres:
        hauteur = barre.get_height()
        plt.text(barre.get_x() + barre.get_width() / 2, hauteur + 1,
                  f"{hauteur:.0f}€", ha="center", fontsize=8)

    plt.tight_layout()
    plt.savefig(chemin_sortie, dpi=150)
    plt.close()
    print(f"✅ Graphique généré : {chemin_sortie}")


# ──────────────────────────────────────────────────────────
# ÉTAPE 4 — EXPORT DU RAPPORT EXCEL
# ──────────────────────────────────────────────────────────
def exporter_rapport_excel(stats, mois_le_plus_cher, nb_valides, nb_ignorees, chemin_graphique, chemin_sortie):
    """Construit un fichier Excel avec les statistiques et le graphique intégré."""
    classeur = Workbook()
    feuille = classeur.active
    feuille.title = "Rapport"

    style_titre = Font(size=14, bold=True, color="FFFFFF")
    remplissage_titre = PatternFill(start_color="3DDC84", end_color="3DDC84", fill_type="solid")
    style_entete = Font(bold=True)

    feuille["A1"] = "Rapport de consommation énergie"
    feuille["A1"].font = style_titre
    feuille["A1"].fill = remplissage_titre
    feuille.merge_cells("A1:D1")

    feuille["A3"] = "Résumé"
    feuille["A3"].font = style_entete
    feuille["A4"] = "Lignes valides analysées"
    feuille["B4"] = nb_valides
    feuille["A5"] = "Lignes ignorées (données invalides)"
    feuille["B5"] = nb_ignorees
    feuille["A6"] = "Mois le plus cher"
    feuille["B6"] = mois_le_plus_cher or "N/A"

    feuille["A8"] = "Type"
    feuille["B8"] = "Relevés"
    feuille["C8"] = "Total kWh"
    feuille["D8"] = "Total €"
    feuille["E8"] = "Moyenne kWh"
    feuille["F8"] = "Moyenne €"
    for col in "ABCDEF":
        feuille[f"{col}8"].font = style_entete

    ligne_excel = 9
    for type_e, valeurs in stats.items():
        feuille[f"A{ligne_excel}"] = type_e.capitalize()
        feuille[f"B{ligne_excel}"] = valeurs["nb_releves"]
        feuille[f"C{ligne_excel}"] = round(valeurs["total_kwh"], 1)
        feuille[f"D{ligne_excel}"] = round(valeurs["total_eur"], 2)
        feuille[f"E{ligne_excel}"] = round(valeurs["moyenne_kwh"], 1)
        feuille[f"F{ligne_excel}"] = round(valeurs["moyenne_eur"], 2)
        ligne_excel += 1

    if os.path.exists(chemin_graphique):
        image = XLImage(chemin_graphique)
        image.width = 560
        image.height = 280
        feuille.add_image(image, "A13")

    for col, largeur in zip("ABCDEF", [22, 12, 12, 12, 14, 12]):
        feuille.column_dimensions[col].width = largeur

    classeur.save(chemin_sortie)
    print(f"✅ Rapport Excel généré : {chemin_sortie}")


# ──────────────────────────────────────────────────────────
# ÉTAPE 5 — SAUVEGARDE DES LIGNES IGNORÉES (traçabilité)
# ──────────────────────────────────────────────────────────
def sauvegarder_lignes_ignorees(lignes_ignorees, chemin_sortie):
    """Sauvegarde en JSON la liste des lignes rejetées, avec la raison exacte."""
    with open(chemin_sortie, "w", encoding="utf-8") as fichier:
        json.dump(lignes_ignorees, fichier, indent=2, ensure_ascii=False)
    if lignes_ignorees:
        print(f"⚠️  {len(lignes_ignorees)} ligne(s) ignorée(s) — détail dans {chemin_sortie}")


# ──────────────────────────────────────────────────────────
# PROGRAMME PRINCIPAL
# ──────────────────────────────────────────────────────────
def main():
    print("=== Analyseur de consommation énergie ===\n")

    lignes_valides, lignes_ignorees = lire_et_nettoyer(FICHIER_ENTREE)
    if not lignes_valides:
        print("Aucune ligne valide trouvée — arrêt du programme.")
        return

    stats, montants_par_mois, mois_le_plus_cher = calculer_statistiques(lignes_valides)

    generer_graphique(montants_par_mois, FICHIER_GRAPHIQUE)

    exporter_rapport_excel(
        stats, mois_le_plus_cher,
        len(lignes_valides), len(lignes_ignorees),
        FICHIER_GRAPHIQUE, FICHIER_RAPPORT_EXCEL
    )

    sauvegarder_lignes_ignorees(lignes_ignorees, FICHIER_LOG_ERREURS)

    print("\n--- Résumé ---")
    for type_e, valeurs in stats.items():
        print(f"{type_e.capitalize():12} | {valeurs['nb_releves']} relevés | "
              f"{valeurs['total_kwh']:.0f} kWh | {valeurs['total_eur']:.2f} €")
    print(f"Mois le plus cher : {mois_le_plus_cher}")
    print("\n✅ Analyse terminée.")


if __name__ == "__main__":
    main()

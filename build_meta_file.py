"""
Construction et mise à jour d'un fichier client principal conforme Meta Ads
pour Custom Audience / Value-Based Audience.

Utilisation :
    python build_meta_file.py fichier1.xlsx fichier2.xlsx fichier3.xlsx fichier4.xlsx

Exemple :
    python build_meta_file.py Anderson_03-07-2026.xlsx guepex_03-07-2026.xlsx Univer_03-07-2026.xlsx Zimou_03-07-2026.xlsx --master rachelle_clients_03-07-2026.csv --reset-master

Fonctionnalités :
- Accepte de 1 à 4 fichiers sources en arguments
- Lit des fichiers Excel / CSV
- Détecte automatiquement la structure des colonnes
- Accepte "Prix" ou "Montant" comme colonne de valeur
- Normalise les noms, wilayas, communes et valeurs
- Formate les numéros algériens au format international attendu par Meta
- Déduplique les clients sur la base du téléphone
- Additionne les valeurs d'achat par client unique
- Sauvegarde un fichier principal compatible Meta
- Génère des statistiques par wilaya
- Sauvegarde un fichier log dédié des wilayas

Auteur : ChatGPT
Contexte : Shopify / COD / Algérie / Meta Ads
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Optional, Tuple

import pandas as pd


# ============================================================
# COLORAMA OPTIONNEL
# ============================================================

try:
    from colorama import Fore, Style, init

    init(autoreset=True)
except ImportError:

    class DummyColor:
        RED = ""
        GREEN = ""
        YELLOW = ""
        CYAN = ""
        MAGENTA = ""
        WHITE = ""
        RESET_ALL = ""

    Fore = Style = DummyColor()


# ============================================================
# CONFIGURATION
# ============================================================

META_COLUMNS = ["phone", "fn", "ln", "ct", "st", "country", "value"]
DEFAULT_COUNTRY = "DZ"

MASTER_FILE = "meta_custom_audience_master.csv"
WILAYA_LOG_FILE = "wilaya_stats_log.txt"


# ============================================================
# OUTILS DE NORMALISATION
# ============================================================

def clean_text(value: object) -> str:
    """Nettoie une valeur texte."""
    if pd.isna(value):
        return ""

    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_phone_dz(phone: object) -> str:
    """Normalise un numéro algérien au format international Meta."""
    raw = clean_text(phone)

    if not raw:
        return ""

    digits = re.sub(r"\D", "", raw)

    if not digits:
        return ""

    if digits.startswith("00213"):
        digits = digits[2:]

    if digits.startswith("213"):
        return digits

    if digits.startswith("0") and len(digits) == 10:
        return "213" + digits[1:]

    if len(digits) == 9:
        return "213" + digits

    return digits


def normalize_value(value: object) -> Optional[float]:
    """Normalise la valeur de commande."""
    if pd.isna(value):
        return None

    text = str(value).strip().replace(",", ".")
    text = re.sub(r"[^\d.]", "", text)

    if not text:
        return None

    try:
        return float(text)
    except ValueError:
        return None


def split_client_name(full_name: object) -> Tuple[str, str]:
    """Sépare un champ Client en prénom et nom."""
    text = clean_text(full_name)

    if not text:
        return "", ""

    parts = text.split(" ")
    first_name = parts[0]
    last_name = " ".join(parts[1:]) if len(parts) > 1 else ""

    return first_name, last_name


# ============================================================
# CHARGEMENT DES FICHIERS
# ============================================================

def read_table(file_path: str | Path) -> pd.DataFrame:
    """Lit un fichier Excel ou CSV en DataFrame."""
    file_path = Path(file_path)

    if not file_path.exists():
        raise FileNotFoundError(f"Fichier introuvable : {file_path}")

    suffix = file_path.suffix.lower()

    if suffix in [".xlsx", ".xls"]:
        return pd.read_excel(file_path)

    if suffix == ".csv":
        try:
            return pd.read_csv(file_path)
        except UnicodeDecodeError:
            return pd.read_csv(file_path, encoding="latin1")

    raise ValueError(f"Format non supporté : {suffix}")


def find_first_existing_column(df: pd.DataFrame, candidates: list[str]) -> str:
    """
    Retourne la première colonne trouvée dans le DataFrame.

    Exemple :
    candidates = ["Prix", "Montant"]
    """
    for column in candidates:
        if column in df.columns:
            return column

    raise ValueError(
        f"Aucune colonne trouvée parmi : {candidates}. "
        f"Colonnes disponibles : {list(df.columns)}"
    )


# ============================================================
# TRANSFORMATIONS PAR TYPE DE FICHIER SOURCE
# ============================================================

def transform_source_type_1(df: pd.DataFrame) -> pd.DataFrame:
    """
    Type 1 :
    Prénom | Nom | Téléphone | Wilaya | Commune | Prix/Montant
    """
    required_columns = ["Prénom", "Nom", "Téléphone", "Wilaya", "Commune"]
    missing = [col for col in required_columns if col not in df.columns]

    if missing:
        raise ValueError(f"Colonnes manquantes pour le fichier type 1 : {missing}")

    value_column = find_first_existing_column(df, ["Prix", "Montant"])

    out = pd.DataFrame()
    out["phone"] = df["Téléphone"].apply(normalize_phone_dz)
    out["fn"] = df["Prénom"].apply(clean_text)
    out["ln"] = df["Nom"].apply(clean_text)
    out["ct"] = df["Commune"].apply(clean_text)
    out["st"] = df["Wilaya"].apply(clean_text)
    out["country"] = DEFAULT_COUNTRY
    out["value"] = df[value_column].apply(normalize_value)

    return out


def transform_source_type_2(df: pd.DataFrame) -> pd.DataFrame:
    """
    Type 2 :
    Client | Téléphone | Wilaya | Commune | Prix/Montant
    """
    required_columns = ["Client", "Téléphone", "Wilaya", "Commune"]
    missing = [col for col in required_columns if col not in df.columns]

    if missing:
        raise ValueError(f"Colonnes manquantes pour le fichier type 2 : {missing}")

    value_column = find_first_existing_column(df, ["Prix", "Montant"])
    names = df["Client"].apply(split_client_name)

    out = pd.DataFrame()
    out["phone"] = df["Téléphone"].apply(normalize_phone_dz)
    out["fn"] = names.apply(lambda x: x[0])
    out["ln"] = names.apply(lambda x: x[1])
    out["ct"] = df["Commune"].apply(clean_text)
    out["st"] = df["Wilaya"].apply(clean_text)
    out["country"] = DEFAULT_COUNTRY
    out["value"] = df[value_column].apply(normalize_value)

    return out


def transform_source_type_3(df: pd.DataFrame) -> pd.DataFrame:
    """
    Type 3 - Guepex :
    prénom | nom | téléphone | commune | wilaya destination | prix

    Ce format est utilisé par les exports Guepex.
    Les noms de colonnes sont en minuscules.
    """
    required_columns = [
        "prénom",
        "nom",
        "téléphone",
        "commune",
        "wilaya destination",
        "prix",
    ]

    missing = [col for col in required_columns if col not in df.columns]

    if missing:
        raise ValueError(f"Colonnes manquantes pour le fichier type 3 : {missing}")

    out = pd.DataFrame()
    out["phone"] = df["téléphone"].apply(normalize_phone_dz)
    out["fn"] = df["prénom"].apply(clean_text)
    out["ln"] = df["nom"].apply(clean_text)
    out["ct"] = df["commune"].apply(clean_text)
    out["st"] = df["wilaya destination"].apply(clean_text)
    out["country"] = DEFAULT_COUNTRY
    out["value"] = df["prix"].apply(normalize_value)

    return out


# ============================================================
# QUALITE / NETTOYAGE
# ============================================================

def print_duplicate_resolution_report(duplicates_report: list[dict]) -> None:
    """Affiche un rapport détaillé des doublons téléphone."""
    if not duplicates_report:
        return

    print(f"\n{Fore.MAGENTA}{'=' * 90}{Style.RESET_ALL}")
    print(f"{Fore.MAGENTA}DOUBLONS TELEPHONE TRAITES - LOGIQUE VALUE-BASED{Style.RESET_ALL}")
    print(f"{Fore.MAGENTA}{'=' * 90}{Style.RESET_ALL}")

    total_groups = len(duplicates_report)
    total_removed = sum(len(item["removed_rows"]) for item in duplicates_report)
    total_added_value = sum(item["total_value"] for item in duplicates_report)

    print(f"{Fore.WHITE}Nombre de téléphones en doublon : {total_groups}{Style.RESET_ALL}")
    print(f"{Fore.WHITE}Nombre de variantes supprimées : {total_removed}{Style.RESET_ALL}")
    print(f"{Fore.WHITE}Valeur totale cumulée concernée : {total_added_value:.2f}{Style.RESET_ALL}")
    print(f"{Fore.MAGENTA}{'-' * 90}{Style.RESET_ALL}")

    for idx, item in enumerate(duplicates_report, start=1):
        phone = item["phone"]
        kept_row = item["kept_row"]
        removed_rows = item["removed_rows"]
        total_value = item["total_value"]
        variants_count = item["variants_count"]

        print(f"\n{Fore.CYAN}[DOUBLON #{idx}]{Style.RESET_ALL} Téléphone : {phone}")
        print(f"{Fore.WHITE}Nombre de variantes : {variants_count}{Style.RESET_ALL}")
        print(f"{Fore.WHITE}Valeur cumulée finale : {total_value:.2f}{Style.RESET_ALL}")

        print(f"\n{Fore.GREEN}LIGNE GARDEE :{Style.RESET_ALL}")
        print(
            f"  phone={kept_row['phone']} | "
            f"fn={kept_row['fn']} | "
            f"ln={kept_row['ln']} | "
            f"wilaya={kept_row['st']} | "
            f"commune={kept_row['ct']} | "
            f"ancienne_value={kept_row['original_value']:.2f} | "
            f"nouvelle_value={total_value:.2f}"
        )

        print(f"{Fore.RED}VARIANTES SUPPRIMEES :{Style.RESET_ALL}")

        for removed in removed_rows:
            print(
                f"  - phone={removed['phone']} | "
                f"fn={removed['fn']} | "
                f"ln={removed['ln']} | "
                f"wilaya={removed['st']} | "
                f"commune={removed['ct']} | "
                f"value={removed['original_value']:.2f}"
            )

        print(f"{Fore.MAGENTA}{'-' * 90}{Style.RESET_ALL}")

    print(
        f"{Fore.YELLOW}[INFO]{Style.RESET_ALL} "
        "Pour Meta Value-Based Audience, la valeur finale représente le total acheté par client unique."
    )
    print(f"{Fore.MAGENTA}{'=' * 90}{Style.RESET_ALL}\n")


def clean_meta_dataframe(
    df: pd.DataFrame,
    show_removed: bool = False,
    show_duplicates_report: bool = False,
) -> pd.DataFrame:
    """
    Nettoie le DataFrame final Meta.

    Règles :
    - supprime les lignes sans téléphone
    - supprime les lignes sans prénom et sans nom
    - déduplique par téléphone
    - garde la variante avec la valeur la plus élevée
    - additionne toutes les valeurs du même téléphone
    """
    df = df.copy()
    df = df[META_COLUMNS]

    for col in ["phone", "fn", "ln", "ct", "st", "country"]:
        df[col] = df[col].apply(clean_text)

    df["country"] = DEFAULT_COUNTRY
    df["value"] = pd.to_numeric(df["value"], errors="coerce").fillna(0)

    removed_rows = []

    no_phone = df[df["phone"] == ""].copy()
    if not no_phone.empty:
        no_phone["delete_reason"] = "Téléphone vide ou invalide"
        removed_rows.append(no_phone)

    df = df[df["phone"] != ""]

    no_name = df[(df["fn"] == "") & (df["ln"] == "")].copy()
    if not no_name.empty:
        no_name["delete_reason"] = "Prénom et nom vides"
        removed_rows.append(no_name)

    df = df[(df["fn"] != "") | (df["ln"] != "")]

    final_rows = []
    duplicates_report = []

    for phone, group in df.groupby("phone", sort=False):
        group = group.copy()

        if len(group) == 1:
            final_rows.append(group.iloc[0])
            continue

        kept_index = group["value"].idxmax()
        kept_row = group.loc[kept_index].copy()
        total_value = group["value"].sum()

        removed_group = group.drop(index=kept_index).copy()
        removed_group["delete_reason"] = "Doublon téléphone supprimé - valeur plus basse"
        removed_rows.append(removed_group)

        kept_report = kept_row.copy()
        kept_report["original_value"] = kept_row["value"]

        removed_report_rows = []

        for _, removed_row in removed_group.iterrows():
            removed_report = removed_row.copy()
            removed_report["original_value"] = removed_row["value"]
            removed_report_rows.append(removed_report)

        duplicates_report.append(
            {
                "phone": phone,
                "variants_count": len(group),
                "kept_row": kept_report,
                "removed_rows": removed_report_rows,
                "total_value": total_value,
            }
        )

        kept_row["value"] = total_value
        final_rows.append(kept_row)

    if final_rows:
        df = pd.DataFrame(final_rows)
    else:
        df = pd.DataFrame(columns=META_COLUMNS)

    df = df[META_COLUMNS].reset_index(drop=True)

    if show_removed:
        if removed_rows:
            removed_df = pd.concat(removed_rows, ignore_index=True)

            print(f"\n{Fore.RED}{'=' * 90}{Style.RESET_ALL}")
            print(f"{Fore.RED}CLIENTS / VARIANTES SUPPRIMES{Style.RESET_ALL}")
            print(f"{Fore.RED}{'=' * 90}{Style.RESET_ALL}")

            display_cols = ["phone", "fn", "ln", "ct", "st", "value", "delete_reason"]
            print(removed_df[display_cols].to_string(index=False))

            print(f"{Fore.RED}{'-' * 90}{Style.RESET_ALL}")
            print(f"{Fore.YELLOW}[INFO]{Style.RESET_ALL} Total lignes supprimées : {len(removed_df)}\n")
        else:
            print(f"{Fore.GREEN}[OK]{Style.RESET_ALL} Aucun client supprimé pendant le nettoyage.")

    if show_duplicates_report:
        print_duplicate_resolution_report(duplicates_report)

    return df


# ============================================================
# FICHIER PRINCIPAL
# ============================================================

def initialize_master_file(master_file: str | Path) -> None:
    """Crée le fichier principal s'il n'existe pas encore."""
    master_path = Path(master_file)

    if not master_path.exists():
        empty_df = pd.DataFrame(columns=META_COLUMNS)
        empty_df.to_csv(master_path, index=False, encoding="utf-8-sig")
        print(f"{Fore.GREEN}[OK]{Style.RESET_ALL} Fichier principal créé : {master_path}")
    else:
        print(f"{Fore.CYAN}[INFO]{Style.RESET_ALL} Fichier principal existe déjà : {master_path}")


def reset_master_file(master_file: str | Path) -> None:
    """Réinitialise le fichier master pour repartir de zéro."""
    master_path = Path(master_file)
    empty_df = pd.DataFrame(columns=META_COLUMNS)
    empty_df.to_csv(master_path, index=False, encoding="utf-8-sig")
    print(f"{Fore.YELLOW}[RESET]{Style.RESET_ALL} Fichier principal réinitialisé : {master_path}")


def load_master_file(master_file: str | Path) -> pd.DataFrame:
    """Charge le fichier principal CSV."""
    initialize_master_file(master_file)

    df = pd.read_csv(master_file, encoding="utf-8-sig")

    for col in META_COLUMNS:
        if col not in df.columns:
            df[col] = None

    return df[META_COLUMNS]


def save_master_file(df: pd.DataFrame, master_file: str | Path) -> None:
    """Sauvegarde le fichier principal."""
    df.to_csv(master_file, index=False, encoding="utf-8-sig")
    print(f"{Fore.GREEN}[OK]{Style.RESET_ALL} Fichier principal sauvegardé : {master_file}")


# ============================================================
# PIPELINE D'IMPORT
# ============================================================

def detect_source_type(file_path: str | Path) -> int:
    """
    Détecte automatiquement le type de fichier source selon les colonnes.

    Type 1 :
    Prénom | Nom | Téléphone | Wilaya | Commune | Prix/Montant

    Type 2 :
    Client | Téléphone | Wilaya | Commune | Prix/Montant

    Type 3 - Guepex :
    prénom | nom | téléphone | commune | wilaya destination | prix
    """
    df = read_table(file_path)
    columns = set(df.columns)

    type_1_base = {"Prénom", "Nom", "Téléphone", "Wilaya", "Commune"}
    type_2_base = {"Client", "Téléphone", "Wilaya", "Commune"}
    type_3_base = {
        "prénom",
        "nom",
        "téléphone",
        "commune",
        "wilaya destination",
        "prix",
    }

    value_columns = {"Prix", "Montant"}

    has_value_column = bool(columns.intersection(value_columns))

    if type_1_base.issubset(columns) and has_value_column:
        return 1

    if type_2_base.issubset(columns) and has_value_column:
        return 2

    if type_3_base.issubset(columns):
        return 3

    raise ValueError(
        f"Impossible de détecter le type du fichier {file_path}.\n"
        f"Colonnes trouvées : {list(df.columns)}\n"
        "Structures acceptées :\n"
        "- Type 1 : Prénom, Nom, Téléphone, Wilaya, Commune, Prix ou Montant\n"
        "- Type 2 : Client, Téléphone, Wilaya, Commune, Prix ou Montant\n"
        "- Type 3 : prénom, nom, téléphone, commune, wilaya destination, prix"
    )


def import_source_file(
    source_file: str | Path,
    source_type: int,
) -> pd.DataFrame:
    """
    Importe un fichier source et le transforme au format Meta.
    """
    raw_df = read_table(source_file)

    if source_type == 1:
        transformed = transform_source_type_1(raw_df)
    elif source_type == 2:
        transformed = transform_source_type_2(raw_df)
    elif source_type == 3:
        transformed = transform_source_type_3(raw_df)
    else:
        raise ValueError("source_type doit être 1, 2 ou 3")

    cleaned = clean_meta_dataframe(
        transformed,
        show_removed=True,
        show_duplicates_report=True,
    )

    return cleaned


def merge_into_master(
    master_file: str | Path,
    imported_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Fusionne les nouvelles données avec le fichier principal.

    Si un téléphone existe déjà :
    - garde la variante avec la valeur la plus élevée
    - additionne les valeurs
    - conserve une seule ligne par téléphone
    """
    master_df = load_master_file(master_file)
    combined = pd.concat([master_df, imported_df], ignore_index=True)

    combined = clean_meta_dataframe(
        combined,
        show_removed=False,
        show_duplicates_report=True,
    )

    save_master_file(combined, master_file)

    return combined


# ============================================================
# STATS WILAYAS
# ============================================================

def build_wilaya_stats(df: pd.DataFrame) -> pd.DataFrame:
    """Construit les stats par wilaya."""
    stats_df = df.copy()

    stats_df["st"] = stats_df["st"].apply(clean_text)
    stats_df = stats_df[stats_df["st"] != ""]

    total = len(stats_df)

    if total == 0:
        return pd.DataFrame(columns=["wilaya", "count", "percentage"])

    summary = (
        stats_df.groupby("st", dropna=False)
        .size()
        .reset_index(name="count")
        .rename(columns={"st": "wilaya"})
        .sort_values(by="count", ascending=False)
        .reset_index(drop=True)
    )

    summary["percentage"] = (summary["count"] / total * 100).round(2)

    return summary


def get_color_for_rank(index: int) -> str:
    """Couleur console selon le classement."""
    if index < 3:
        return Fore.GREEN

    if index < 10:
        return Fore.YELLOW

    return Fore.CYAN


def format_wilaya_stats_text(stats_df: pd.DataFrame, total_clients: int) -> str:
    """Formate le texte du log wilayas."""
    lines = []
    lines.append("=" * 72)
    lines.append("REPARTITION DES CLIENTS LIVRES PAR WILAYA")
    lines.append("=" * 72)
    lines.append(f"Total clients uniques retenus dans le master : {total_clients}")
    lines.append("")

    if stats_df.empty:
        lines.append("Aucune wilaya exploitable trouvée.")
        return "\n".join(lines)

    lines.append(f"{'RANG':<6}{'WILAYA':<25}{'NB CLIENTS':<15}{'% TOTAL':<10}")
    lines.append("-" * 72)

    for idx, row in stats_df.iterrows():
        rank = idx + 1
        wilaya = str(row["wilaya"])
        count = int(row["count"])
        pct = float(row["percentage"])

        lines.append(f"{rank:<6}{wilaya:<25}{count:<15}{pct:<10.2f}")

    lines.append("-" * 72)
    lines.append("Lecture business :")
    lines.append("- Les wilayas en haut du classement sont les plus représentées dans ta base clients livrés.")
    lines.append("- Elles peuvent servir pour prioriser le ciblage géographique, la logistique et l'analyse de rentabilité.")
    lines.append("=" * 72)

    return "\n".join(lines)


def print_wilaya_stats_colored(stats_df: pd.DataFrame, total_clients: int) -> None:
    """Affiche les stats wilayas dans le terminal avec couleurs."""
    print(f"\n{Fore.MAGENTA}{'=' * 72}{Style.RESET_ALL}")
    print(f"{Fore.MAGENTA}REPARTITION DES CLIENTS LIVRES PAR WILAYA{Style.RESET_ALL}")
    print(f"{Fore.MAGENTA}{'=' * 72}{Style.RESET_ALL}")
    print(f"{Fore.WHITE}Total clients uniques retenus dans le master : {total_clients}{Style.RESET_ALL}\n")

    if stats_df.empty:
        print(f"{Fore.RED}[ALERTE]{Style.RESET_ALL} Aucune wilaya exploitable trouvée.")
        return

    print(f"{Fore.WHITE}{'RANG':<6}{'WILAYA':<25}{'NB CLIENTS':<15}{'% TOTAL':<10}{Style.RESET_ALL}")
    print(f"{Fore.WHITE}{'-' * 72}{Style.RESET_ALL}")

    for idx, row in stats_df.iterrows():
        color = get_color_for_rank(idx)
        rank = idx + 1
        wilaya = str(row["wilaya"])
        count = int(row["count"])
        pct = float(row["percentage"])

        print(f"{color}{rank:<6}{wilaya:<25}{count:<15}{pct:<10.2f}{Style.RESET_ALL}")

    print(f"{Fore.WHITE}{'-' * 72}{Style.RESET_ALL}")
    print(f"{Fore.CYAN}[INFO]{Style.RESET_ALL} Les 3 premières wilayas sont tes zones les plus représentées.")
    print(f"{Fore.YELLOW}[INFO]{Style.RESET_ALL} A croiser ensuite avec CPA confirmé / livré et taux de refus.\n")


def save_wilaya_log(stats_df: pd.DataFrame, total_clients: int, log_file: str | Path) -> None:
    """Sauvegarde le log wilayas dans un fichier texte."""
    report_text = format_wilaya_stats_text(stats_df, total_clients)
    Path(log_file).write_text(report_text, encoding="utf-8")

    print(f"{Fore.GREEN}[OK]{Style.RESET_ALL} Log wilayas sauvegardé : {log_file}")


# ============================================================
# ARGUMENTS CLI
# ============================================================

def parse_args() -> argparse.Namespace:
    """Parse les arguments passés en ligne de commande."""
    parser = argparse.ArgumentParser(
        description="Construit un fichier Meta Custom Audience à partir de 1 à 4 fichiers sources."
    )

    parser.add_argument(
        "files",
        nargs="+",
        help="Fichiers sources Excel/CSV à importer. Maximum 4 fichiers.",
    )

    parser.add_argument(
        "--master",
        default=MASTER_FILE,
        help=f"Nom du fichier master de sortie. Défaut : {MASTER_FILE}",
    )

    parser.add_argument(
        "--wilaya-log",
        default=WILAYA_LOG_FILE,
        help=f"Nom du fichier log wilayas. Défaut : {WILAYA_LOG_FILE}",
    )

    parser.add_argument(
        "--reset-master",
        action="store_true",
        help="Réinitialise le fichier master avant l'import.",
    )

    return parser.parse_args()


# ============================================================
# MAIN
# ============================================================

def main() -> None:
    """Point d'entrée principal du script."""
    args = parse_args()

    if len(args.files) > 4:
        raise ValueError("Tu peux fournir au maximum 4 fichiers sources.")

    master_file = args.master
    wilaya_log_file = args.wilaya_log

    if args.reset_master:
        reset_master_file(master_file)
    else:
        initialize_master_file(master_file)

    merged_df = None

    for source_file in args.files:
        source_path = Path(source_file)

        if not source_path.exists():
            print(f"{Fore.YELLOW}[WARN]{Style.RESET_ALL} Fichier introuvable : {source_file}")
            continue

        source_type = detect_source_type(source_path)

        print(
            f"{Fore.CYAN}[INFO]{Style.RESET_ALL} Import du fichier : "
            f"{source_file} | type détecté : {source_type}"
        )

        imported_df = import_source_file(source_path, source_type=source_type)
        merged_df = merge_into_master(master_file, imported_df)

        print(
            f"{Fore.GREEN}[OK]{Style.RESET_ALL} Après import {source_file} : "
            f"{len(merged_df)} clients uniques"
        )

    if merged_df is not None:
        wilaya_stats = build_wilaya_stats(merged_df)

        print_wilaya_stats_colored(
            wilaya_stats,
            total_clients=len(merged_df),
        )

        save_wilaya_log(
            wilaya_stats,
            total_clients=len(merged_df),
            log_file=wilaya_log_file,
        )

    print(f"{Fore.GREEN}[FIN]{Style.RESET_ALL} Pipeline terminé.")


if __name__ == "__main__":
    main()
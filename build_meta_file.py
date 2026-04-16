"""
Construction et mise à jour d'un fichier client principal conforme Meta Ads
pour Custom Audience / Value-Based Audience.

Fonctionnalités :
- Crée un fichier principal Meta s'il n'existe pas
- Lit plusieurs fichiers sources Excel / CSV
- Gère plusieurs structures de colonnes
- Normalise les noms, wilayas, communes, valeurs
- Formate les numéros algériens au format international attendu par Meta
- Déduplique les clients sur la base du téléphone
- Sauvegarde un fichier principal enrichi au fur et à mesure
- Génère des statistiques par wilaya (nombre + pourcentage)
- Sauvegarde un fichier log dédié des wilayas

Auteur : ChatGPT
Contexte : Shopify / COD / Algérie / Meta Ads
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional, Tuple

import pandas as pd

# ============================================================
# COLORAMA (optionnel)
# ============================================================

try:
    from colorama import Fore, Style, init
    init(autoreset=True)
    COLOR_ENABLED = True
except ImportError:
    COLOR_ENABLED = False

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
    """
    Nettoie une valeur texte.
    """
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_phone_dz(phone: object) -> str:
    """
    Normalise un numéro algérien au format international Meta.
    """
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
    """
    Normalise la valeur de commande.
    """
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
    """
    Sépare un champ 'Client' en prénom et nom.
    """
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
    """
    Lit un fichier Excel ou CSV en DataFrame.
    """
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


# ============================================================
# TRANSFORMATIONS PAR TYPE DE FICHIER SOURCE
# ============================================================

def transform_source_type_1(df: pd.DataFrame) -> pd.DataFrame:
    """
    Type 1 :
    Prénom | Nom | Téléphone | Wilaya | Commune | Prix
    """
    required_columns = ["Prénom", "Nom", "Téléphone", "Wilaya", "Commune", "Prix"]
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(f"Colonnes manquantes pour le fichier type 1 : {missing}")

    out = pd.DataFrame()
    out["phone"] = df["Téléphone"].apply(normalize_phone_dz)
    out["fn"] = df["Prénom"].apply(clean_text)
    out["ln"] = df["Nom"].apply(clean_text)
    out["ct"] = df["Commune"].apply(clean_text)
    out["st"] = df["Wilaya"].apply(clean_text)
    out["country"] = DEFAULT_COUNTRY
    out["value"] = df["Prix"].apply(normalize_value)

    return out


def transform_source_type_2(df: pd.DataFrame) -> pd.DataFrame:
    """
    Type 2 :
    Client | Téléphone | Wilaya | Commune | Prix
    """
    required_columns = ["Client", "Téléphone", "Wilaya", "Commune", "Prix"]
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(f"Colonnes manquantes pour le fichier type 2 : {missing}")

    names = df["Client"].apply(split_client_name)

    out = pd.DataFrame()
    out["phone"] = df["Téléphone"].apply(normalize_phone_dz)
    out["fn"] = names.apply(lambda x: x[0])
    out["ln"] = names.apply(lambda x: x[1])
    out["ct"] = df["Commune"].apply(clean_text)
    out["st"] = df["Wilaya"].apply(clean_text)
    out["country"] = DEFAULT_COUNTRY
    out["value"] = df["Prix"].apply(normalize_value)

    return out


# ============================================================
# QUALITE / NETTOYAGE
# ============================================================

def clean_meta_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Nettoie le DataFrame final Meta.
    """
    df = df.copy()
    df = df[META_COLUMNS]

    for col in ["phone", "fn", "ln", "ct", "st", "country"]:
        df[col] = df[col].apply(clean_text)

    df = df[df["phone"] != ""]
    df = df[(df["fn"] != "") | (df["ln"] != "")]
    df["country"] = DEFAULT_COUNTRY

    df = df.drop_duplicates(subset=["phone"], keep="last")
    df = df.reset_index(drop=True)

    return df


# ============================================================
# FICHIER PRINCIPAL
# ============================================================

def initialize_master_file(master_file: str | Path) -> None:
    """
    Crée le fichier principal s'il n'existe pas encore.
    """
    master_path = Path(master_file)
    if not master_path.exists():
        empty_df = pd.DataFrame(columns=META_COLUMNS)
        empty_df.to_csv(master_path, index=False, encoding="utf-8-sig")
        print(f"{Fore.GREEN}[OK]{Style.RESET_ALL} Fichier principal créé : {master_path}")
    else:
        print(f"{Fore.CYAN}[INFO]{Style.RESET_ALL} Fichier principal existe déjà : {master_path}")


def load_master_file(master_file: str | Path) -> pd.DataFrame:
    """
    Charge le fichier principal CSV.
    """
    initialize_master_file(master_file)
    df = pd.read_csv(master_file, encoding="utf-8-sig")
    for col in META_COLUMNS:
        if col not in df.columns:
            df[col] = None
    return df[META_COLUMNS]


def save_master_file(df: pd.DataFrame, master_file: str | Path) -> None:
    """
    Sauvegarde le fichier principal.
    """
    df.to_csv(master_file, index=False, encoding="utf-8-sig")
    print(f"{Fore.GREEN}[OK]{Style.RESET_ALL} Fichier principal sauvegardé : {master_file}")


# ============================================================
# PIPELINE D'IMPORT
# ============================================================

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
    else:
        raise ValueError("source_type doit être 1 ou 2")

    cleaned = clean_meta_dataframe(transformed)
    return cleaned


def merge_into_master(
    master_file: str | Path,
    imported_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Fusionne les nouvelles données avec le fichier principal.
    """
    master_df = load_master_file(master_file)

    combined = pd.concat([master_df, imported_df], ignore_index=True)
    combined = clean_meta_dataframe(combined)

    save_master_file(combined, master_file)
    return combined


# ============================================================
# STATS WILAYAS
# ============================================================

def build_wilaya_stats(df: pd.DataFrame) -> pd.DataFrame:
    """
    Construit les stats par wilaya :
    - nombre de clients
    - pourcentage du total
    """
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
    """
    Couleur console selon le classement.
    """
    if index < 3:
        return Fore.GREEN
    if index < 10:
        return Fore.YELLOW
    return Fore.CYAN


def format_wilaya_stats_text(stats_df: pd.DataFrame, total_clients: int) -> str:
    """
    Formate le texte du log wilayas.
    """
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
    """
    Affiche les stats wilayas dans le terminal avec couleurs.
    """
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
    """
    Sauvegarde le log wilayas dans un fichier texte.
    """
    report_text = format_wilaya_stats_text(stats_df, total_clients)
    Path(log_file).write_text(report_text, encoding="utf-8")
    print(f"{Fore.GREEN}[OK]{Style.RESET_ALL} Log wilayas sauvegardé : {log_file}")


# ============================================================
# EXEMPLE D'UTILISATION
# ============================================================

def main() -> None:
    """
    Exemple complet d'utilisation.
    """
    master_file = MASTER_FILE
    wilaya_log_file = WILAYA_LOG_FILE

    # Ton fichier source livré
    source_file_1 = "clients_livres_rachelle_jusqua_16Avr2026.xlsx"
    # source_file_2 = "clients_source_type_2.xlsx"

    initialize_master_file(master_file)

    if Path(source_file_1).exists():
        df1 = import_source_file(source_file_1, source_type=1)
        merged_df = merge_into_master(master_file, df1)

        print(f"{Fore.CYAN}[INFO]{Style.RESET_ALL} Après import fichier 1 : {len(merged_df)} clients uniques")

        # Stats par wilaya sur le fichier final Meta
        wilaya_stats = build_wilaya_stats(merged_df)

        # Affichage console coloré
        print_wilaya_stats_colored(wilaya_stats, total_clients=len(merged_df))

        # Sauvegarde log texte
        save_wilaya_log(wilaya_stats, total_clients=len(merged_df), log_file=wilaya_log_file)

    # Exemple futur si tu ajoutes un autre fichier
    # if Path(source_file_2).exists():
    #     df2 = import_source_file(source_file_2, source_type=2)
    #     merged_df = merge_into_master(master_file, df2)
    #     print(f"{Fore.CYAN}[INFO]{Style.RESET_ALL} Après import fichier 2 : {len(merged_df)} clients uniques")
    #     wilaya_stats = build_wilaya_stats(merged_df)
    #     print_wilaya_stats_colored(wilaya_stats, total_clients=len(merged_df))
    #     save_wilaya_log(wilaya_stats, total_clients=len(merged_df), log_file=wilaya_log_file)

    print(f"{Fore.GREEN}[FIN]{Style.RESET_ALL} Pipeline terminé.")


if __name__ == "__main__":
    main()
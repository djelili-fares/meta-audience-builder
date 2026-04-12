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

Auteur : ChatGPT
Contexte : Shopify / COD / Algérie / Meta Ads
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional, Tuple

import pandas as pd


# ============================================================
# CONFIGURATION
# ============================================================

# Colonnes finales du fichier principal Meta
META_COLUMNS = ["phone", "fn", "ln", "ct", "st", "country", "value"]

# Pays fixe pour ton cas
DEFAULT_COUNTRY = "DZ"


# ============================================================
# OUTILS DE NORMALISATION
# ============================================================

def clean_text(value: object) -> str:
    """
    Nettoie une valeur texte :
    - gère les NaN / None
    - supprime les espaces inutiles
    - conserve le texte lisible

    Args:
        value: valeur brute

    Returns:
        Texte nettoyé
    """
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_phone_dz(phone: object) -> str:
    """
    Normalise un numéro algérien vers le format international sans '+' ni '00',
    recommandé pour Meta.

    Exemples :
    - 0670303673      -> 213670303673
    - 670303673       -> 213670303673
    - +213670303673   -> 213670303673
    - 00213670303673  -> 213670303673

    Règle métier :
    - on enlève tous les caractères non numériques
    - si le numéro commence par 0 et fait 10 chiffres -> on remplace le 0 par 213
    - si le numéro commence déjà par 213 -> on le garde
    - si le numéro commence par 00213 -> on remplace par 213
    - si le numéro fait 9 chiffres (sans 0) -> on ajoute 213 devant

    Args:
        phone: numéro brut

    Returns:
        Numéro normalisé ou chaîne vide si invalide
    """
    raw = clean_text(phone)
    if not raw:
        return ""

    # Garde uniquement les chiffres
    digits = re.sub(r"\D", "", raw)

    if not digits:
        return ""

    # Cas 00213...
    if digits.startswith("00213"):
        digits = digits[2:]  # 00213xxxx -> 213xxxx

    # Cas +213... déjà géré par la suppression des non chiffres -> 213...
    if digits.startswith("213"):
        # Vérification minimale
        return digits

    # Cas local 0XXXXXXXXX (10 chiffres)
    if digits.startswith("0") and len(digits) == 10:
        return "213" + digits[1:]

    # Cas local sans 0 : 9 chiffres
    if len(digits) == 9:
        return "213" + digits

    # Si rien ne correspond, on retourne quand même le nombre brut nettoyé
    # pour inspection manuelle éventuelle.
    return digits


def normalize_value(value: object) -> Optional[float]:
    """
    Normalise la valeur de commande.

    Args:
        value: prix brut

    Returns:
        float si possible, sinon None
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

    Logique :
    - premier mot => prénom
    - le reste => nom
    Cela évite de perdre de l'information quand il y a plus de 2 mots.

    Exemples :
    - "Zakii Redaa" -> ("Zakii", "Redaa")
    - "Zahra Benmalek" -> ("Zahra", "Benmalek")
    - "Taimo" -> ("Taimo", "")
    - "Amina Ben Ali" -> ("Amina", "Ben Ali")

    Args:
        full_name: nom complet brut

    Returns:
        Tuple (prenom, nom)
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

    Formats supportés :
    - .xlsx
    - .xls
    - .csv

    Args:
        file_path: chemin du fichier source

    Returns:
        DataFrame pandas

    Raises:
        ValueError: si extension non supportée
    """
    file_path = Path(file_path)

    if not file_path.exists():
        raise FileNotFoundError(f"Fichier introuvable : {file_path}")

    suffix = file_path.suffix.lower()

    if suffix in [".xlsx", ".xls"]:
        return pd.read_excel(file_path)

    if suffix == ".csv":
        # tentative standard UTF-8
        try:
            return pd.read_csv(file_path)
        except UnicodeDecodeError:
            # fallback fréquent
            return pd.read_csv(file_path, encoding="latin1")

    raise ValueError(f"Format non supporté : {suffix}")


# ============================================================
# TRANSFORMATIONS PAR TYPE DE FICHIER SOURCE
# ============================================================

def transform_source_type_1(df: pd.DataFrame) -> pd.DataFrame:
    """
    Transforme le premier type de fichier source :

    Colonnes attendues :
    - Prénom
    - Nom
    - Téléphone
    - Wilaya
    - Commune
    - Prix

    Retourne un DataFrame au format Meta :
    - phone
    - fn
    - ln
    - ct
    - st
    - country
    - value
    """
    required_columns = ["Prénom", "Nom", "Téléphone", "Wilaya", "Commune", "Prix"]
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(
            f"Colonnes manquantes pour le fichier type 1 : {missing}"
        )

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
    Transforme le deuxième type de fichier source :

    Colonnes attendues :
    - Client
    - Téléphone
    - Wilaya
    - Commune
    - Prix

    Le champ 'Client' est séparé en :
    - premier mot => fn
    - reste => ln
    """
    required_columns = ["Client", "Téléphone", "Wilaya", "Commune", "Prix"]
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(
            f"Colonnes manquantes pour le fichier type 2 : {missing}"
        )

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
    Nettoie le DataFrame final Meta :
    - garde uniquement les colonnes Meta
    - supprime les lignes sans téléphone
    - supprime les lignes sans prénom ET sans nom
    - déduplique sur le téléphone
    """
    df = df.copy()

    # Garde les colonnes dans l'ordre attendu
    df = df[META_COLUMNS]

    # Normalisation défensive supplémentaire
    for col in ["phone", "fn", "ln", "ct", "st", "country"]:
        df[col] = df[col].apply(clean_text)

    # Suppression lignes inutilisables
    df = df[df["phone"] != ""]
    df = df[(df["fn"] != "") | (df["ln"] != "")]

    # country forcé à DZ
    df["country"] = DEFAULT_COUNTRY

    # Déduplication sur le téléphone
    # keep="last" = la dernière occurrence importée écrase la précédente
    df = df.drop_duplicates(subset=["phone"], keep="last")

    # Réindexation
    df = df.reset_index(drop=True)

    return df


# ============================================================
# FICHIER PRINCIPAL
# ============================================================

def initialize_master_file(master_file: str | Path) -> None:
    """
    Crée le fichier principal s'il n'existe pas encore.

    Args:
        master_file: chemin du fichier principal CSV
    """
    master_path = Path(master_file)
    if not master_path.exists():
        empty_df = pd.DataFrame(columns=META_COLUMNS)
        empty_df.to_csv(master_path, index=False, encoding="utf-8-sig")
        print(f"[OK] Fichier principal créé : {master_path}")
    else:
        print(f"[INFO] Fichier principal existe déjà : {master_path}")


def load_master_file(master_file: str | Path) -> pd.DataFrame:
    """
    Charge le fichier principal CSV.
    S'il n'existe pas, il est créé vide.

    Args:
        master_file: chemin du fichier principal

    Returns:
        DataFrame du fichier principal
    """
    initialize_master_file(master_file)
    df = pd.read_csv(master_file, encoding="utf-8-sig")
    for col in META_COLUMNS:
        if col not in df.columns:
            df[col] = None
    return df[META_COLUMNS]


def save_master_file(df: pd.DataFrame, master_file: str | Path) -> None:
    """
    Sauvegarde le fichier principal en CSV UTF-8 BOM
    pour une bonne compatibilité Excel.

    Args:
        df: DataFrame à sauvegarder
        master_file: chemin du fichier de sortie
    """
    df.to_csv(master_file, index=False, encoding="utf-8-sig")
    print(f"[OK] Fichier principal sauvegardé : {master_file}")


# ============================================================
# PIPELINE D'IMPORT
# ============================================================

def import_source_file(
    source_file: str | Path,
    source_type: int,
) -> pd.DataFrame:
    """
    Importe un fichier source et le transforme au format Meta.

    Args:
        source_file: chemin du fichier source
        source_type: 1 ou 2 selon la structure du fichier

    Returns:
        DataFrame transformé et nettoyé
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

    Logique :
    - charge le fichier principal
    - concatène ancien + nouveau
    - déduplique sur téléphone
    - conserve la dernière version

    Args:
        master_file: chemin du fichier principal
        imported_df: nouvelles données à intégrer

    Returns:
        DataFrame fusionné
    """
    master_df = load_master_file(master_file)

    combined = pd.concat([master_df, imported_df], ignore_index=True)
    combined = clean_meta_dataframe(combined)

    save_master_file(combined, master_file)
    return combined


# ============================================================
# EXEMPLE D'UTILISATION
# ============================================================

def main() -> None:
    """
    Exemple complet d'utilisation.

    Adapte simplement les chemins à tes vrais fichiers.
    """
    master_file = "meta_custom_audience_master.csv"

    # Exemples de fichiers sources
    source_file_1 = "clients_livrés_zimou.xlsx"
    # source_file_2 = "clients_source_type_2.xlsx"

    # Création du fichier principal si besoin
    initialize_master_file(master_file)

    # Import du fichier 1
    if Path(source_file_1).exists():
        df1 = import_source_file(source_file_1, source_type=1)
        merged_1 = merge_into_master(master_file, df1)
        print(f"[INFO] Après import fichier 1 : {len(merged_1)} clients uniques")

    # # Import du fichier 2
    # if Path(source_file_2).exists():
    #     df2 = import_source_file(source_file_2, source_type=2)
    #     merged_2 = merge_into_master(master_file, df2)
    #     print(f"[INFO] Après import fichier 2 : {len(merged_2)} clients uniques")

    print("[FIN] Pipeline terminé.")


if __name__ == "__main__":
    main()
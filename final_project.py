"""
FINAL PROJECT - PHAN TICH HE THONG GIAO THONG CONG CONG LONDON (TfL)

Mon hoc: Nhap mon Ky thuat Du lieu
Chu de: Xay dung pipeline ETL de phan tich he thong nha ga TfL

Huong dan chay trong VSCode:
    python final_project.py

Huong dan chay trong Google Colab:
    1. Upload file final_project.py va cac file du lieu vao cung 1 folder
    2. Chay lenh: !python final_project.py

File du lieu can dat chung thu muc:
    - stations .kml hoac stations.kml
    - TfL_stations.csv hoac stations.csv
    - Stops.csv

Code duoc viet theo muc sinh vien nam 2 nganh Data Science:
    - Tach thanh tung buoc ETL ro rang
    - Dung pandas, numpy, scikit-learn, sqlite3, folium
    - In bao cao bang tieng Viet, giu lai mot so tu chuyen mon nhu ETL, KMeans, cluster
"""

from __future__ import annotations

import csv
import json
import os
import re
import sys
import sqlite3
import warnings
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.cluster import KMeans
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import StandardScaler

# Coerce stdout/stderr to UTF-8 to prevent encoding crash in Windows console
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

warnings.filterwarnings("ignore")


# ============================================================================
# 0. CAU HINH CHUNG
# ============================================================================

KML_CANDIDATES = ["stations .kml", "stations.kml"]
TFL_CANDIDATES = ["TfL_stations.csv", "stations.csv"]
STOPS_CANDIDATES = ["Stops.csv"]

YEARS = [2017, 2018, 2019, 2020, 2021]
N_CLUSTERS = 6

LINE_SUFFIXES = [
    " underground station",
    " overground station",
    " rail station",
    " dlr station",
    " elizabeth line station",
    " tram stop",
    " station",
    # Additional transport mode suffixes to align TfL and KML data
    " lu",
    " lo",
    " nr",
    " dlr",
    " el",
    " tfl",
    " (dis)",
    " (bak)",
    " (h&c)",
    " (d & p)",
    " h&c",
    " dis",
    " bak",
    " d & p",
    " bakerloo",
    " circle",
    " central",
    " hammersmith & city",
    " hammersmith and city",
]


STOPTYPE_PRIORITY = {
    "MET": 1,
    "RLY": 2,
    "RSE": 3,
    "TMU": 4,
    "DLR": 5,
    "PLT": 6,
    "RPL": 7,
}

CLUSTER_NAMES = [
    "Siêu trung tâm",
    "Ga lớn",
    "Ga trung bình",
    "Ga nhỏ",
    "Ga ít khách",
    "Ga rất ít khách",
]


@dataclass
class MergeReport:
    kml_rows: int
    tfl_rows: int
    stops_rows: int
    matched_tfl: int
    matched_stops: int


# ============================================================================
# 1. HAM HO TRO
# ============================================================================

def get_project_folder() -> Path:
    """Lay folder dang chua file code. Cach nay chay tot trong VSCode va Colab."""
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path.cwd()


def find_file(folder: Path, candidates: list[str]) -> Path:
    for file_name in candidates:
        path = folder / file_name
        if path.exists():
            return path
    raise FileNotFoundError(f"Khong tim thay file trong cac ten: {candidates}")


def ensure_package(package_name: str, import_name: str | None = None) -> bool:
    import_name = import_name or package_name
    try:
        __import__(import_name)
        return True
    except ImportError:
        print(f"[Info] Cai dat goi Python thieu: {package_name}...")
        import subprocess
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
            return True
        except subprocess.CalledProcessError as e:
            print(f"[Error] Khong cai duoc {package_name}: {e}")
            return False



def normalize_station_name(name: object) -> str:
    """Chuan hoa ten ga de merge du lieu tu nhieu nguon."""
    if pd.isna(name):
        return ""

    text = str(name).strip().lower()
    text = text.replace("&", " and ")
    text = text.replace("st.", "st")
    text = re.sub(r"\s+", " ", text)

    # Strip suffixes repeatedly to clean up multiple suffixes
    changed = True
    while changed:
        changed = False
        for suffix in LINE_SUFFIXES:
            if text.endswith(suffix):
                text = text[: -len(suffix)].strip()
                changed = True
                break

    text = re.sub(r"[^\w\s']", " ", text)
    text = re.sub(r"\s+", " ", text).strip()

    # Manual aliases for complex mismatched names
    if text in ["bank", "monument"]:
        return "bank and monument"
    if text == "crossharbour and london arena":
        return "crossharbour"
    if text in ["heathrow terminals 123", "heathrow terminals 2 and 3", "heathrow terminals 1 2 3"]:
        return "heathrow terminals 1 2 3"
    if text in ["hammersmith d and p", "hammersmith (d and p)", "hammersmith"]:
        return "hammersmith"

    return text



def first_non_empty(values: list[object]) -> str:
    for value in values:
        if pd.notna(value):
            text = str(value).strip()
            if text and text.lower() != "nan":
                return text
    return ""


def print_title(title: str) -> None:
    print("\n" + "=" * 72)
    print(title)
    print("=" * 72)


# ============================================================================
# 2. EXTRACT - DOC DU LIEU
# ============================================================================

def load_kml_stations(kml_path: Path) -> pd.DataFrame:
    """Doc file KML de lay ten ga va toa do GPS."""
    tree = ET.parse(kml_path)
    root = tree.getroot()

    # Strip namespaces to make the parser schema-agnostic
    for elem in root.iter():
        if elem.tag.startswith('{'):
            elem.tag = elem.tag.split('}', 1)[1]

    rows = []
    skipped = 0

    for placemark in root.findall(".//Placemark"):
        name_el = placemark.find("name")
        coord_el = placemark.find(".//coordinates")

        if name_el is None or coord_el is None or not name_el.text or not coord_el.text:
            skipped += 1
            continue

        parts = coord_el.text.strip().split(",")
        if len(parts) < 2:
            skipped += 1
            continue

        try:
            lon = float(parts[0])
            lat = float(parts[1])
        except ValueError:
            skipped += 1
            continue

        station = re.sub(r"\s+", " ", name_el.text.strip())
        station = re.sub(r"\s+Station$", "", station).strip()

        rows.append(
            {
                "station": station,
                "station_key": normalize_station_name(station),
                "lat": lat,
                "lon": lon,
            }
        )

    df = pd.DataFrame(rows).drop_duplicates(subset=["station_key"]).reset_index(drop=True)
    print(f"KML: doc duoc {len(df)} ga, bo qua {skipped} placemark loi")
    return df



def load_tfl_csv(tfl_path: Path) -> pd.DataFrame:
    """
    Doc file TfL_stations.csv.

    File du lieu hien tai co header dung nhung moi dong du lieu bi boc trong 1 o.
    Vi vay can tach lai bang csv.reader truoc khi dua vao DataFrame.
    """
    with open(tfl_path, "r", encoding="utf-8-sig", newline="") as handle:
        lines = handle.read().splitlines()

    if not lines:
        raise ValueError("File TfL CSV rong")

    header = next(csv.reader([lines[0]]))
    if header and header[0] == "":
        header[0] = "row_id"

    rows = []
    for line in lines[1:]:
        if not line.strip():
            continue

        first_pass = next(csv.reader([line]))

        # BUG FIX #2: Cai tien logic xu ly CSV parsing
        # Chi parse lai neu first_pass chi co 1 phan tu va cac dau phay trong do
        if len(first_pass) == len(header):
            row = first_pass
        elif len(first_pass) == 1 and "," in first_pass[0]:
            # Chi parse lai neu co dau phay, chi ra rang du lieu bi boc trong 1 o
            row = next(csv.reader([first_pass[0]]))
            if len(row) != len(header):
                # Neu sau parse van khong dung so cot, huy bo
                continue
        else:
            row = first_pass

        if len(row) < len(header):
            row = row + [""] * (len(header) - len(row))
        if len(row) > len(header):
            row = row[: len(header)]

        rows.append(row)

    df = pd.DataFrame(rows, columns=header)

    if "Station" not in df.columns:
        raise ValueError("TfL CSV thieu cot Station")

    df["station_csv"] = df["Station"].astype(str).str.strip()
    df["station_key"] = df["station_csv"].map(normalize_station_name)

    # Gop va cong don cac ga bi trung key (nhu Liverpool Street LU & NR)
    passenger_cols = [c for c in df.columns if c.startswith("En/Ex")]
    for col in passenger_cols:
        df[col] = pd.to_numeric(
            df[col].astype(str).str.replace(",", "").str.strip(),
            errors="coerce"
        ).fillna(0)

    def combine_lines(series):
        vals = []
        for val in series.dropna().astype(str):
            for sub in val.split(","):
                s_sub = sub.strip()
                if s_sub:
                    vals.append(s_sub)
        return ", ".join(sorted(set(vals)))

    agg_dict = {
        "Station": "first",
        "station_csv": "first",
    }
    for col in passenger_cols:
        agg_dict[col] = "sum"
    for col in ["LINES", "NETWORK"]:
        if col in df.columns:
            agg_dict[col] = combine_lines
    for col in ["London Underground", "Elizabeth Line", "London Overground", "DLR", "Night Tube?"]:
        if col in df.columns:
            agg_dict[col] = "first"

    df = df.groupby("station_key", as_index=False).agg(agg_dict)

    print(f"TfL CSV: doc va gop luu luong duoc {len(df)} ga doc lap")
    return df



def load_stops_csv(stops_path: Path) -> pd.DataFrame:
    """Doc Stops.csv tu NaPTAN va lay metadata can thiet."""
    df = pd.read_csv(stops_path, low_memory=False)

    required_cols = ["CommonName", "Longitude", "Latitude", "StopType"]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Stops.csv thieu cac cot: {missing_cols}")

    df["station_key"] = df["CommonName"].map(normalize_station_name)
    df["stop_lat"] = pd.to_numeric(df["Latitude"], errors="coerce")
    df["stop_lon"] = pd.to_numeric(df["Longitude"], errors="coerce")
    df["stop_priority"] = df["StopType"].map(STOPTYPE_PRIORITY).fillna(99)
    df["borough"] = df.apply(
        lambda row: first_non_empty([row.get("Town"), row.get("ParentLocalityName"), row.get("LocalityName")]),
        axis=1,
    )

    df.loc[df["station_key"].eq(""), "station_key"] = np.nan
    df.loc[df["borough"].eq(""), "borough"] = np.nan
    df = df.dropna(subset=["station_key", "stop_lat", "stop_lon"]).copy()

    print(f"Stops CSV: doc duoc {len(df)} ban ghi hop le")
    return df


# ============================================================================
# 3. TRANSFORM - GHEP, LAM SACH, TAO DAC TRUNG
# ============================================================================

def nearest_stop_match(merged_df: pd.DataFrame, stops_df: pd.DataFrame) -> pd.DataFrame:
    """
    Ghep Stops.csv theo ten ga da chuan hoa.

    Mot ten ga co the xuat hien nhieu lan trong NaPTAN, nen chon ban ghi uu tien
    StopType phu hop va gan toa do KML nhat.
    """
    candidates = pd.merge(
        merged_df[["station", "station_key", "lat", "lon"]],
        stops_df[["station_key", "CommonName", "borough", "stop_lat", "stop_lon", "StopType", "stop_priority"]],
        on="station_key",
        how="left",
    )

    # BUG FIX #1: Kiem tra xem co match thanh cong khong (khong can kiem tra empty cho left merge)
    # Thay vao do, kiem tra xem cot CommonName co NaN khong
    if candidates["CommonName"].isna().all():
        merged_df["borough"] = np.nan
        merged_df["stop_name"] = np.nan
        merged_df["stop_type"] = np.nan
        merged_df["stop_distance"] = np.nan
        return merged_df

    candidates["stop_distance"] = np.sqrt(
        (candidates["lat"] - candidates["stop_lat"]) ** 2
        + (candidates["lon"] - candidates["stop_lon"]) ** 2
    )
    candidates = candidates.sort_values(["station", "stop_priority", "stop_distance"])
    best = candidates.drop_duplicates(subset=["station"], keep="first")

    result = pd.merge(
        merged_df,
        best[["station", "borough", "CommonName", "StopType", "stop_distance"]],
        on="station",
        how="left",
    )
    result = result.rename(columns={"CommonName": "stop_name", "StopType": "stop_type"})
    return result


def merge_sources(kml_df: pd.DataFrame, tfl_df: pd.DataFrame, stops_df: pd.DataFrame) -> tuple[pd.DataFrame, MergeReport]:
    """Ghep 3 nguon: KML + TfL CSV + Stops CSV."""
    df = pd.merge(kml_df, tfl_df, on="station_key", how="left")

    probe_cols = [col for col in ["Station", "LINES", "NETWORK", "En/Ex 2021"] if col in df.columns]
    matched_tfl = int(df[probe_cols].notna().any(axis=1).sum()) if probe_cols else 0

    df = nearest_stop_match(df, stops_df)
    matched_stops = int(df["borough"].notna().sum()) if "borough" in df.columns else 0

    report = MergeReport(
        kml_rows=len(kml_df),
        tfl_rows=len(tfl_df),
        stops_rows=len(stops_df),
        matched_tfl=matched_tfl,
        matched_stops=matched_stops,
    )
    return df, report


def clean_and_engineer(df: pd.DataFrame) -> pd.DataFrame:
    """Lam sach du lieu va tao cac cot phan tich."""
    df = df.copy()

    passenger_cols = []
    for year in YEARS:
        raw_col = f"En/Ex {year}"
        new_col = f"passengers_{year}"

        if raw_col in df.columns:
            df[new_col] = pd.to_numeric(
                df[raw_col].astype(str).str.replace(",", "").str.strip(),
                errors="coerce",
            )
        else:
            df[new_col] = np.nan

        passenger_cols.append(new_col)

    if "LINES" in df.columns:
        line_text = df["LINES"].fillna("").astype(str).str.strip()
        df["num_lines"] = np.where(line_text.eq(""), 0, line_text.str.count(",") + 1)
    else:
        df["num_lines"] = 0

    df["avg_passengers"] = df[passenger_cols].mean(axis=1)

    # BUG FIX #3: Cai tien xu ly division by zero
    # Thay vi replace 0 bang NaN roi chia, hay dung dieu kien de tranh loi
    # Neu passengers_2019 = 0 hoac NaN, gan covid_impact_pct = NaN
    df["covid_impact_pct"] = np.where(
        (df["passengers_2019"] == 0) | df["passengers_2019"].isna(),
        np.nan,
        (df["passengers_2020"] - df["passengers_2019"]) / df["passengers_2019"] * 100
    )
    # Tuong tu cho recovery_rate_pct
    df["recovery_rate_pct"] = np.where(
        (df["passengers_2020"] == 0) | df["passengers_2020"].isna(),
        np.nan,
        (df["passengers_2021"] - df["passengers_2020"]) / df["passengers_2020"] * 100
    )

    df = add_trend_analysis(df)

    before_rows = len(df)
    df = df.dropna(subset=["lat", "lon", "passengers_2021", "num_lines"]).copy()
    df["num_lines"] = df["num_lines"].astype(int)

    print(f"Sau lam sach: giu lai {len(df)}/{before_rows} ga du dieu kien phan tich")
    return df


def add_trend_analysis(df: pd.DataFrame) -> pd.DataFrame:
    """Tinh xu huong hanh khach bang LinearRegression tren giai doan 2017-2021."""
    df = df.copy()
    year_matrix = np.array(YEARS).reshape(-1, 1)
    passenger_cols = [f"passengers_{year}" for year in YEARS]

    slopes = []
    for _, row in df.iterrows():
        passengers = row[passenger_cols].to_numpy(dtype=float)
        # BUG FIX #4: Cai tien xu ly NaN trong trend analysis
        # Thay vi bo qua toan bo dong neu co NaN, hay co the su dung cac nam co du lieu
        # Loc ra cac nam co du lieu (khong NaN va khong 0)
        valid_mask = ~np.isnan(passengers) & (passengers > 0)
        
        if valid_mask.sum() < 2:
            # Neu co it hon 2 diem du lieu, khong the fit duong thang
            slopes.append(np.nan)
            continue

        # Chi fit voi cac diem du lieu hop le
        valid_years = year_matrix[valid_mask]
        valid_passengers = passengers[valid_mask]
        
        model = LinearRegression()
        model.fit(valid_years, valid_passengers)
        slopes.append(float(model.coef_[0]))

    def trend_label(slope: float) -> str:
        if pd.isna(slope):
            return "Chưa rõ"
        if slope > 150_000:
            return "Tăng mạnh"
        if slope > 20_000:
            return "Tăng nhẹ"
        if slope >= -20_000:
            return "Ổn định"
        if slope > -150_000:
            return "Giảm nhẹ"
        return "Giảm mạnh"

    df["trend_slope"] = slopes
    df["trend_category"] = df["trend_slope"].apply(trend_label)
    return df


# ============================================================================
# 4. CLUSTERING VA PHAN TICH
# ============================================================================

def run_kmeans_clustering(df: pd.DataFrame, n_clusters: int = N_CLUSTERS) -> pd.DataFrame:
    """Phan cum nha ga bang KMeans."""
    df = df.copy()
    features = ["passengers_2021", "num_lines", "lat", "lon"]
    matrix = df[features].to_numpy(dtype=float)
    matrix_scaled = StandardScaler().fit_transform(matrix)

    actual_clusters = min(n_clusters, len(df))
    if actual_clusters < 1:
        raise ValueError("Khong du du lieu de chay KMeans")

    # BUG FIX #5: Canh bao neu so cum bi giam
    if actual_clusters < n_clusters:
        print(f"\n[CANH BAO] So cum yeu cau {n_clusters} nhung du lieu chi co {len(df)} ga.")
        print(f"[CANH BAO] Giam so cum xuong {actual_clusters} de phu hop.\n")

    model = KMeans(n_clusters=actual_clusters, random_state=42, n_init=10)
    df["cluster_id"] = model.fit_predict(matrix_scaled)

    cluster_means = (
        df.groupby("cluster_id")["passengers_2021"]
        .mean()
        .sort_values(ascending=False)
    )
    name_map = {}
    for rank, cluster_id in enumerate(cluster_means.index.tolist()):
        name_map[int(cluster_id)] = CLUSTER_NAMES[rank] if rank < len(CLUSTER_NAMES) else f"Cum {rank + 1}"

    df["cluster_name"] = df["cluster_id"].map(name_map)
    return df


def create_cluster_summary(df: pd.DataFrame) -> pd.DataFrame:
    return (
        df.groupby(["cluster_id", "cluster_name"], as_index=False)
        .agg(
            so_ga=("station", "count"),
            hanh_khach_tb_2021=("passengers_2021", "mean"),
            tong_hanh_khach_2021=("passengers_2021", "sum"),
            so_tuyen_tb=("num_lines", "mean"),
            covid_impact_tb=("covid_impact_pct", "mean"),
            recovery_tb=("recovery_rate_pct", "mean"),
        )
        .sort_values("hanh_khach_tb_2021", ascending=False)
        .round(2)
    )


def make_sqlite_safe_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Tao ban sao DataFrame co ten cot an toan cho SQLite.

    SQLite khong phan biet hoa/thuong trong ten cot, nen `Station` va `station`
    bi xem la trung nhau. CSV/Excel van giu ten cot goc, chi SQLite dung ban nay.
    """
    safe_df = df.copy()
    new_columns = []
    used_columns = set()

    for col in safe_df.columns:
        clean_col = str(col).strip().lower()
        clean_col = re.sub(r"[^a-z0-9_]+", "_", clean_col)
        clean_col = re.sub(r"_+", "_", clean_col).strip("_")
        if not clean_col:
            clean_col = "column"

        base_col = clean_col
        counter = 2
        while clean_col in used_columns:
            clean_col = f"{base_col}_{counter}"
            counter += 1

        used_columns.add(clean_col)
        new_columns.append(clean_col)

    safe_df.columns = new_columns
    return safe_df


def print_analysis_report(df: pd.DataFrame, merge_report: MergeReport) -> None:
    print_title("BAO CAO TOM TAT PIPELINE")

    print("1. Extract - Thu thap du lieu")
    print(f"   KML: {merge_report.kml_rows} ga")
    print(f"   TfL CSV: {merge_report.tfl_rows} dong")
    print(f"   Stops CSV: {merge_report.stops_rows} ban ghi hop le")

    print("\n2. Transform - Lam sach va ghep du lieu")
    print(f"   So ga ghep duoc voi TfL CSV: {merge_report.matched_tfl}/{merge_report.kml_rows}")
    print(f"   So ga ghep duoc voi Stops CSV: {merge_report.matched_stops}/{merge_report.kml_rows}")
    print(f"   So ga duoc dua vao phan tich: {len(df)}")

    print("\n3. Clustering - Phan cum KMeans")
    cluster_summary = create_cluster_summary(df)
    for _, row in cluster_summary.iterrows():
        print(
            f"   Cluster {int(row['cluster_id'])} - {row['cluster_name']}: "
            f"{int(row['so_ga'])} ga, "
            f"hanh khach TB 2021 = {row['hanh_khach_tb_2021']:,.0f}, "
            f"so tuyen TB = {row['so_tuyen_tb']:.2f}"
        )

    print("\n4. COVID Analysis")
    total_2019 = df["passengers_2019"].sum()
    total_2020 = df["passengers_2020"].sum()
    total_2021 = df["passengers_2021"].sum()
    covid_change = (total_2020 - total_2019) / total_2019 * 100 if total_2019 else np.nan
    recovery_change = (total_2021 - total_2020) / total_2020 * 100 if total_2020 else np.nan
    print(f"   Tong hanh khach 2019: {total_2019:,.0f}")
    print(f"   Tong hanh khach 2020: {total_2020:,.0f}")
    print(f"   Tong hanh khach 2021: {total_2021:,.0f}")
    print(f"   Tac dong COVID-19 nam 2019 -> 2020: {covid_change:.2f}%")
    print(f"   Muc phuc hoi 2020 -> 2021: {recovery_change:.2f}%")

    print("\n5. Trend Analysis")
    trend_counts = df["trend_category"].value_counts()
    for trend_name, count in trend_counts.items():
        print(f"   {trend_name}: {count} ga")

    top5 = df.nlargest(5, "passengers_2021")[["station", "passengers_2021", "cluster_name"]]
    print("\nTop 5 ga co hanh khach 2021 cao nhat:")
    for idx, (_, row) in enumerate(top5.iterrows(), start=1):
        print(f"   {idx}. {row['station']} - {row['passengers_2021']:,.0f} ({row['cluster_name']})")

    print("\nKet luan:")
    print("   Pipeline da hoan thanh dung cac buoc trong bao cao: Extract, Transform, Load,")
    print("   KMeans clustering, COVID Analysis va Visualization. Ket qua cho thay luu luong")
    print("   hanh khach giam manh trong nam 2020 do COVID-19 va co dau hieu phuc hoi nam 2021.")


def build_summary_text(df: pd.DataFrame) -> str:
    """Tao noi dung tom tat de luu ra file txt."""
    total_2019 = df["passengers_2019"].sum()
    total_2020 = df["passengers_2020"].sum()
    total_2021 = df["passengers_2021"].sum()
    covid_change = (total_2020 - total_2019) / total_2019 * 100 if total_2019 else np.nan
    recovery_change = (total_2021 - total_2020) / total_2020 * 100 if total_2020 else np.nan

    lines = [
        "BAO CAO TOM TAT - PHAN TICH HE THONG TfL",
        "=" * 60,
        f"So ga duoc phan tich: {len(df)}",
        f"Tong hanh khach 2019: {total_2019:,.0f}",
        f"Tong hanh khach 2020: {total_2020:,.0f}",
        f"Tong hanh khach 2021: {total_2021:,.0f}",
        f"Tac dong COVID-19 2019 -> 2020: {covid_change:.2f}%",
        f"Muc phuc hoi 2020 -> 2021: {recovery_change:.2f}%",
        "",
        "THONG KE CLUSTER",
        "-" * 60,
    ]

    for _, row in create_cluster_summary(df).iterrows():
        lines.append(
            f"Cluster {int(row['cluster_id'])} - {row['cluster_name']}: "
            f"{int(row['so_ga'])} ga, "
            f"hanh khach TB 2021 = {row['hanh_khach_tb_2021']:,.0f}, "
            f"COVID impact TB = {row['covid_impact_tb']:.2f}%"
        )

    lines.extend(["", "XU HUONG HANH KHACH", "-" * 60])
    for trend_name, count in df["trend_category"].value_counts().items():
        lines.append(f"{trend_name}: {count} ga")

    lines.extend(
        [
            "",
            "KET LUAN",
            "-" * 60,
            "Pipeline da thuc hien day du cac buoc trong bao cao do an:",
            "Extract, Transform, Load, KMeans clustering, COVID Analysis, Trend Analysis va Visualization.",
            "Ket qua cho thay nam 2020 co muc sut giam hanh khach ro ret do COVID-19,",
            "sau do co dau hieu phuc hoi vao nam 2021.",
        ]
    )
    return "\n".join(lines)


# ============================================================================
# 5. LOAD - LUU TRU KET QUA
# ============================================================================

def save_outputs(df: pd.DataFrame, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    csv_path = output_dir / "london_tfl_cleaned.csv"
    excel_path = output_dir / "london_tfl_results.xlsx"
    db_path = output_dir / "london_tfl.db"
    summary_path = output_dir / "analysis_summary.txt"
    cluster_summary = create_cluster_summary(df)

    df.to_csv(csv_path, index=False)
    print(f"Da luu CSV: {csv_path}")

    # Tu dong kiem tra va cai dat openpyxl neu thieu
    ensure_package("openpyxl")

    try:
        import openpyxl
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="All_Stations", index=False)
            cluster_summary.to_excel(writer, sheet_name="Cluster_Summary", index=False)
        print(f"Da luu Excel: {excel_path}")
    except ImportError:
        print("Chua cai openpyxl va khong the tu dong cai dat, bo qua file Excel.")

    summary_path.write_text(build_summary_text(df), encoding="utf-8")
    print(f"Da luu bao cao tom tat: {summary_path}")

    sqlite_df = make_sqlite_safe_dataframe(df)
    sqlite_cluster_summary = make_sqlite_safe_dataframe(cluster_summary)
    with sqlite3.connect(db_path) as conn:
        sqlite_df.to_sql("fact_stations", conn, if_exists="replace", index=False)
        sqlite_cluster_summary.to_sql("dim_clusters", conn, if_exists="replace", index=False)
    print(f"Da luu SQLite: {db_path}")


def log_unmatched_stations(kml_df: pd.DataFrame, final_df: pd.DataFrame, output_dir: Path) -> None:
    """Ghi log danh sach cac ga trong file KML khong the ghep du lieu."""
    output_dir.mkdir(parents=True, exist_ok=True)
    log_path = output_dir / "unmatched_stations_log.txt"

    final_keys = set(final_df["station_key"].tolist())
    unmatched = kml_df[~kml_df["station_key"].isin(final_keys)].copy()

    lines = [
        "DANH SACH CAC GA TRONG FILE KML BI LOAI BO HOAC KHONG MATCH DU LIEU PIPELINE",
        "=" * 80,
        f"Tong so ga trong KML ban dau: {len(kml_df)}",
        f"So ga duoc giu lai phân tích: {len(final_df)}",
        f"So ga bi loai bo: {len(unmatched)}",
        "",
        "Chi tiet cac ga bi loai bo (do thieu toa do, thieu tuyen, hoac thieu du lieu khach 2021):",
        "-" * 80,
    ]

    for idx, (_, row) in enumerate(unmatched.iterrows(), start=1):
        lines.append(f"{idx:02d}. Tên ga: {row['station']} (lat: {row['lat']}, lon: {row['lon']})")

    log_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"Da luu log cac ga khong khop: {log_path}")



# ============================================================================
# 6. VISUALIZATION - BAN DO FOLIUM
# ============================================================================

def create_folium_map(df: pd.DataFrame, output_dir: Path) -> None:
    map_path = output_dir / "london_tfl_map.html"
    center_lat = float(df["lat"].mean())
    center_lon = float(df["lon"].mean())

    stations = []
    for _, row in df.iterrows():
        stations.append(
            {
                "station": str(row["station"]),
                "cluster_name": str(row["cluster_name"]),
                "borough": str(row["borough"]) if pd.notna(row.get("borough")) else "Không xác định",
                "lat": float(row["lat"]),
                "lon": float(row["lon"]),
                "num_lines": int(row["num_lines"]),
                "lines": str(row["LINES"]) if "LINES" in row and pd.notna(row["LINES"]) else "",
                "trend_category": str(row["trend_category"]),
                "passengers_2017": int(row["passengers_2017"]) if pd.notna(row["passengers_2017"]) else 0,
                "passengers_2018": int(row["passengers_2018"]) if pd.notna(row["passengers_2018"]) else 0,
                "passengers_2019": int(row["passengers_2019"]) if pd.notna(row["passengers_2019"]) else 0,
                "passengers_2020": int(row["passengers_2020"]) if pd.notna(row["passengers_2020"]) else 0,
                "passengers_2021": int(row["passengers_2021"]) if pd.notna(row["passengers_2021"]) else 0,
            }
        )

    categories = sorted({station["cluster_name"] for station in stations})
    station_data = json.dumps(stations, ensure_ascii=False)
    category_data = json.dumps(categories, ensure_ascii=False)

    html = """<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Bản đồ tương tác TfL London</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/leaflet@1.9.3/dist/leaflet.css" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.2.2/dist/css/bootstrap.min.css" />
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
        html, body { height: 100vh; overflow: hidden; margin: 0; padding: 0; background: #080d18; color: #e2e8f0; }
        #app-shell { position: relative; display: flex; height: 100vh; overflow: hidden; font-family: Inter, Arial, sans-serif; }
        #sidebar { width: 360px; min-width: 360px; max-width: 360px; background: linear-gradient(180deg, #f8fafc, #eef2f7); border-radius: 28px; box-shadow: 0 28px 60px rgba(15, 23, 42, 0.18); margin: 1rem; padding: 1rem; overflow-y: auto; transition: all 0.3s ease; color: #1e293b; }
        #sidebar.collapsed { width: 0; min-width: 0; margin: 1rem 0; padding: 0; opacity: 0; pointer-events: none; }
        #sidebar.collapsed .sidebar-card { opacity: 0; }
        #sidebar h4 { margin-top: 0; font-size: 1.2rem; font-weight: 700; color: #0f172a; }
        #toggle-sidebar-btn { position: absolute; top: 1rem; left: 360px; width: 40px; height: 40px; border: none; border-radius: 999px; background: #0d6efd; color: #ffffff; font-size: 1rem; cursor: pointer; box-shadow: 0 12px 28px rgba(13, 110, 253, 0.24); display: flex; align-items: center; justify-content: center; transition: left 0.3s ease, transform 0.3s ease; z-index: 1000; }
        .sidebar-card.collapsed-card > *:not(.sidebar-section-title) { display: none; }
        .sidebar-section-title { cursor: pointer; position: relative; }
        .sidebar-section-title:after { content: '▼'; position: absolute; right: 0.6rem; top: 0; font-size: 0.78rem; transform: translateY(3px); transition: transform 0.2s ease; }
        .sidebar-card.collapsed-card .sidebar-section-title:after { transform: rotate(-90deg) translateY(0); }
        #sidebar.collapsed + #toggle-sidebar-btn { left: 0; transform: rotate(180deg); }
        .custom-cluster-icon {
            background: linear-gradient(180deg, rgba(59,130,246,0.95), rgba(37,99,235,0.95));
            color: #ffffff;
            border-radius: 50%;
            width: 44px;
            height: 44px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 13px;
            font-weight: 800;
            box-shadow: 0 6px 18px rgba(2,6,23,0.28);
            border: 2px solid rgba(255, 255, 255, 0.18);
        }
        .custom-cluster-icon.cluster-small { width: 40px; height: 40px; font-size: 12px; }
        .custom-cluster-icon.cluster-medium { width: 52px; height: 52px; font-size: 14px; }
        .custom-cluster-icon.cluster-large { width: 66px; height: 66px; font-size: 16px; }
        .custom-cluster-icon .count { display:block; line-height:1; }
        .custom-cluster-icon .sub { display:block; font-size:10px; opacity:0.9; margin-top:2px; }
        .leaflet-tooltip.my-tooltip { background: rgba(2,6,23,0.95); color: #fff; border-radius:6px; padding:6px 8px; border: none; box-shadow: 0 6px 18px rgba(2,6,23,0.28); }
        .slider-group label, .slider-group .form-label, .slider-group .text-secondary, .slider-group strong { color: #333333 !important; opacity: 1 !important; }
        .filter-chip-container { display: flex; flex-wrap: wrap; gap: 0.5rem; }
        .filter-chip { display: inline-flex; align-items: center; justify-content: center; padding: 0.55rem 0.9rem; border-radius: 999px; border: 1px solid var(--chip-color, #cbd5e1); background: #f8fafc; color: var(--chip-color, #334155); cursor: pointer; transition: all 0.18s ease; font-size: 0.9rem; }
        .filter-chip.active { background: var(--chip-color, #0d6efd); color: #ffffff; border-color: var(--chip-color, #0d6efd); box-shadow: 0 10px 24px rgba(13, 110, 253, 0.18); }
        .sidebar-card { background: #ffffff; border-radius: 20px; padding: 1rem; margin-bottom: 1rem; box-shadow: 0 10px 28px rgba(15, 23, 42, 0.08); }
        .sidebar-section-title { margin-bottom: 0.8rem; font-size: 0.82rem; letter-spacing: 0.12em; text-transform: uppercase; color: #64748b; }
        #map { flex: 1; height: 100%; background: #080d18; }
        .slider-group { padding: 0.85rem 0.85rem 0.75rem; background: #f1f5f9; border: 1px solid #e2e8f0; border-radius: 18px; }
        input[type=range].form-range { accent-color: #0d6efd; height: 6px; background: #cbd5e1; border-radius: 999px; margin-top: 0.8rem; }
        input[type=range].form-range::-webkit-slider-thumb { appearance: none; width: 22px; height: 22px; border-radius: 50%; background: #0d6efd; border: 3px solid #ffffff; box-shadow: 0 10px 22px rgba(13, 110, 253, 0.28); }
        input[type=range].form-range::-moz-range-thumb { width: 22px; height: 22px; border-radius: 50%; background: #0d6efd; border: 3px solid #ffffff; box-shadow: 0 10px 22px rgba(13, 110, 253, 0.28); }
        .slider-group label, .slider-group .form-label, .slider-group .text-secondary, .slider-group strong { color: #333333 !important; opacity: 1 !important; }
        #app-shell { position: relative; }
        .filter-chip-container { display: flex; flex-wrap: wrap; gap: 0.5rem; }
        .filter-chip { display: inline-flex; align-items: center; justify-content: center; padding: 0.55rem 0.9rem; border-radius: 999px; border: 1px solid var(--chip-color, #cbd5e1); background: #f8fafc; color: var(--chip-color, #334155); cursor: pointer; transition: all 0.18s ease; font-size: 0.9rem; }
        .filter-chip.active { background: var(--chip-color, #0d6efd); color: #ffffff; border-color: var(--chip-color, #0d6efd); box-shadow: 0 10px 24px rgba(0, 0, 0, 0.12); }
        #toggle-sidebar-btn { position: absolute; top: 1rem; left: 360px; width: 40px; height: 40px; border: none; border-radius: 999px; background: #0d6efd; color: #ffffff; font-size: 1rem; cursor: pointer; box-shadow: 0 12px 28px rgba(13, 110, 253, 0.24); display: flex; align-items: center; justify-content: center; transition: left 0.3s ease, transform 0.3s ease; z-index: 1000; }
        #sidebar.collapsed + #toggle-sidebar-btn { left: 0; transform: rotate(180deg); }
        .btn-reset { width: 100%; font-weight: 700; }
        .legend-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 0.6rem; margin-top: 0.6rem; }
        .legend-item { display: flex; align-items: center; gap: 0.55rem; font-size: 0.88rem; color: #334155; }
        .legend-swatch { width: 14px; height: 14px; border-radius: 5px; display: inline-block; box-shadow: 0 0 0 1px rgba(15, 23, 42, 0.08); }
        .analytics-summary { display: grid; grid-template-columns: 1fr 1fr; gap: 0.8rem; margin-bottom: 1rem; }
        .analytics-card { background: #f8fafc; border-radius: 18px; padding: 0.95rem 1rem; border: 1px solid #e2e8f0; }
        .analytics-value { font-size: 1.7rem; font-weight: 800; line-height: 1; color: #0f172a; }
        .analytics-label { margin-top: 0.3rem; color: #64748b; font-size: 0.82rem; }
        .analytics-summary span, .analytics-summary text, .chart-bar-row, .chart-bar-row span, .station-label, .station-label span { color: #0f172a !important; }
        .chart-bar-row { display: flex; align-items: center; gap: 0.5rem; margin-bottom: 0.65rem; font-size: 0.9rem; }
        .chart-bar { flex: 1; background: #e2e8f0; border-radius: 999px; height: 10px; overflow: hidden; }
        .chart-bar div { height: 100%; background: #0d6efd; border-radius: 999px; }
        .station-bar-row { display: grid; gap: 0.35rem; margin-bottom: 0.95rem; }
        .station-label { display: flex; justify-content: space-between; font-size: 0.9rem; font-weight: 600; color: #0f172a; }
        .mini-bar { height: 10px; background: #e2e8f0; border-radius: 999px; overflow: hidden; }
        .mini-bar-content { height: 100%; border-radius: 999px; background: linear-gradient(90deg, #0d6efd, #38bdf8); }
        .insight-card { background: #f8fafc; border-radius: 18px; padding: 1rem; border: 1px solid #e2e8f0; }
        .insight-card p { margin-bottom: 0.7rem; color: #334155; font-size: 0.92rem; }
        .insight-card strong { color: #0d6efd; }
        .trend-overview { display: grid; gap: 0.5rem; margin-top: 0.6rem; }
        .trend-overview-item { display: flex; justify-content: space-between; align-items: center; padding: 0.45rem 0.75rem; background: #ffffff; border-radius: 12px; border: 1px solid #eef2f7; color: #0f172a; font-size: 0.9rem; }
        .trend-overview-item .swatch { width: 12px; height: 12px; border-radius: 4px; display: inline-block; margin-right: 0.6rem; box-shadow: 0 0 0 1px rgba(15,23,42,0.04); }
        .trend-overview-item .label { display: inline-flex; align-items: center; gap: 0.4rem; font-weight: 600; }
        .leaflet-popup-content-wrapper { border-radius: 1rem; box-shadow: 0 16px 48px rgba(15, 23, 42, 0.18); }
        .station-popup .leaflet-popup-content { padding: 1rem; }
        .popup-card { font-family: Inter, Arial, sans-serif; font-size: 13px; color: #0f172a; }
        .popup-title { margin: 0 0 0.5rem; font-size: 16px; font-weight: 700; }
        .popup-subtitle { display: flex; flex-wrap: wrap; gap: 0.5rem; margin-bottom: 0.8rem; color: #475569; font-size: 12px; }
        .popup-badge { display: inline-flex; align-items: center; padding: 0.25rem 0.6rem; border-radius: 999px; background: rgba(13, 110, 253, 0.12); color: #0d6efd; font-weight: 700; font-size: 11px; }
        .popup-field { display: flex; justify-content: space-between; gap: 0.75rem; margin-bottom: 0.35rem; }
        .popup-field-label { color: #475569; }
        .popup-trend { font-weight: 700; color: #0d6efd; }
        .modal-content { background: #ffffff; color: #0f172a; border-radius: 1rem; border: none; box-shadow: 0 24px 80px rgba(15, 23, 42, 0.18); }
        .modal-header, .modal-footer { border: none; }
        .modal-title { color: #0f172a; }
        .modal-body { color: #334155; line-height: 1.6; }
        .modal-body ul { padding-left: 1.2rem; margin: 0; }
        .modal-body li { margin-bottom: 0.75rem; }
        #station-comment { padding: 1rem; }
        .station-comment-card { display: flex; flex-wrap: wrap; gap: 1rem; align-items: flex-start; }
        .station-comment-summary { flex: 1 1 240px; min-width: 240px; }
        .station-comment-sparkline { width: 100%; max-width: 220px; }

        /* Custom scrollbar for sidebar */
        #sidebar::-webkit-scrollbar { width: 6px; }
        #sidebar::-webkit-scrollbar-track { background: transparent; }
        #sidebar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 3px; }
        #sidebar::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
        
        #sidebar.dark-theme::-webkit-scrollbar-thumb { background: #475569; }
        #sidebar.dark-theme::-webkit-scrollbar-thumb:hover { background: #64748b; }

        /* Dark theme styles */
        #sidebar.dark-theme {
            background: linear-gradient(180deg, #0f172a, #1e293b);
            color: #f8fafc;
            box-shadow: 0 28px 60px rgba(0, 0, 0, 0.5);
        }
        #sidebar.dark-theme h4 {
            color: #f8fafc;
        }
        #sidebar.dark-theme .sidebar-card {
            background: #1e293b;
            color: #cbd5e1;
            box-shadow: 0 10px 28px rgba(0, 0, 0, 0.2);
            border: 1px solid rgba(255, 255, 255, 0.05);
        }
        #sidebar.dark-theme .sidebar-section-title {
            color: #94a3b8;
        }
        #sidebar.dark-theme .slider-group {
            background: #0f172a;
            border-color: rgba(255, 255, 255, 0.05);
        }
        #sidebar.dark-theme .slider-group label,
        #sidebar.dark-theme .slider-group .form-label,
        #sidebar.dark-theme .slider-group .text-secondary,
        #sidebar.dark-theme .slider-group strong {
            color: #cbd5e1 !important;
        }
        #sidebar.dark-theme .filter-chip {
            background: #0f172a;
            color: #94a3b8;
            border-color: rgba(255, 255, 255, 0.1);
        }
        #sidebar.dark-theme .filter-chip.active {
            background: var(--chip-color, #0d6efd);
            color: #ffffff;
            border-color: var(--chip-color, #0d6efd);
        }
        #sidebar.dark-theme .analytics-card {
            background: #0f172a;
            border-color: rgba(255, 255, 255, 0.05);
        }
        #sidebar.dark-theme .analytics-value {
            color: #f8fafc;
        }
        #sidebar.dark-theme .analytics-label {
            color: #94a3b8;
        }
        #sidebar.dark-theme .chart-bar {
            background: #0f172a;
        }
        #sidebar.dark-theme .chart-bar-row,
        #sidebar.dark-theme .chart-bar-row span,
        #sidebar.dark-theme .station-label,
        #sidebar.dark-theme .station-label span {
            color: #f8fafc !important;
        }
        #sidebar.dark-theme .mini-bar {
            background: #0f172a;
        }
        #sidebar.dark-theme .insight-card {
            background: #0f172a;
            border-color: rgba(255, 255, 255, 0.05);
        }
        #sidebar.dark-theme .insight-card p {
            color: #cbd5e1;
        }
        #sidebar.dark-theme .trend-overview-item {
            background: #1e293b;
            border-color: rgba(255, 255, 255, 0.05);
            color: #cbd5e1;
        }
        #sidebar.dark-theme #station-comment {
            color: #cbd5e1;
        }
        #sidebar.dark-theme #basemap-select,
        #sidebar.dark-theme #search-scope,
        #sidebar.dark-theme #station-search,
        #sidebar.dark-theme #borough-filter {
            background-color: #0f172a;
            color: #f8fafc;
            border-color: rgba(255, 255, 255, 0.1);
        }
        #sidebar.dark-theme #basemap-select option,
        #sidebar.dark-theme #search-scope option,
        #sidebar.dark-theme #borough-filter option {
            background-color: #0f172a;
            color: #f8fafc;
        }
        #sidebar label {
            color: #1e293b;
            font-weight: 500;
        }
        #sidebar.dark-theme label {
            color: #cbd5e1;
        }
    </style>
</head>
<body>
    <div id="app-shell">
        <div id="sidebar">
            <h4>Bảng điều khiển TfL London</h4>
            <div class="sidebar-card">
                <div class="sidebar-section-title">Nền bản đồ</div>
                <div style="margin-top:.6rem;">
                    <select id="basemap-select" class="form-select form-select-sm">
                        <option value="default">Default (Light)</option>
                        <option value="dark">Dark</option>
                        <option value="transport">Giao thông</option>
                    </select>
                </div>
            </div>
            <div class="sidebar-card">
                <div class="sidebar-section-title">KHOẢNG THỜI GIAN</div>
                <div class="slider-group">
                    <label class="form-label mb-2">Chọn năm phân tích</label>
                    <input id="time-slider" class="form-range" type="range" min="2017" max="2021" step="1" value="2021" />
                    <div class="d-flex justify-content-between align-items-center mt-2 text-secondary small">
                        <span>2017</span>
                        <strong id="time-slider-value">2021</strong>
                        <span>2021</span>
                    </div>
                </div>
            </div>
            <div class="sidebar-card">
                <div class="sidebar-section-title">Bộ lọc cụm ga</div>
                <div id="category-filters" class="filter-chip-container"></div>
                <div style="margin-top:0.6rem; display:flex; gap:.5rem; align-items:center;">
                    <select id="search-scope" class="form-select form-select-sm" style="width:130px;">
                        <option value="name">Tên</option>
                        <option value="line">Tuyến</option>
                        <option value="borough">Borough</option>
                    </select>
                    <div style="flex:1; display:flex; gap:.4rem; align-items:center;">
                        <input id="station-search" class="form-control form-control-sm" type="search" placeholder="Tìm (ví dụ: Oxford)" />
                        <button id="clear-search" class="btn btn-sm btn-outline-secondary" title="Xóa">✕</button>
                    </div>
                    <button id="locate-btn" class="btn btn-outline-primary btn-sm" title="Tìm và di chuyển tới ga">Tìm</button>
                </div>
                <div style="margin-top:0.65rem;">
                    <select id="borough-filter" class="form-select form-select-sm">
                        <option value="">-- Lọc theo borough --</option>
                    </select>
                </div>
                <button id="reset-filters" class="btn btn-primary btn-reset mt-3">Đặt lại bộ lọc</button>
                <button id="export-csv-btn" class="btn btn-outline-success w-100 mt-2">Xuất CSV dữ liệu đang lọc</button>
            </div>
            <div class="sidebar-card">
                <div class="sidebar-section-title">Chú thích màu sắc</div>
                <div id="cluster-legend" class="legend-grid"></div>
            </div>
            <div class="sidebar-card">
                <div class="sidebar-section-title">Bản đồ nhiệt</div>
                <div style="margin-top:0.6rem; display:flex; align-items:center; gap:.5rem;">
                    <input id="heatmap-toggle" type="checkbox" />
                    <label for="heatmap-toggle" style="margin:0;">Bật bản đồ nhiệt</label>
                </div>
            </div>
            <div class="sidebar-card">
                <div class="sidebar-section-title">THỐNG KÊ CHI TIẾT</div>
                <div class="analytics-summary">
                    <div class="analytics-card">
                        <div class="analytics-label">Tổng số ga</div>
                        <div class="analytics-value" id="summary-count">0</div>
                    </div>
                    <div class="analytics-card">
                        <div class="analytics-label">Trung bình hành khách</div>
                        <div class="analytics-value" id="summary-passengers">0</div>
                    </div>
                </div>
                <div id="analyticsChart"></div>
            </div>
            <div class="sidebar-card">
                <div class="sidebar-section-title">TOP GA ĐÔNG KHÁCH NHẤT</div>
                <div id="top-stations" class="mini-bar-row"></div>
            </div>
            <div class="sidebar-card">
                <div class="sidebar-section-title">Nhận xét xu hướng</div>
                <div id="data-insights" class="insight-card">
                    <div id="insight-summary">
                        <p><strong>Nhấn vào một ga trên bản đồ</strong> để xem nhận xét chi tiết.</p>
                    </div>
                    <div id="trend-overview" style="margin-top: 0.85rem; color: #334155; font-size: 0.92rem;"></div>
                </div>
            </div>
            <div class="sidebar-card">
                <div class="sidebar-section-title">Nhận xét ga được chọn</div>
                <div id="station-comment" class="insight-card" style="color: #334155; font-size: 0.92rem;">
                    <p>Chưa có ga được chọn.</p>
                </div>
            </div>
            <div style="text-align:center; margin-top:.6rem;">
                <button id="help-btn" class="btn btn-sm btn-outline-secondary">Hướng dẫn sử dụng</button>
            </div>
        </div>
                <button id="toggle-sidebar-btn" title="Ẩn/hiện Sidebar">◀</button>
                <div id="map"></div>

                <!-- Help modal -->
                <div class="modal fade" id="helpModal" tabindex="-1" aria-hidden="true">
                    <div class="modal-dialog modal-lg modal-dialog-centered">
                        <div class="modal-content">
                            <div class="modal-header">
                                <h5 class="modal-title">Hướng dẫn sử dụng bản đồ TfL</h5>
                                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                            </div>
                            <div class="modal-body">
                                <ul>
                                    <li>Sử dụng thanh trượt <strong>Chọn năm phân tích</strong> để thay đổi dữ liệu hiển thị.</li>
                                    <li>Nhập tên ga vào ô <strong>Tìm theo tên ga...</strong> và nhấn Enter hoặc nút <strong>Tìm</strong> để phóng tới ga tương ứng.</li>
                                    <li>Nhấn các chip trong <strong>Bộ lọc cụm ga</strong> để bật/tắt các nhóm ga; dùng <strong>Lọc theo borough</strong> để thu hẹp theo khu vực.</li>
                                    <li>Di chuột lên marker để xem tooltip nhanh; nhấp vào marker để mở popup chi tiết, chọn <strong>Tới ga</strong> để phóng tới vị trí, hoặc <strong>Nhận xét</strong> để xem phân tích chi tiết bên thanh bên.</li>
                                    <li>Bảng <strong>TOP GA ĐÔNG KHÁCH NHẤT</strong> và biểu đồ hiển thị chỉ những ga nằm trong khung bản đồ hiện tại.</li>
                                    <li>Sử dụng nút <strong>Đặt lại bộ lọc</strong> để phục hồi trạng thái ban đầu.</li>
                                </ul>
                            </div>
                            <div class="modal-footer">
                                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Đóng</button>
                            </div>
                        </div>
                    </div>
                </div>
    </div>
    <script src="https://cdn.jsdelivr.net/npm/leaflet@1.9.3/dist/leaflet.js"></script>
    <script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js"></script>
    <script src="https://unpkg.com/leaflet.heat/dist/leaflet-heat.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.2.2/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        const stations = __STATION_DATA__;
        const categories = __CATEGORY_DATA__;
        const clusterColors = ["#38bdf8", "#6366f1", "#f97316", "#fb7185", "#22c55e", "#7c3aed", "#f59e0b"];
        const categoryColors = Object.fromEntries(categories.map((category, index) => [category, clusterColors[index % clusterColors.length]]));
        const map = L.map('map', { zoomControl: false }).setView([__CENTER_LAT__, __CENTER_LON__], 10);
        // Basemap layers
        const baseLayers = {
            'default': L.tileLayer('https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png', { maxZoom: 19, attribution: '&copy; CartoDB & OpenStreetMap' }),
            'dark': L.tileLayer('https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png', { maxZoom: 19, attribution: '&copy; CartoDB & OpenStreetMap' }),
            'transport': L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 19, attribution: '&copy; OpenStreetMap contributors' }),
        };
        let currentBase = baseLayers['default'];
        currentBase.addTo(map);
        L.control.zoom({ position: 'topright' }).addTo(map);

        function getRadiusForPassengers(value, maxValue) {
            const minRadius = 6;
            const maxRadius = 28;
            if (!maxValue || maxValue <= 0) {
                return minRadius;
            }
            const ratio = Math.sqrt(value || 0) / Math.sqrt(maxValue);
            return Math.max(minRadius, Math.round(minRadius + ratio * (maxRadius - minRadius)));
        }

        function updateMarkerSizesForFilteredMarkers() {
            const filtered = markerList.filter(marker => clusterGroup.hasLayer(marker));
            const maxValue = filtered.reduce((max, marker) => Math.max(max, marker.stationData[`passengers_${selectedYear}`] || 0), 0);
            filtered.forEach(marker => {
                const value = marker.stationData[`passengers_${selectedYear}`] || 0;
                marker.setRadius(getRadiusForPassengers(value, maxValue));
            });
        }

        const clusterGroup = L.markerClusterGroup({
            chunkedLoading: true,
            disableClusteringAtZoom: 13,
            spiderfyOnMaxZoom: true,
            iconCreateFunction(cluster) {
                const markers = cluster.getAllChildMarkers();
                const count = markers.length;
                const totalPassengers = markers.reduce((s, m) => s + (m.stationData && (m.stationData[`passengers_${selectedYear}`] || 0)), 0);
                let sizeClass = 'cluster-small';
                if (count > 50) sizeClass = 'cluster-large';
                else if (count > 12) sizeClass = 'cluster-medium';
                const formatted = new Intl.NumberFormat('vi-VN').format(totalPassengers);
                const html = `<div class="count">${count}</div><div class="sub">${formatted}<div style="font-size:10px;">lượt</div></div>`;
                return L.divIcon({
                    html: html,
                    className: `custom-cluster-icon ${sizeClass}`,
                    iconSize: null,
                    iconAnchor: [0, 0],
                    popupAnchor: [0, -10],
                });
            },
        });

        const markerList = [];
        let selectedYear = 2021;

        function buildPopup(station, year) {
            const passengers = station[`passengers_${year}`] || 0;
            // Chuẩn hoá nhãn xu hướng sang tiếng Việt
            function normalizeTrend(t) {
                if (!t) return 'Chưa rõ';
                const s = String(t).trim().toLowerCase();
                if (s.includes('tăng mạnh') || s.includes('tang manh')) return 'Tăng mạnh';
                if (s.includes('tăng nhẹ') || s.includes('tang nhe')) return 'Tăng nhẹ';
                if (s.includes('giảm nhẹ') || s.includes('giam nhe')) return 'Giảm nhẹ';
                if (s.includes('giảm mạnh') || s.includes('giam manh')) return 'Giảm mạnh';
                if (s.includes('ổn định') || s.includes('on dinh') || s.includes('stable')) return 'Ổn định';
                if (s.includes('tăng') || s.includes('tang')) return 'Tăng';
                if (s.includes('giảm') || s.includes('giam')) return 'Giảm';
                return String(t);
            }
            const trend = normalizeTrend(station.trend_category);
            return `
                <div class="popup-card">
                    <h5 class="popup-title">${station.station}</h5>
                    <div class="popup-subtitle"><span class="popup-badge">${station.cluster_name}</span><span>${station.borough}</span></div>
                    <div class="popup-field"><span class="popup-field-label">Tuyến</span><span>${station.num_lines}</span></div>
                    <div class="popup-field"><span class="popup-field-label">Hành khách ${year}</span><span>${passengers.toLocaleString('vi-VN')}</span></div>
                    <div class="popup-field"><span class="popup-field-label">Xu hướng</span><span class="popup-trend">${trend}</span></div>
                        <div class="popup-sparkline" style="margin-top:.5rem">${renderSparkline(station)}</div>
                    <div style="margin-top:0.5rem; display:flex; gap:.5rem;">
                        <button class="btn btn-sm btn-primary zoom-to-btn" data-lat="${station.lat}" data-lon="${station.lon}" data-name="${station.station}">Tới ga</button>
                        <button class="btn btn-sm btn-outline-secondary comment-btn" data-name="${station.station}">Nhận xét</button>
                    </div>
                </div>
            `;
        }

        function createMarkers() {
            const maxVisiblePassengers = Math.max(...stations.map(station => station[`passengers_${selectedYear}`] || 0));
            stations.forEach(station => {
                const color = categoryColors[station.cluster_name] || '#0d6efd';
                const passengers = station[`passengers_${selectedYear}`] || 0;
                const marker = L.circleMarker([station.lat, station.lon], {
                    radius: getRadiusForPassengers(passengers, maxVisiblePassengers),
                    color: color,
                    fillColor: color,
                    fillOpacity: 0.85,
                    weight: 1.5,
                    opacity: 1,
                });
                marker.stationData = station;
                marker.bindTooltip(`<div style="font-weight:700">${station.station}</div><div style="font-size:0.85rem;">${station.cluster_name} • ${passengers.toLocaleString('vi-VN')} lượt</div><div style="font-size:0.78rem; color:#334155;">Xu hướng: ${normalizeTrendGlobal(station.trend_category)}</div>`, { direction: 'top', offset: [0, -12], opacity: 0.95, className: 'my-tooltip' });
                marker.bindPopup(buildPopup(station, selectedYear), { maxWidth: 340, className: 'station-popup' });
                marker.on('mouseover', () => marker.openTooltip());
                marker.on('mouseout', () => marker.closeTooltip());
                marker.on('click', () => { marker.openPopup(); renderStationComment(station); });
                markerList.push(marker);
            });
            clusterGroup.addLayers(markerList);
            map.addLayer(clusterGroup);
        }

        function renderCategoryFilters() {
            const container = document.getElementById('category-filters');
            container.innerHTML = categories.map(category => `
                <button type=\"button\" class=\"filter-chip active\" data-category=\"${category}\" style=\"--chip-color:${categoryColors[category]};\">${category}</button>
            `).join('');
        }

        function renderStationComment(station) {
            const passengers2017 = station.passengers_2017 || 0;
            const passengers2020 = station.passengers_2020 || 0;
            const passengers2021 = station.passengers_2021 || 0;
            const recoveryPercent = passengers2020 ? Math.round(((passengers2021 - passengers2020) / passengers2020) * 100) : null;
            const compare2017 = passengers2017 ? Math.round(((passengers2021 - passengers2017) / passengers2017) * 100) : null;
            const recoveryText = recoveryPercent !== null ? `${recoveryPercent >= 0 ? 'tăng' : 'giảm'} ${Math.abs(recoveryPercent)}%` : 'không có dữ liệu năm 2020';
            const compareText = compare2017 !== null ? `${compare2017 >= 0 ? 'tăng' : 'giảm'} ${Math.abs(compare2017)}%` : 'không có dữ liệu năm 2017';
            // Chuẩn hoá nhãn xu hướng sang tiếng Việt (nhỏ)
            const trendRemark = normalizeTrendGlobal(station.trend_category);
            document.getElementById('station-comment').innerHTML = `
                <div class="station-comment-card">
                    <div class="station-comment-summary">
                        <p><strong>Ga ${station.station}</strong> — thuộc nhóm <strong>${station.cluster_name}</strong>.</p>
                        <p>Năm ${selectedYear} ghi nhận <strong>${(station[`passengers_${selectedYear}`] || 0).toLocaleString('vi-VN')}</strong> lượt khách.</p>
                        <p>Xu hướng 5 năm gần nhất: <strong>${trendRemark}</strong>.</p>
                        <p>So với năm 2020: ${recoveryText}. So với năm 2017: ${compareText}.</p>
                        <p><em>Ghi chú:</em> Kích thước marker tỉ lệ với số hành khách so với các ga đang hiển thị theo bộ lọc.</p>
                    </div>
                    <div class="station-comment-sparkline">${renderSparkline(station)}</div>
                </div>
            `;
        }

        function renderLegend() {
            const legend = document.getElementById('cluster-legend');
            if (!legend) return;
            legend.innerHTML = categories.map(category => `
                <div class=\"legend-item\"><span class=\"legend-swatch\" style=\"background:${categoryColors[category]}\"></span>${category}</div>
            `).join('');
        }

        // Populate borough filter based on station data
        function renderBoroughOptions() {
            const boroughs = Array.from(new Set(stations.map(s => s.borough || '').filter(b => b && b !== 'Unknown'))).sort();
            const sel = document.getElementById('borough-filter');
            if (!sel) return;
            sel.innerHTML = '<option value="">-- Lọc theo borough --</option>' + boroughs.map(b => `<option value="${b}">${b}</option>`).join('');
        }

        function locateStationByName(name) {
            if (!name) return;
            const lower = name.trim().toLowerCase();
            const found = markerList.find(m => m.stationData.station && m.stationData.station.toLowerCase().includes(lower));
            if (found) {
                map.setView(found.getLatLng(), Math.max(map.getZoom(), 14), { animate: true });
                found.openPopup();
                renderStationComment(found.stationData);
            } else {
                alert('Không tìm thấy ga phù hợp.');
            }
        }

        // Normalise string: remove diacritics + lowercase for fuzzy match
        function normaliseString(s) {
            if (!s) return '';
            try { return s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase(); } catch (e) { return s.toLowerCase(); }
        }

        // Normalize trend to Vietnamese (global)
        function normalizeTrendGlobal(t) {
            if (!t) return 'Chưa rõ';
            const s = String(t).trim().toLowerCase();
            if (s.includes('tăng mạnh') || s.includes('tang manh')) return 'Tăng mạnh';
            if (s.includes('tăng nhẹ') || s.includes('tang nhe')) return 'Tăng nhẹ';
            if (s.includes('giảm nhẹ') || s.includes('giam nhe')) return 'Giảm nhẹ';
            if (s.includes('giảm mạnh') || s.includes('giam manh')) return 'Giảm mạnh';
            if (s.includes('ổn định') || s.includes('on dinh') || s.includes('stable')) return 'Ổn định';
            if (s.includes('tăng') || s.includes('tang')) return 'Tăng';
            if (s.includes('giảm') || s.includes('giam')) return 'Giảm';
            return String(t);
        }

        // Render small sparkline SVG for station 2017-2021
        function renderSparkline(station) {
            const years = [2017,2018,2019,2020,2021];
            const vals = years.map(y => station[`passengers_${y}`] || 0);
            const max = Math.max(...vals, 1);
            const min = Math.min(...vals);
            const w = 140, h = 36, pad = 4;
            const step = (w - pad*2) / (vals.length - 1);
            const points = vals.map((v,i) => {
                const x = pad + i * step;
                const y = h - pad - ((v - min) / (max - min || 1)) * (h - pad*2);
                return `${x},${y}`;
            }).join(' ');
            const area = `${points} ${w-pad},${h-pad} ${pad},${h-pad}`;
            return `<div style="width:${w}px"><svg width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" xmlns="http://www.w3.org/2000/svg"><polyline fill="none" stroke="#0d6efd" stroke-width="2" points="${points}" stroke-linecap="round" stroke-linejoin="round"/><polygon points="${area}" fill="rgba(13,110,253,0.08)" stroke="none"/></svg></div>`;
        }

        // Build a heat layer from an array of station objects
        let heatLayer = null;
        function buildHeatLayer(stationArray) {
            const points = stationArray.map(s => {
                const val = s[`passengers_${selectedYear}`] || 0;
                return [s.lat, s.lon, Math.max(0.1, val)];
            });
            const maxVal = Math.max(...points.map(p => p[2]), 1);
            // Sử dụng căn bậc hai kết hợp offset tối thiểu để làm nổi bật rõ nét các ga nhỏ và trung bình, triệt tiêu độ mờ nhạt
            const norm = points.map(p => {
                const rawIntensity = p[2] / maxVal;
                const scaledIntensity = 0.3 + 0.7 * Math.sqrt(rawIntensity);
                return [p[0], p[1], Math.max(0.1, scaledIntensity)];
            });
            return L.heatLayer(norm, {
                radius: 28,
                blur: 12,
                maxZoom: 17,
                minOpacity: 0.45,
                gradient: {
                    0.3: '#3b82f6',  // Xanh dương (lưu lượng thấp, hiển thị cực rõ nét)
                    0.5: '#10b981',  // Xanh lá (lưu lượng trung bình)
                    0.7: '#eab308',  // Vàng (lưu lượng khá)
                    0.85: '#f97316', // Cam (lưu lượng cao)
                    1.0: '#ef4444'   // Đỏ (lưu lượng cực lớn - siêu trung tâm)
                }
            });
        }

        // Debounce util
        function debounce(fn, wait) {
            let t = null;
            return function(...args) { clearTimeout(t); t = setTimeout(() => fn.apply(this, args), wait); };
        }

        function zoomToStation(lat, lon, name) {
            if (!lat || !lon) return;
            map.setView([lat, lon], 15, { animate: true });
            const found = markerList.find(m => Math.abs(m.getLatLng().lat - lat) < 1e-6 && Math.abs(m.getLatLng().lng - lon) < 1e-6);
            if (found) {
                found.openPopup();
                renderStationComment(found.stationData);
                // temporary highlight
                try {
                    const prevStyle = { color: found.options.color, weight: found.options.weight };
                    found.setStyle({ color: '#000000', weight: 3 });
                    setTimeout(() => { try { found.setStyle(prevStyle); } catch (e) {} }, 1800);
                } catch (e) {}
            }
        }

        // Delegate popup button clicks
        document.addEventListener('click', (e) => {
            const t = e.target;
            if (t && t.classList && t.classList.contains('zoom-to-btn')) {
                const lat = parseFloat(t.dataset.lat);
                const lon = parseFloat(t.dataset.lon);
                zoomToStation(lat, lon, t.dataset.name);
            }
            if (t && t.classList && t.classList.contains('comment-btn')) {
                const name = t.dataset.name;
                if (name) locateStationByName(name);
            }
        });

        function updateInsights(selectedCategories, visibleCount, busiestStation, clusterCounts, avgPassengers) {
            const activeCount = selectedCategories.size;
            const chosen = activeCount === categories.length ? 'Tất cả nhóm' : `${activeCount} nhóm`;
            const busiestText = busiestStation ? `${busiestStation.stationData.station} (${new Intl.NumberFormat('vi-VN').format(busiestStation.stationData[`passengers_${selectedYear}`] || 0)} khách)` : 'Chưa có ga nổi bật';
            const breakdownEntries = Object.entries(clusterCounts).sort((a, b) => b[1] - a[1]);
            document.getElementById('insight-summary').innerHTML = `
                <p><strong>${chosen}</strong> đang được hiển thị theo bộ lọc.</p>
                <p><strong>Tổng số ga:</strong> ${visibleCount}</p>
                <p><strong>Ga đông khách nhất:</strong> ${busiestText}</p>
            `;
            // Hiển thị rõ từng cụm với swatch màu và số ga
            const breakdownItems = breakdownEntries.length ? breakdownEntries.map(([category, count]) => `
                <div class="trend-overview-item"><div class="label"><span class="swatch" style="background:${categoryColors[category] || '#cbd5e1'}"></span><span>${category}</span></div><div><strong>${count} ga</strong></div></div>
            `).join('') : '';
            const avgItem = `<div class="trend-overview-item"><div class="label">Hành khách trung bình (các ga hiển thị)</div><div><strong>${avgPassengers ? new Intl.NumberFormat('vi-VN').format(avgPassengers) + ' lượt' : '0 lượt'}</strong></div></div>`;
            const noteItem = `<div class="trend-overview-item"><div class="label">Ghi chú</div><div><strong>Kích thước marker tỉ lệ với hành khách trong nhóm được chọn</strong></div></div>`;
            document.getElementById('trend-overview').innerHTML = breakdownItems || '<div>Không có dữ liệu theo khu.</div>';
            document.getElementById('trend-overview').innerHTML += avgItem + noteItem;
        }

        function updateAnalytics() {
            const visible = markerList.filter(marker => clusterGroup.hasLayer(marker) && map.getBounds().contains(marker.getLatLng()));
            const filtered = markerList.filter(marker => clusterGroup.hasLayer(marker));
            const visibleCount = visible.length;
            const totalPassengers = visible.reduce((sum, marker) => sum + (marker.stationData[`passengers_${selectedYear}`] || 0), 0);
            document.getElementById('summary-count').textContent = visibleCount;
            document.getElementById('summary-passengers').textContent = visibleCount ? new Intl.NumberFormat('vi-VN').format(Math.round(totalPassengers / visibleCount)) : '0';
            updateMarkerSizesForFilteredMarkers();

            const counts = visible.reduce((acc, marker) => {
                const key = marker.stationData.cluster_name || 'Không xác định';
                acc[key] = (acc[key] || 0) + 1;
                return acc;
            }, {});

            const chart = document.getElementById('analyticsChart');
            chart.innerHTML = Object.entries(counts).sort((a, b) => b[1] - a[1]).map(([category, count]) => {
                const width = visibleCount ? Math.round((count / visibleCount) * 100) : 0;
                const barColor = categoryColors[category] || '#0d6efd';
                return `<div class=\"chart-bar-row\"><span>${category}</span><div class=\"chart-bar\"><div style=\"width:${width}%; background:${barColor};\"></div></div><span>${count}</span></div>`;
            }).join('') || '<div class=\"text-muted\"><em>Không có dữ liệu hiển thị</em></div>';

            const topStations = visible.slice().sort((a, b) => (b.stationData[`passengers_${selectedYear}`] || 0) - (a.stationData[`passengers_${selectedYear}`] || 0)).slice(0, 4);
            const maxValue = topStations.length ? topStations[0].stationData[`passengers_${selectedYear}`] || 1 : 1;
            document.getElementById('top-stations').innerHTML = topStations.map(marker => {
                const value = marker.stationData[`passengers_${selectedYear}`] || 0;
                const width = maxValue ? Math.round((value / maxValue) * 100) : 0;
                return `
                    <div class="station-bar-row">
                        <div class="station-label"><span>${marker.stationData.station}</span><span>${new Intl.NumberFormat('vi-VN').format(value)}</span></div>
                        <div class="mini-bar"><div class="mini-bar-content" style="width:${width}%; background:${categoryColors[marker.stationData.cluster_name] || '#0d6efd'}"></div></div>
                    </div>
                `;
            }).join('') || '<div class="text-muted"><em>Không có trạm để hiển thị</em></div>';

            const selectedCategories = new Set(Array.from(document.querySelectorAll('.filter-chip.active')).map(el => el.dataset.category));
            const busiestStation = topStations[0] || null;
            const avgPassengers = visibleCount ? Math.round(totalPassengers / visibleCount) : 0;
            updateInsights(selectedCategories, visibleCount, busiestStation, counts, avgPassengers);
        }

        function applyFilters() {
            const selectedCategories = new Set(Array.from(document.querySelectorAll('.filter-chip.active')).map(el => el.dataset.category));
            selectedYear = parseInt(document.getElementById('time-slider').value, 10);
            document.getElementById('time-slider-value').textContent = selectedYear;
            const borough = document.getElementById('borough-filter') ? document.getElementById('borough-filter').value : '';
            const searchInputEl = document.getElementById('station-search');
            const searchTermRaw = searchInputEl ? searchInputEl.value.trim() : '';
            const searchTerm = normaliseString(searchTermRaw);
            const searchScope = document.getElementById('search-scope') ? document.getElementById('search-scope').value : 'name';

            markerList.forEach(marker => {
                const cat = marker.stationData.cluster_name;
                const boroughMatch = !borough || (marker.stationData.borough && marker.stationData.borough === borough);
                let searchMatch = true;
                if (searchTerm) {
                    if (searchScope === 'name') searchMatch = normaliseString(marker.stationData.station).includes(searchTerm);
                    else if (searchScope === 'line') searchMatch = String(marker.stationData.num_lines || '').includes(searchTermRaw) || (marker.stationData.lines && normaliseString(marker.stationData.lines).includes(searchTerm));
                    else if (searchScope === 'borough') searchMatch = normaliseString(marker.stationData.borough).includes(searchTerm);
                }
                const matchesCategory = selectedCategories.has(cat);
                const passengerValue = marker.stationData[`passengers_${selectedYear}`] || 0;
                const show = matchesCategory && boroughMatch && searchMatch;

                if (show) {
                    if (!clusterGroup.hasLayer(marker)) {
                        clusterGroup.addLayer(marker);
                    }
                    marker.getPopup().setContent(buildPopup(marker.stationData, selectedYear));
                    const visibleMarkers = markerList.filter(m => {
                        const cat2 = m.stationData.cluster_name;
                        const boroughMatch2 = !borough || (m.stationData.borough && m.stationData.borough === borough);
                        const searchMatch2 = !searchTerm || (m.stationData.station && m.stationData.station.toLowerCase().includes(searchTerm));
                        return selectedCategories.has(cat2) && boroughMatch2 && searchMatch2;
                    });
                    const maxVisiblePassengers = visibleMarkers.reduce((max, m) => Math.max(max, m.stationData[`passengers_${selectedYear}`] || 0), 0);
                    marker.setRadius(getRadiusForPassengers(passengerValue, maxVisiblePassengers));
                    marker.unbindTooltip();
                    // highlight matched term in tooltip
                    let tooltipName = marker.stationData.station;
                    if (searchTerm && searchScope === 'name') {
                        const needle = searchTermRaw.toLowerCase();
                        const hay = String(marker.stationData.station || '').toLowerCase();
                        const idx = hay.indexOf(needle);
                        if (idx >= 0) {
                            const orig = String(marker.stationData.station || '');
                            tooltipName = orig.slice(0, idx) + `<mark style="background: #fffbcc; padding:0 .15rem; border-radius:3px; color:#000">` + orig.slice(idx, idx + needle.length) + `</mark>` + orig.slice(idx + needle.length);
                        }
                    }
                    marker.bindTooltip(`<div style="font-weight:700">${tooltipName}</div><div style="font-size:0.85rem;">${marker.stationData.cluster_name} • ${passengerValue.toLocaleString('vi-VN')} lượt</div><div style="font-size:0.78rem; color:#334155;">Xu hướng: ${normalizeTrendGlobal(marker.stationData.trend_category)}</div>`, { direction: 'top', offset: [0, -12], opacity: 0.95, className: 'my-tooltip' });
                } else {
                    if (clusterGroup.hasLayer(marker)) {
                        clusterGroup.removeLayer(marker);
                    }
                }
            });
            updateAnalytics();

            // Update heatmap if active
            const heatOn = document.getElementById('heatmap-toggle') && document.getElementById('heatmap-toggle').checked;
            if (heatOn) {
                try { if (heatLayer) map.removeLayer(heatLayer); } catch (e) {}
                const visibleStations = markerList.filter(m => clusterGroup.hasLayer(m)).map(m => m.stationData);
                heatLayer = buildHeatLayer(visibleStations);
                map.addLayer(heatLayer);
            } else {
                try { if (heatLayer) { map.removeLayer(heatLayer); heatLayer = null; } } catch (e) {}
                if (!map.hasLayer(clusterGroup)) map.addLayer(clusterGroup);
            }

            // Refresh cluster icons (to update passenger sums)
            if (clusterGroup && typeof clusterGroup.refreshClusters === 'function') try { clusterGroup.refreshClusters(); } catch (e) {}
        }

        document.addEventListener('DOMContentLoaded', () => {
            renderCategoryFilters();
            renderLegend();
            renderBoroughOptions();
            createMarkers();
            document.querySelectorAll('.filter-chip').forEach(el => el.addEventListener('click', () => {
                el.classList.toggle('active');
                applyFilters();
            }));
            document.getElementById('time-slider').addEventListener('input', applyFilters);
            // search and borough listeners with debounce
            const searchInput = document.getElementById('station-search');
            if (searchInput) {
                const debounced = debounce(applyFilters, 300);
                searchInput.addEventListener('input', debounced);
                searchInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') { locateStationByName(searchInput.value); } });
            }
            const clearBtn = document.getElementById('clear-search');
            if (clearBtn) clearBtn.addEventListener('click', () => { if (searchInput) { searchInput.value = ''; searchInput.focus(); applyFilters(); } });
            const locateBtn = document.getElementById('locate-btn');
            if (locateBtn) locateBtn.addEventListener('click', () => locateStationByName(document.getElementById('station-search').value));
            const boroughSel = document.getElementById('borough-filter');
            if (boroughSel) boroughSel.addEventListener('change', applyFilters);
            const scopeSel = document.getElementById('search-scope');
            if (scopeSel) scopeSel.addEventListener('change', applyFilters);
            document.getElementById('reset-filters').addEventListener('click', () => {
                document.querySelectorAll('.filter-chip').forEach(el => el.classList.add('active'));
                document.getElementById('time-slider').value = '2021';
                if (document.getElementById('borough-filter')) document.getElementById('borough-filter').value = '';
                if (document.getElementById('station-search')) document.getElementById('station-search').value = '';
                applyFilters();
            });

            // Export CSV listener
            const exportBtn = document.getElementById('export-csv-btn');
            if (exportBtn) {
                exportBtn.addEventListener('click', () => {
                    const visible = markerList.filter(marker => clusterGroup.hasLayer(marker));
                    if (visible.length === 0) {
                        alert('Không có dữ liệu để xuất.');
                        return;
                    }
                    
                    let csvContent = "\uFEFF"; // UTF-8 BOM
                    csvContent += "Tên ga,Cụm phân loại,Borough,Số tuyến,Danh sách tuyến,Hành khách 2017,Hành khách 2018,Hành khách 2019,Hành khách 2020,Hành khách 2021,Xu hướng\\n";
                    
                    visible.forEach(m => {
                        const s = m.stationData;
                        const row = [
                            `"${s.station.replace(/"/g, '""')}"`,
                            `"${s.cluster_name.replace(/"/g, '""')}"`,
                            `"${s.borough.replace(/"/g, '""')}"`,
                            s.num_lines,
                            `"${(s.lines || '').replace(/"/g, '""')}"`,
                            s.passengers_2017,
                            s.passengers_2018,
                            s.passengers_2019,
                            s.passengers_2020,
                            s.passengers_2021,
                            `"${s.trend_category.replace(/"/g, '""')}"`
                        ].join(",");
                        csvContent += row + "\\n";
                    });
                    
                    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
                    const url = URL.createObjectURL(blob);
                    const link = document.createElement("a");
                    link.setAttribute("href", url);
                    link.setAttribute("download", `tfl_stations_filtered_${selectedYear}.csv`);
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                });
            }
            const helpBtn = document.getElementById('help-btn');
            if (helpBtn) helpBtn.addEventListener('click', () => {
                const modalEl = document.getElementById('helpModal');
                const modal = new bootstrap.Modal(modalEl);
                modal.show();
            });
            // Sidebar accordion behaviour: click title to collapse/expand card
            document.querySelectorAll('#sidebar .sidebar-card .sidebar-section-title').forEach(title => {
                title.addEventListener('click', (e) => {
                    const card = title.closest('.sidebar-card');
                    if (card) card.classList.toggle('collapsed-card');
                });
            });

            // Basemap selector
            const basemapSel = document.getElementById('basemap-select');
            if (basemapSel) {
                basemapSel.addEventListener('change', (e) => {
                    const v = e.target.value || 'default';
                    if (currentBase) map.removeLayer(currentBase);
                    currentBase = baseLayers[v] || baseLayers['default'];
                    currentBase.addTo(map);

                    const sidebar = document.getElementById('sidebar');
                    if (sidebar) {
                        if (v === 'dark') {
                            sidebar.classList.add('dark-theme');
                        } else {
                            sidebar.classList.remove('dark-theme');
                        }
                    }
                });
            }
            // Heatmap toggle
            const heatToggle = document.getElementById('heatmap-toggle');
            if (heatToggle) {
                heatToggle.addEventListener('change', (e) => {
                    const on = e.target.checked;
                    if (on) {
                        try { if (clusterGroup) map.removeLayer(clusterGroup); } catch (err) {}
                        const visibleStations = markerList.filter(m => clusterGroup.hasLayer(m)).map(m => m.stationData);
                        heatLayer = buildHeatLayer(visibleStations);
                        map.addLayer(heatLayer);
                    } else {
                        try { if (heatLayer) { map.removeLayer(heatLayer); heatLayer = null; } } catch (err) {}
                        if (!map.hasLayer(clusterGroup)) map.addLayer(clusterGroup);
                    }
                });
            }
            document.getElementById('toggle-sidebar-btn').addEventListener('click', () => {
                const sidebar = document.getElementById('sidebar');
                sidebar.classList.toggle('collapsed');
                document.getElementById('toggle-sidebar-btn').textContent = sidebar.classList.contains('collapsed') ? '▶' : '◀';
                setTimeout(() => { map.invalidateSize(); }, 300);
            });
            map.on('moveend zoomend', updateAnalytics);
            applyFilters();
        });
    </script>
</body>
</html>"""

    html = html.replace("__STATION_DATA__", station_data)
    html = html.replace("__CATEGORY_DATA__", category_data)
    html = html.replace("__CENTER_LAT__", str(center_lat))
    html = html.replace("__CENTER_LON__", str(center_lon))

    output_dir.mkdir(parents=True, exist_ok=True)
    map_path.write_text(html, encoding='utf-8')
    print(f"Da luu ban do tuong tac: {map_path}")

    # Copy to root as FINAL_MAP.html for convenience
    root_map_path = output_dir.parent / "FINAL_MAP.html"
    try:
        import shutil
        shutil.copy2(map_path, root_map_path)
        print(f"Da copy ban do vao root: {root_map_path}")
    except Exception as e:
        print(f"Loi copy FINAL_MAP.html: {e}")


# ============================================================================
# 7. MAIN - CHAY TOAN BO PIPELINE
# ============================================================================

def run_pipeline() -> pd.DataFrame:
    project_dir = get_project_folder()
    output_dir = project_dir / "outputs"

    kml_path = find_file(project_dir, KML_CANDIDATES)
    tfl_path = find_file(project_dir, TFL_CANDIDATES)
    stops_path = find_file(project_dir, STOPS_CANDIDATES)

    print_title("BAT DAU PIPELINE TfL")
    print(f"Thu muc du lieu: {project_dir}")
    print(f"File KML: {kml_path.name}")
    print(f"File TfL CSV: {tfl_path.name}")
    print(f"File Stops CSV: {stops_path.name}")

    print_title("BUOC 1 - EXTRACT")
    kml_df = load_kml_stations(kml_path)
    tfl_df = load_tfl_csv(tfl_path)
    stops_df = load_stops_csv(stops_path)

    print_title("BUOC 2 - TRANSFORM")
    merged_df, merge_report = merge_sources(kml_df, tfl_df, stops_df)
    final_df = clean_and_engineer(merged_df)

    # Ghi log cac ga khong khop hoac bi loai bo
    log_unmatched_stations(kml_df, final_df, output_dir)

    print_title("BUOC 3 - CLUSTERING VA ANALYSIS")
    final_df = run_kmeans_clustering(final_df, n_clusters=N_CLUSTERS)
    print_analysis_report(final_df, merge_report)

    print_title("BUOC 4 - LOAD")
    save_outputs(final_df, output_dir)

    print_title("BUOC 5 - VISUALIZATION")
    create_folium_map(final_df, output_dir)

    print_title("BUOC 6 - GENERATE PDF REPORT")
    try:
        from generate_pdf_report import generate_pdf_report as make_pdf
        pdf_path = output_dir / "TfL_Project_Report.pdf"
        db_path = output_dir / "london_tfl.db"
        make_pdf(db_path, pdf_path)
        
        # Copy to root as Bao_Cao_TfL.pdf for convenience
        root_pdf_path = output_dir.parent / "Bao_Cao_TfL.pdf"
        try:
            import shutil
            shutil.copy2(pdf_path, root_pdf_path)
            print(f"Da copy bao cao PDF vao root: {root_pdf_path}")
        except Exception as e:
            print(f"Loi copy Bao_Cao_TfL.pdf: {e}")
    except Exception as e:
        print(f"Loi khi sinh bao cao PDF: {e}")

    print_title("HOAN THANH")
    print(f"So ga cuoi cung: {len(final_df)}")
    print(f"So cot cuoi cung: {len(final_df.columns)}")
    print(f"Thu muc ket qua: {output_dir}")

    return final_df


if __name__ == "__main__":
    run_pipeline()

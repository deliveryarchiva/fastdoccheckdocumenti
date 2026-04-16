"""
Data loader for Archiva File DB.

Archiva sheet column indices (0-based):
  7  -> file_name
  94 -> PIVA
  97 -> RAGIONE_SOCIALE
  115 -> r_object_id  (original Postel ID, used as join key)

Postel sheet column indices (0-based):
  0  -> r_object_id
  2  -> object_name  (filename)
  32 -> pt_ragione_sociale
  88 -> pt_piva

Filename date pattern: PIVA_YYYY_M_D_rest...
"""

import re
import logging
from datetime import date

logger = logging.getLogger(__name__)

_FLOAT_RE = re.compile(r"^(\d+)\.0$")


def _clean_str(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    # Remove trailing .0 from numeric strings (pandas float artifact)
    m = _FLOAT_RE.match(s)
    if m:
        return m.group(1)
    if s.lower() in ("nan", "none", "nat", "nulldate"):
        return ""
    return s


def parse_date_from_filename(filename: str):
    """Extract date from filename pattern: PIVA_YYYY_M_D_..."""
    if not filename:
        return None
    try:
        parts = filename.split("_")
        if len(parts) >= 4:
            year  = int(parts[1])
            month = int(parts[2])
            day   = int(parts[3])
            if 2000 <= year <= 2099 and 1 <= month <= 12 and 1 <= day <= 31:
                return date(year, month, day)
    except (ValueError, IndexError):
        pass
    return None


def load_data(excel_path: str) -> list:
    """
    Load both sheets, merge on r_object_id, return a unified list of records.
    Uses a pickle cache alongside the Excel file to speed up repeated startups.
    On first load from Excel: ~60-90s. Subsequent loads from cache: ~2-5s.
    """
    import os
    import pickle

    cache_path = excel_path + ".cache.pkl"

    # Use cache if it exists and is newer than the Excel file
    if os.path.exists(cache_path):
        excel_mtime = os.path.getmtime(excel_path)
        cache_mtime = os.path.getmtime(cache_path)
        if cache_mtime >= excel_mtime:
            try:
                logger.info(f"Loading from cache {cache_path} …")
                import time; t0 = time.time()
                with open(cache_path, "rb") as f:
                    records = pickle.load(f)
                logger.info(f"Cache loaded in {time.time()-t0:.1f}s — {len(records):,} records")
                return records
            except Exception as e:
                logger.warning(f"Cache load failed ({e}), reloading from Excel")

    # Load from Excel
    try:
        import pandas as pd
        records = _load_with_pandas(excel_path, pd)
    except Exception as e:
        logger.warning(f"pandas load failed ({e}), falling back to openpyxl")
        records = _load_with_openpyxl(excel_path)

    # Save cache
    try:
        with open(cache_path, "wb") as f:
            pickle.dump(records, f, protocol=pickle.HIGHEST_PROTOCOL)
        logger.info(f"Cache saved to {cache_path}")
    except Exception as e:
        logger.warning(f"Could not save cache: {e}")

    return records


def _load_with_pandas(excel_path: str, pd) -> list:
    logger.info("Loading Archiva sheet with pandas...")
    df_a = pd.read_excel(
        excel_path,
        sheet_name="Estrazione Archiva",
        usecols=[7, 94, 97, 115],
        header=0,
        dtype=object,
        engine="openpyxl",
    )
    df_a.columns = ["nome_file", "piva", "ragione_sociale", "r_object_id"]

    logger.info("Loading Postel sheet with pandas...")
    df_p = pd.read_excel(
        excel_path,
        sheet_name="Estrazione Postel",
        usecols=[0, 2, 32, 88],
        header=0,
        dtype=object,
        engine="openpyxl",
    )
    df_p.columns = ["r_object_id", "nome_file", "ragione_sociale", "piva"]

    logger.info("Building unified record set...")
    return _build_records(df_a.values.tolist(), df_p.values.tolist())


def _load_with_openpyxl(excel_path: str) -> list:
    from openpyxl import load_workbook

    A_COLS = [7, 94, 97, 115]
    P_COLS = [0, 2, 32, 88]

    def extract_cols(row, cols):
        return [row[c] if c < len(row) else None for c in cols]

    logger.info("Loading Archiva sheet with openpyxl...")
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    rows_a, rows_p = [], []

    ws_a = wb["Estrazione Archiva"]
    for i, row in enumerate(ws_a.iter_rows(values_only=True)):
        if i == 0:
            continue
        rows_a.append(extract_cols(row, A_COLS))

    ws_p = wb["Estrazione Postel"]
    for i, row in enumerate(ws_p.iter_rows(values_only=True)):
        if i == 0:
            continue
        rows_p.append(extract_cols(row, P_COLS))

    wb.close()
    return _build_records(rows_a, rows_p)


def _build_records(rows_archiva: list, rows_postel: list) -> list:
    """
    rows_archiva: list of [nome_file, piva, ragione_sociale, r_object_id]
    rows_postel:  list of [r_object_id, nome_file, ragione_sociale, piva]
    """
    # Build Archiva index: r_object_id -> dict
    archiva_map: dict[str, dict] = {}
    for row in rows_archiva:
        nf, pv, rs, rid = row
        rid = _clean_str(rid)
        if not rid:
            continue
        nf  = _clean_str(nf)
        pv  = _clean_str(pv)
        rs  = _clean_str(rs)
        dt  = parse_date_from_filename(nf)
        archiva_map[rid] = {
            "r_object_id":    rid,
            "nome_file":      nf,
            "piva":           pv,
            "ragione_sociale": rs,
            "data_doc":       dt.isoformat() if dt else "",
            "in_archiva":     "SI",
            "in_postel":      "NO",
        }

    # Process Postel: update or add
    for row in rows_postel:
        rid, nf, rs, pv = row
        rid = _clean_str(rid)
        if not rid:
            continue
        nf = _clean_str(nf)
        pv = _clean_str(pv)
        rs = _clean_str(rs)
        dt = parse_date_from_filename(nf)

        if rid in archiva_map:
            archiva_map[rid]["in_postel"] = "SI"
            # Fill missing data from Postel record
            if not archiva_map[rid]["piva"]:
                archiva_map[rid]["piva"] = pv
            if not archiva_map[rid]["ragione_sociale"]:
                archiva_map[rid]["ragione_sociale"] = rs
        else:
            archiva_map[rid] = {
                "r_object_id":    rid,
                "nome_file":      nf,
                "piva":           pv,
                "ragione_sociale": rs,
                "data_doc":       dt.isoformat() if dt else "",
                "in_archiva":     "NO",
                "in_postel":      "SI",
            }

    records = list(archiva_map.values())
    logger.info(
        f"Records: total={len(records)}, "
        f"archiva={sum(1 for r in records if r['in_archiva']=='SI')}, "
        f"postel={sum(1 for r in records if r['in_postel']=='SI')}, "
        f"entrambi={sum(1 for r in records if r['in_archiva']=='SI' and r['in_postel']=='SI')}"
    )
    return records


def search_records(
    records: list,
    ragione_sociale: str = None,
    piva: str = None,
    nome_file: str = None,
    file_match: str = "partial",
    date_from: str = None,
    date_to: str = None,
) -> list:
    rs_q  = ragione_sociale.strip().lower() if ragione_sociale else None
    pv_q  = piva.strip().lower() if piva else None
    nf_q  = nome_file.strip().lower() if nome_file else None
    dt_f  = date_from.strip() if date_from else None
    dt_t  = date_to.strip() if date_to else None

    results = []
    for r in records:
        if rs_q and rs_q not in r["ragione_sociale"].lower():
            continue
        if pv_q and pv_q not in r["piva"].lower():
            continue
        if nf_q:
            fn = r["nome_file"].lower()
            if file_match == "exact":
                if nf_q != fn:
                    continue
            else:
                if nf_q not in fn:
                    continue
        if dt_f or dt_t:
            dd = r["data_doc"]
            if not dd:
                continue
            if dt_f and dd < dt_f:
                continue
            if dt_t and dd > dt_t:
                continue
        results.append(r)
    return results


def compute_stats(records: list) -> dict:
    total     = len(records)
    archiva   = sum(1 for r in records if r["in_archiva"] == "SI")
    postel    = sum(1 for r in records if r["in_postel"] == "SI")
    entrambi  = sum(1 for r in records if r["in_archiva"] == "SI" and r["in_postel"] == "SI")
    solo_archiva = archiva - entrambi
    solo_postel  = postel  - entrambi
    return {
        "total":        total,
        "in_archiva":   archiva,
        "in_postel":    postel,
        "entrambi":     entrambi,
        "solo_archiva": solo_archiva,
        "solo_postel":  solo_postel,
    }

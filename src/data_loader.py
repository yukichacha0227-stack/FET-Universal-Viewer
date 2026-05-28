from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd


DEFAULT_COLUMNS = ["No.", "Temp", "Mag", "Isd", "Vsd", "Vbg"]
REQUIRED_COLUMNS = ("Isd", "Vsd", "Vbg")

COLUMN_ALIASES = {
    "no": "No.",
    "no.": "No.",
    "number": "No.",
    "temp": "Temp",
    "temperature": "Temp",
    "温度": "Temp",
    "mag": "Mag",
    "magnetic field": "Mag",
    "field": "Mag",
    "磁場": "Mag",
    "isd": "Isd",
    "i_sd": "Isd",
    "ids": "Isd",
    "id": "Isd",
    "drain current": "Isd",
    "vsd": "Vsd",
    "v_sd": "Vsd",
    "vds": "Vsd",
    "vd": "Vsd",
    "drain voltage": "Vsd",
    "vbg": "Vbg",
    "v_bg": "Vbg",
    "vg": "Vbg",
    "gate voltage": "Vbg",
}


@dataclass(frozen=True)
class MergeResult:
    dataframe: pd.DataFrame
    source_frames: list[pd.DataFrame]
    processed_count: int


def read_measurement_file(file_path: str | Path) -> pd.DataFrame:
    """Read one FET measurement file and normalize common column names."""
    path = Path(file_path)
    errors: list[str] = []

    for has_header in (True, False):
        try:
            df = _read_table(path, has_header=has_header)
            if df.empty:
                continue
            df = _normalize_columns(df, has_header=has_header)
            df = _clean_measurement_values(df)
            _validate_required_columns(df, path)
            df = df.dropna(subset=list(REQUIRED_COLUMNS)).reset_index(drop=True)
            df["_SourceFile"] = path.name
            return df
        except Exception as exc:  # Collect both header strategies for a useful message.
            errors.append(str(exc))

    details = " / ".join(dict.fromkeys(errors))
    raise ValueError(f"Could not read {path.name}: {details}")


def merge_measurement_files(file_paths: Iterable[str | Path]) -> MergeResult:
    """Read and concatenate multiple FET measurement files."""
    source_frames: list[pd.DataFrame] = []

    for file_index, file_path in enumerate(file_paths):
        df = read_measurement_file(file_path)
        df["_Sort_ID"] = file_index * 1_000_000 + np.arange(len(df))
        source_frames.append(df)

    if not source_frames:
        raise ValueError("No files were selected.")

    merged = pd.concat(source_frames, ignore_index=True)
    merged["Vsd_R"] = merged["Vsd"].round(2)
    merged["Vbg_R"] = merged["Vbg"].round(2)

    return MergeResult(
        dataframe=merged,
        source_frames=source_frames,
        processed_count=len(source_frames),
    )


def write_merged_excel(result: MergeResult, output_path: str | Path) -> None:
    """Write merged data and side-by-side source tables to an Excel workbook."""
    output_path = Path(output_path)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        result.dataframe.drop(columns=["_Sort_ID"], errors="ignore").to_excel(
            writer,
            sheet_name="Merged_Data",
            index=False,
        )

        current_col = 0
        for df in result.source_frames:
            export_df = df.drop(columns=["_Sort_ID"], errors="ignore")
            export_df.to_excel(
                writer,
                sheet_name="Raw_Data",
                startcol=current_col,
                index=False,
            )
            current_col += len(export_df.columns) + 1


def _read_table(path: Path, has_header: bool) -> pd.DataFrame:
    header = 0 if has_header else None
    encoding_errors: list[str] = []

    for encoding in ("utf-8-sig", "cp932", "shift_jis"):
        for sep in (None, r"\s+"):
            try:
                df = pd.read_csv(
                    path,
                    sep=sep,
                    engine="python",
                    header=header,
                    encoding=encoding,
                )
                if len(df.columns) > 1:
                    return df
                encoding_errors.append(f"{encoding}: detected only one column")
            except UnicodeDecodeError as exc:
                encoding_errors.append(f"{encoding}: {exc}")
                break
            except pd.errors.ParserError as exc:
                encoding_errors.append(f"{encoding}: {exc}")

    joined = " / ".join(encoding_errors)
    raise ValueError(f"Unable to parse table ({joined})")


def _normalize_columns(df: pd.DataFrame, has_header: bool) -> pd.DataFrame:
    df = df.copy()

    if has_header:
        normalized = []
        for col in df.columns:
            raw = str(col).strip()
            key = raw.lower()
            normalized.append(COLUMN_ALIASES.get(key, raw))
        df.columns = normalized
    else:
        width = len(df.columns)
        fallback = DEFAULT_COLUMNS + [f"Extra_{i}" for i in range(width - len(DEFAULT_COLUMNS))]
        df.columns = fallback[:width]

    drop_cols = [col for col in df.columns if str(col).startswith("Unnamed:")]
    if drop_cols:
        df = df.drop(columns=drop_cols)

    return df


def _clean_measurement_values(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.replace({"#######": np.nan, "########": np.nan, "OVER": np.nan, "Overflow": np.nan})

    numeric_columns = [col for col in [*DEFAULT_COLUMNS, *REQUIRED_COLUMNS] if col in df.columns]
    for col in dict.fromkeys(numeric_columns):
        df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def _validate_required_columns(df: pd.DataFrame, path: Path) -> None:
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(f"{path.name} is missing required columns: {', '.join(missing)}")

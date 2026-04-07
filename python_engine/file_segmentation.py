import os
import pandas as pd
import numpy as np


def segment_by_column(file_path: str, column: str,
                       output_dir: str = "segments") -> dict:
    """
    Split an Excel file into multiple files based on unique values
    in a given column. One file per unique value.
    Example: segment by 'Region' -> North.xlsx, South.xlsx, East.xlsx
    """
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        return {"error": f"Could not read file: {str(e)}"}

    if column not in df.columns:
        return {"error": f"Column '{column}' not found. Available: {list(df.columns)}"}

    os.makedirs(output_dir, exist_ok=True)

    unique_values = df[column].dropna().unique()
    created_files = []

    for val in unique_values:
        segment_df = df[df[column] == val].reset_index(drop=True)
        safe_name = str(val).replace("/", "-").replace("\\", "-").replace(" ", "_")
        out_path = os.path.join(output_dir, f"{safe_name}.xlsx")
        segment_df.to_excel(out_path, index=False)
        created_files.append({
            "value": str(val),
            "rows": len(segment_df),
            "path": out_path,
        })

    return {
        "message": f"Segmented into {len(created_files)} files by '{column}'",
        "output_dir": output_dir,
        "segments": created_files,
    }


def segment_by_row_count(file_path: str, chunk_size: int = 1000,
                          output_dir: str = "segments") -> dict:
    """
    Split a large Excel file into smaller chunks of fixed row count.
    Example: 5000 rows with chunk_size=1000 -> 5 files of 1000 rows each
    """
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        return {"error": f"Could not read file: {str(e)}"}

    if chunk_size < 1:
        return {"error": "chunk_size must be at least 1"}

    os.makedirs(output_dir, exist_ok=True)

    total = len(df)
    chunks = [df.iloc[i: i + chunk_size] for i in range(0, total, chunk_size)]
    created_files = []

    base_name = os.path.splitext(os.path.basename(file_path))[0]

    for idx, chunk in enumerate(chunks, start=1):
        out_path = os.path.join(output_dir, f"{base_name}_part{idx}.xlsx")
        chunk.reset_index(drop=True).to_excel(out_path, index=False)
        created_files.append({
            "part": idx,
            "rows": len(chunk),
            "path": out_path,
        })

    return {
        "message": f"Split into {len(chunks)} parts of max {chunk_size} rows",
        "total_rows": total,
        "output_dir": output_dir,
        "segments": created_files,
    }


def segment_by_date_column(file_path: str, date_column: str,
                            freq: str = "M",
                            output_dir: str = "segments") -> dict:
    """
    Split Excel data by date periods.
    freq: 'D' = daily, 'W' = weekly, 'M' = monthly, 'Q' = quarterly, 'Y' = yearly
    """
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        return {"error": f"Could not read file: {str(e)}"}

    if date_column not in df.columns:
        return {"error": f"Column '{date_column}' not found. Available: {list(df.columns)}"}

    try:
        df[date_column] = pd.to_datetime(df[date_column])
    except Exception:
        return {"error": f"Column '{date_column}' could not be parsed as dates"}

    os.makedirs(output_dir, exist_ok=True)

    freq_map = {"D": "day", "W": "week", "M": "month", "Q": "quarter", "Y": "year"}
    df["_period"] = df[date_column].dt.to_period(freq)
    created_files = []

    for period, group in df.groupby("_period"):
        segment_df = group.drop(columns=["_period"]).reset_index(drop=True)
        safe_name = str(period).replace("/", "-")
        out_path = os.path.join(output_dir, f"{safe_name}.xlsx")
        segment_df.to_excel(out_path, index=False)
        created_files.append({
            "period": str(period),
            "rows": len(segment_df),
            "path": out_path,
        })

    label = freq_map.get(freq, freq)
    return {
        "message": f"Segmented into {len(created_files)} files by {label}",
        "output_dir": output_dir,
        "segments": created_files,
    }


def merge_excel_files(file_paths: list, output_path: str = "merged_output.xlsx") -> dict:
    """
    Merge multiple Excel files into one single file.
    All files must have compatible columns.
    """
    if not file_paths:
        return {"error": "No file paths provided"}

    dfs = []
    for fp in file_paths:
        try:
            dfs.append(pd.read_excel(fp))
        except Exception as e:
            return {"error": f"Could not read {fp}: {str(e)}"}

    try:
        merged = pd.concat(dfs, ignore_index=True)
        merged.to_excel(output_path, index=False)
        return {
            "message": f"Merged {len(file_paths)} files into {output_path}",
            "total_rows": len(merged),
            "output_path": output_path,
        }
    except Exception as e:
        return {"error": f"Merge failed: {str(e)}"}


def get_file_info(file_path: str) -> dict:
    """
    Quick summary of an Excel file — size, rows, cols, column names.
    Uses numpy for numeric summary.
    """
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        return {"error": str(e)}

    numeric_df = df.select_dtypes(include=[np.number])
    file_size_kb = round(os.path.getsize(file_path) / 1024, 2) if os.path.exists(file_path) else 0

    return {
        "file_path": file_path,
        "file_size_kb": file_size_kb,
        "total_rows": len(df),
        "total_columns": len(df.columns),
        "columns": list(df.columns),
        "numeric_columns": list(numeric_df.columns),
        "numeric_summary": {
            col: {
                "mean":   round(float(np.mean(df[col].dropna())), 2),
                "std":    round(float(np.std(df[col].dropna())), 2),
                "min":    round(float(np.min(df[col].dropna())), 2),
                "max":    round(float(np.max(df[col].dropna())), 2),
            }
            for col in numeric_df.columns
        },
    }

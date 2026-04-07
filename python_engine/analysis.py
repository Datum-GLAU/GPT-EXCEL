import numpy as np
import pandas as pd


def analyze_data(file_path: str) -> dict:
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        return {"error": f"Could not read file: {str(e)}"}

    numeric_df = df.select_dtypes(include=[np.number])

    summary = {
        "total_rows":      int(len(df)),
        "total_columns":   int(len(df.columns)),
        "columns":         list(df.columns),
        "data_types":      df.dtypes.astype(str).to_dict(),
        "null_values":     df.isnull().sum().to_dict(),
        "null_percentage": (df.isnull().mean() * 100).round(2).to_dict(),
        "duplicate_rows":  int(df.duplicated().sum()),
        "numeric_columns": list(numeric_df.columns),
        "text_columns":    list(df.select_dtypes(exclude=[np.number]).columns),
    }

    if not numeric_df.empty:
        np_stats = {}
        for col in numeric_df.columns:
            values = numeric_df[col].dropna().to_numpy()
            if len(values) == 0:
                continue
            np_stats[col] = {
                "mean":     round(float(np.mean(values)), 2),
                "median":   round(float(np.median(values)), 2),
                "std":      round(float(np.std(values)), 2),
                "variance": round(float(np.var(values)), 2),
                "min":      round(float(np.min(values)), 2),
                "max":      round(float(np.max(values)), 2),
                "range":    round(float(np.max(values) - np.min(values)), 2),
                "q25":      round(float(np.percentile(values, 25)), 2),
                "q75":      round(float(np.percentile(values, 75)), 2),
            }
        summary["numpy_stats"] = np_stats

    try:
        stats = df.describe(include="all", datetime_is_numeric=True).round(2).fillna("")
        summary["statistics"] = stats.to_dict()
    except Exception:
        summary["statistics"] = "Not applicable"

    return summary

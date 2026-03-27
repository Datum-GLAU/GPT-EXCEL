import pandas as pd

def analyze_data(file_path):

    df = pd.read_excel(file_path)

    summary = {
        "total_rows": len(df),
        "total_columns": len(df.columns),
        "columns": list(df.columns),
        "data_types": df.dtypes.astype(str).to_dict(),
        "null_values": df.isnull().sum().to_dict()
    }

    try:
        summary["statistics"] = df.describe().to_dict()
    except:
        summary["statistics"] = "Not applicable"

    return summary

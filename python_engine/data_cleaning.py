import pandas as pd


def clean_data(file_path: str, output_path: str = "cleaned_data.xlsx") -> str:
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        return f"Could not read file: {str(e)}"
 
    original_rows = len(df)
 
    df = df.drop_duplicates().copy()
    rows_after_dedup = len(df)

    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].fillna(df[col].median())
        elif pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].ffill().bfill()
        else:
            df[col] = df[col].fillna("Unknown")

    df.columns = [col.strip().lower().replace(" ", "_") for col in df.columns]

    try:
        df.to_excel(output_path, index=False)
    except Exception as e:
        return f"Could not save cleaned file: {str(e)}"

    removed_rows = original_rows - rows_after_dedup
    return (
        f"Data cleaned successfully. "
        f"Removed {removed_rows} duplicate rows. "
        f"Filled missing values. Saved to: {output_path}"
    )

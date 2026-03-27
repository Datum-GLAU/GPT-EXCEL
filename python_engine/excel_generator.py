# -------- BASIC EXCEL GENERATION (optional demo) --------

# def create_excel():
#     data = {
#         "Name": ["Aman", "Riya", "Rahul", "Sneha"],
#         "Score": [85, 90, 78, 92]
#     }

#     df = pd.DataFrame(data)
#     df.to_excel("output.xlsx", index=False)

#     return "Excel file generated successfully"

import pandas as pd

def read_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        return df.to_dict(orient="records")
    except Exception as e:
        return {"error": str(e)}

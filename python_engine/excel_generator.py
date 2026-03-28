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


def generate_advanced_excel(file_path):

    import pandas as pd

    df = pd.read_excel(file_path)

    with pd.ExcelWriter("final_output.xlsx", engine="xlsxwriter") as writer:

        df.to_excel(writer, sheet_name="Raw Data", index=False)

        summary = df.describe()
        summary.to_excel(writer, sheet_name="Summary")

        workbook = writer.book
        worksheet = writer.sheets["Raw Data"]

        header_format = workbook.add_format({
            'bold': True,
            'border': 1
        })

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        for i, col in enumerate(df.columns):
            worksheet.set_column(i, i, 20)

    return "Advanced Excel generated"

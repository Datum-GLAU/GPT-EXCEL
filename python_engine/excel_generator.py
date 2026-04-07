import pandas as pd


def read_excel(file_path: str, limit: int | None = None) -> list | dict:
    try:
        df = pd.read_excel(file_path)
        if limit is not None:
            df = df.head(limit)
        return df.to_dict(orient="records")
    except Exception as e:
        return {"error": str(e)}


def generate_advanced_excel(file_path: str, output_path: str = "final_output.xlsx") -> str:
    try:
        df = pd.read_excel(file_path)

        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="Raw Data", index=False)

            summary = df.describe(include="all", datetime_is_numeric=True).round(2)
            summary.to_excel(writer, sheet_name="Summary")

            workbook = writer.book
            worksheet = writer.sheets["Raw Data"]

            header_format = workbook.add_format({
                "bold": True,
                "border": 1,
                "bg_color": "#4472C4",
                "font_color": "#FFFFFF",
            })
 
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 20)

        return f"Advanced Excel generated: {output_path}"

    except Exception as e:
        return f"Error generating Excel: {str(e)}"


def create_excel_template(
    output_path: str = "excel_template.xlsx",
    rows: int = 10,
    include_sample_data: bool = True,
) -> str:
    columns = ["Name", "Category", "Amount", "Date", "Status"]
    if include_sample_data:
        data = [
            ["Item 1", "Office", 1200, "2026-01-05", "Open"],
            ["Item 2", "Sales", 950, "2026-01-12", "Closed"],
            ["Item 3", "HR", 780, "2026-02-02", "Open"],
        ]
        while len(data) < rows:
            index = len(data) + 1
            data.append([f"Item {index}", "General", 0, "2026-01-01", "Pending"])
    else:
        data = [["", "", "", "", ""] for _ in range(rows)]

    df = pd.DataFrame(data[:rows], columns=columns)

    try:
        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="Template", index=False)
            worksheet = writer.sheets["Template"]
            for idx, _ in enumerate(columns):
                worksheet.set_column(idx, idx, 18)
        return f"Excel template generated: {output_path}"
    except Exception as e:
        return f"Error generating template: {str(e)}"

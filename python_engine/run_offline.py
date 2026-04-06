import os
from analysis import analyze_data
from charts import create_chart
from report_generator import generate_report
from data_cleaning import clean_data
from excel_generator import generate_advanced_excel


def run_all(file_path: str):
    if not os.path.exists(file_path):
        print(f"[Offline] File not found: {file_path}")
        return

    print(f"\n{'='*50}")
    print(f"  GPT EXCEL — Offline Runner")
    print(f"  File: {file_path}")
    print(f"{'='*50}\n")

    print("[1/5] Analyzing data...")
    analysis = analyze_data(file_path)
    print(f"      Rows: {analysis['total_rows']} | Cols: {analysis['total_columns']}")
    print(f"      Duplicates: {analysis['duplicate_rows']}")

    print("\n[2/5] Cleaning data...")
    clean_result = clean_data(file_path, output_path="offline_cleaned.xlsx")
    print(f"      {clean_result}")

    print("\n[3/5] Generating chart...")
    chart_result = create_chart(file_path, chart_type="auto", output_path="offline_chart.png")
    print(f"      {chart_result}")

    print("\n[4/5] Generating report...")
    report_result = generate_report(file_path, output_path="offline_report.txt")
    print(f"      {report_result}")

    print("\n[5/5] Generating advanced Excel...")
    excel_result = generate_advanced_excel(file_path, output_path="offline_output.xlsx")
    print(f"      {excel_result}")

    print(f"\n{'='*50}")
    print("  All tasks complete!")
    print(f"{'='*50}\n")


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python run_offline.py <path_to_excel_file>")
        print("Example: python run_offline.py data.xlsx")
    else:
        run_all(sys.argv[1])

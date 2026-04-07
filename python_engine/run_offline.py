import os
import sys
import argparse
from analysis import analyze_data
from charts import create_chart
from report_generator import generate_report
from data_cleaning import clean_data
from excel_generator import create_excel_template, generate_advanced_excel, read_excel
from word_generator import generate_word_report
from ppt_generator import generate_ppt
from file_segmentation import get_file_info, merge_excel_files, segment_by_row_count
from SQLite_storage import save_file_record, save_analysis, save_generated_file, log_automation
from utils import ensure_directory, unique_output_path


def run_all(file_path: str, output_dir: str = "offline_outputs"):
    if not os.path.exists(file_path):
        print(f"[Offline] File not found: {file_path}")
        return

    ensure_directory(output_dir)
    out = lambda name: unique_output_path(output_dir, name)

    print(f"\n{'='*55}")
    print(f"  XtronExcel — Offline Runner v2.0")
    print(f"  File : {file_path}")
    print(f"  Out  : {output_dir}/")
    print(f"{'='*55}\n")

    file_id = save_file_record(os.path.basename(file_path), file_path)

    print("[1/8] File info ...")
    info = get_file_info(file_path)
    print(f"      {info.get('total_rows')} rows | {info.get('total_columns')} cols | {info.get('file_size_kb')} KB")

    print("\n[2/8] Analyzing data ...")
    analysis = analyze_data(file_path)
    save_analysis(file_id, "analyze", analysis)
    print(f"      Rows: {analysis['total_rows']} | Duplicates: {analysis['duplicate_rows']}")

    print("\n[3/8] Cleaning data ...")
    clean_path = out("cleaned_data.xlsx")
    msg = clean_data(file_path, output_path=clean_path)
    save_generated_file(file_id, "cleaned_excel", clean_path)
    print(f"      {msg}")

    print("\n[4/8] Generating chart ...")
    chart_path = out("chart.png")
    msg = create_chart(file_path, chart_type="auto", output_path=chart_path)
    save_generated_file(file_id, "chart", chart_path)
    log_automation("chart", "success", msg)
    print(f"      {msg}")

    print("\n[5/8] Generating text report ...")
    report_path = out("report.txt")
    msg = generate_report(file_path, output_path=report_path)
    save_generated_file(file_id, "report_txt", report_path)
    print(f"      {msg}")

    print("\n[6/8] Generating advanced Excel ...")
    excel_path = out("final_output.xlsx")
    msg = generate_advanced_excel(file_path, output_path=excel_path)
    save_generated_file(file_id, "advanced_excel", excel_path)
    print(f"      {msg}")

    print("\n[7/8] Generating Word document ...")
    word_path = out("output_report.docx")
    msg = generate_word_report(file_path, output_path=word_path)
    save_generated_file(file_id, "word_doc", word_path)
    print(f"      {msg}")

    print("\n[8/8] Generating PowerPoint ...")
    ppt_path = out("output_report.pptx")
    msg = generate_ppt(file_path, output_path=ppt_path)
    save_generated_file(file_id, "pptx", ppt_path)
    log_automation("ppt", "success", msg)
    print(f"      {msg}")

    print(f"\n{'='*55}")
    print("  All 8 tasks complete!")
    print(f"  Outputs saved in: {output_dir}/")
    print(f"  SQLite record ID: {file_id}")
    print(f"{'='*55}\n")


def run_single_command(command: str, file_path: str | None, output_dir: str, extra_files: list[str] | None = None):
    ensure_directory(output_dir)

    if command == "template":
        output_path = unique_output_path(output_dir, "excel_template.xlsx")
        print(create_excel_template(output_path=output_path))
        return

    if command == "merge":
        if not extra_files or len(extra_files) < 2:
            print("Merge requires at least two input files.")
            return
        output_path = unique_output_path(output_dir, "merged_output.xlsx")
        print(merge_excel_files(extra_files, output_path=output_path))
        return

    if not file_path:
        print(f"Command '{command}' requires an Excel file path.")
        return

    if not os.path.exists(file_path):
        print(f"Input file not found: {file_path}")
        return

    if command == "info":
        print(get_file_info(file_path))
    elif command == "preview":
        print(read_excel(file_path, limit=10))
    elif command == "analyze":
        print(analyze_data(file_path))
    elif command == "clean":
        output_path = unique_output_path(output_dir, "cleaned_data.xlsx")
        print(clean_data(file_path, output_path=output_path))
    elif command == "chart":
        output_path = unique_output_path(output_dir, "chart.png")
        print(create_chart(file_path, output_path=output_path))
    elif command == "report":
        output_path = unique_output_path(output_dir, "report.txt")
        print(generate_report(file_path, output_path=output_path))
    elif command == "excel":
        output_path = unique_output_path(output_dir, "advanced_output.xlsx")
        print(generate_advanced_excel(file_path, output_path=output_path))
    elif command == "word":
        output_path = unique_output_path(output_dir, "report.docx")
        print(generate_word_report(file_path, output_path=output_path))
    elif command == "ppt":
        output_path = unique_output_path(output_dir, "report.pptx")
        print(generate_ppt(file_path, output_path=output_path))
    elif command == "split":
        print(segment_by_row_count(file_path, chunk_size=500, output_dir=output_dir))
    elif command == "pipeline":
        run_all(file_path, output_dir)
    else:
        print(f"Unknown command: {command}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Offline utilities for the GPT Excel Python engine")
    parser.add_argument(
        "command",
        choices=["pipeline", "info", "preview", "analyze", "clean", "chart", "report", "excel", "word", "ppt", "split", "template", "merge"],
        nargs="?",
        default="pipeline",
        help="Offline command to run",
    )
    parser.add_argument("file", nargs="?", help="Path to an Excel file")
    parser.add_argument("--output-dir", default="offline_outputs", help="Directory for generated files")
    parser.add_argument("--files", nargs="*", help="Additional files for merge")
    return parser


if __name__ == "__main__":
    parser = build_parser()
    args = parser.parse_args()
    extra_files = [args.file] + (args.files or []) if args.command == "merge" and args.file else args.files
    run_single_command(args.command, args.file, args.output_dir, extra_files)

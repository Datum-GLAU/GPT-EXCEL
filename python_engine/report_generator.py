import pandas as pd
from analysis import analyze_data


def generate_report(file_path: str, output_path: str = "report.txt") -> str:
    try:
        pd.read_excel(file_path)
    except Exception as e:
        return f"Could not read file: {str(e)}"

    analysis = analyze_data(file_path)

    lines = [
        "=" * 50,
        "         XtronExcel — DATA REPORT",
        "=" * 50,
        f"File        : {file_path}",
        f"Total Rows  : {analysis['total_rows']}",
        f"Total Cols  : {analysis['total_columns']}",
        f"Duplicates  : {analysis['duplicate_rows']}",
        "",
        "COLUMNS:",
    ]

    for col in analysis["columns"]:
        dtype = analysis["data_types"].get(col, "unknown")
        nulls = analysis["null_values"].get(col, 0)
        null_pct = analysis["null_percentage"].get(col, 0.0)
        lines.append(f"  - {col} ({dtype}) | Nulls: {nulls} ({null_pct}%)")

    lines += ["", "STATISTICS:"]

    if isinstance(analysis.get("statistics"), dict):
        for col, stats in analysis["statistics"].items():
            lines.append(f"\n  [{col}]")
            for stat_name, val in stats.items():
                if val == "":
                    continue
                try:
                    rendered = round(float(val), 2)
                except Exception:
                    rendered = val
                lines.append(f"    {stat_name}: {rendered}")
    else:
        lines.append("  Not applicable")

    lines.append("\n" + "=" * 50)

    report_text = "\n".join(lines)

    try:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(report_text)
    except Exception as e:
        return f"Could not save report: {str(e)}"

    return f"Report generated: {output_path}"

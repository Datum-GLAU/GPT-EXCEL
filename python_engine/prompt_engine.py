from analysis import analyze_data
from charts import create_chart
from report_generator import generate_report
from data_cleaning import clean_data
from excel_generator import read_excel, generate_advanced_excel, create_excel_template
from word_generator import generate_word_report
from ppt_generator import generate_ppt
from file_segmentation import (
    segment_by_column, segment_by_row_count,
    segment_by_date_column, merge_excel_files, get_file_info
)


KEYWORD_MAP = {
    "analyze":  ["analyze", "analysis", "summary", "stats", "statistics", "summarize"],
    "chart":    ["chart", "graph", "plot", "visualize", "visualization", "bar", "line", "pie", "scatter"],
    "report":   ["report", "generate report", "create report", "txt report"],
    "clean":    ["clean", "cleaning", "remove duplicates", "fix nulls", "preprocess"],
    "read":     ["read", "show data", "display", "preview", "view"],
    "excel":    ["excel output", "generate excel", "advanced excel", "format excel"],
    "word":     ["word", "docx", "document", "word report", "word doc"],
    "ppt":      ["ppt", "powerpoint", "presentation", "slides", "deck"],
    "segment":  ["segment", "split", "divide", "chunk", "partition"],
    "merge":    ["merge", "combine", "join files"],
    "template": ["template", "blank excel", "sample excel", "create sheet"],
    "info":     ["info", "file info", "details", "file details"],
}


def detect_task(prompt: str) -> str:
    prompt_lower = prompt.lower()
    for task, keywords in KEYWORD_MAP.items():
        for kw in keywords:
            if kw in prompt_lower:
                return task
    return "unknown"


def detect_chart_type(prompt: str) -> str:
    prompt_lower = prompt.lower()
    if "line" in prompt_lower:   return "line"
    if "pie" in prompt_lower:    return "pie"
    if "scatter" in prompt_lower: return "scatter"
    return "auto"


def detect_segment_column(prompt: str, file_path: str) -> str | None:
    """Try to find a column name mentioned in the prompt."""
    try:
        import pandas as pd
        df = pd.read_excel(file_path)
        prompt_lower = prompt.lower()
        for col in df.columns:
            if col.lower() in prompt_lower:
                return col
    except Exception:
        pass
    return None


def process_prompt(file_path: str, prompt: str) -> dict:
    task = detect_task(prompt)

    if task == "analyze":
        return {"task": "analyze", "result": analyze_data(file_path)}

    elif task == "chart":
        chart_type = detect_chart_type(prompt)
        result = create_chart(file_path, chart_type)
        return {"task": "chart", "chart_type": chart_type, "result": result}

    elif task == "report":
        return {"task": "report", "result": generate_report(file_path)}

    elif task == "clean":
        return {"task": "clean", "result": clean_data(file_path)}

    elif task == "read":
        return {"task": "read", "result": read_excel(file_path)}

    elif task == "excel":
        return {"task": "excel", "result": generate_advanced_excel(file_path)}

    elif task == "word":
        return {"task": "word", "result": generate_word_report(file_path)}

    elif task == "ppt":
        return {"task": "ppt", "result": generate_ppt(file_path)}

    elif task == "segment":
        col = detect_segment_column(prompt, file_path)
        if col:
            return {"task": "segment", "result": segment_by_column(file_path, col)}
        return {"task": "segment", "result": segment_by_row_count(file_path, chunk_size=500)}

    elif task == "template":
        return {"task": "template", "result": create_excel_template()}

    elif task == "info":
        return {"task": "info", "result": get_file_info(file_path)}

    else:
        return {
            "task": "unknown",
            "result": (
                "Could not understand your prompt. Try: "
                "'analyze data', 'create chart', 'generate report', "
                "'clean data', 'show data', 'generate excel', "
                "'create word doc', 'create ppt', 'segment by Region', "
                "'file info'."
            ),
        }

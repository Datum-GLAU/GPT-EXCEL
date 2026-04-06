from analysis import analyze_data
from charts import create_chart
from report_generator import generate_report
from data_cleaning import clean_data
from excel_generator import read_excel, generate_advanced_excel


KEYWORD_MAP = {
    "analyze":       ["analyze", "analysis", "summary", "stats", "statistics", "summarize"],
    "chart":         ["chart", "graph", "plot", "visualize", "visualization", "bar", "line", "pie", "scatter"],
    "report":        ["report", "generate report", "create report"],
    "clean":         ["clean", "cleaning", "remove duplicates", "fix nulls", "preprocess"],
    "read":          ["read", "show data", "display", "preview", "view"],
    "excel":         ["excel output", "generate excel", "advanced excel", "format excel"],
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
    if "line" in prompt_lower:
        return "line"
    if "pie" in prompt_lower:
        return "pie"
    if "scatter" in prompt_lower:
        return "scatter"
    return "auto"


def process_prompt(file_path: str, prompt: str) -> dict:
    task = detect_task(prompt)

    if task == "analyze":
        result = analyze_data(file_path)
        return {"task": "analyze", "result": result}

    elif task == "chart":
        chart_type = detect_chart_type(prompt)
        result = create_chart(file_path, chart_type)
        return {"task": "chart", "chart_type": chart_type, "result": result}

    elif task == "report":
        result = generate_report(file_path)
        return {"task": "report", "result": result}

    elif task == "clean":
        result = clean_data(file_path)
        return {"task": "clean", "result": result}

    elif task == "read":
        result = read_excel(file_path)
        return {"task": "read", "result": result}

    elif task == "excel":
        result = generate_advanced_excel(file_path)
        return {"task": "excel", "result": result}

    else:
        return {
            "task": "unknown",
            "result": (
                "Sorry, I could not understand your prompt. "
                "Try: 'analyze data', 'create chart', 'generate report', "
                "'clean data', 'show data', or 'generate excel'."
            ),
        }

import pandas as pd
import matplotlib.pyplot as plt


def create_chart(file_path: str, chart_type: str = "auto", output_path: str = "chart.png") -> str:
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        return f"Could not read file: {str(e)}"
 
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    non_numeric_cols = df.select_dtypes(exclude="number").columns.tolist()
 
    if len(numeric_cols) == 0:
        return "No numeric data found for chart generation"
 
    x_col = non_numeric_cols[0] if non_numeric_cols else df.columns[0]
    y_col = numeric_cols[0]

    plt.figure(figsize=(10, 6))

    try:
        if chart_type == "line":
            plt.plot(df[x_col], df[y_col], marker="o", linewidth=2)
        elif chart_type == "pie":
            plt.pie(df[y_col], labels=df[x_col], autopct="%1.1f%%", startangle=140)
        elif chart_type == "scatter":
            if len(numeric_cols) >= 2:
                plt.scatter(df[numeric_cols[0]], df[numeric_cols[1]], alpha=0.7)
                plt.xlabel(numeric_cols[0])
                plt.ylabel(numeric_cols[1])
                plt.title(f"{numeric_cols[1]} vs {numeric_cols[0]}")
            else:
                plt.scatter(range(len(df)), df[y_col], alpha=0.7)
        else:
            value_counts = df[x_col].astype(str).value_counts().head(20)
            if x_col == y_col or len(value_counts) < len(df[x_col]):
                plt.bar(value_counts.index, value_counts.values, color="#4472C4")
                plt.xlabel(x_col)
                plt.ylabel("Count")
                plt.title(f"Frequency of {x_col}")
                plt.xticks(rotation=45, ha="right")
                plt.tight_layout()
                plt.savefig(output_path, dpi=150)
                return f"Chart saved: {output_path}"
            plt.bar(df[x_col].astype(str), df[y_col], color="#4472C4")

        if chart_type != "pie" and chart_type != "scatter":
            plt.xlabel(x_col)
            plt.ylabel(y_col)
            plt.title(f"{y_col} vs {x_col}")
            plt.xticks(rotation=45, ha="right")

        plt.tight_layout()
        plt.savefig(output_path, dpi=150)
        return f"Chart saved: {output_path}"

    except Exception as e:
        return f"Chart generation error: {str(e)}"
    finally:
        plt.close()

import time
from analysis import analyze_data
from charts import create_chart
from report_generator import generate_report

FILE_PATH = "uploaded_file.xlsx"

def run_automation():

    print("Automation started...")

    while True:
        try:
            print("Running scheduled tasks...")

            analyze_data(FILE_PATH)
            create_chart(FILE_PATH)
            generate_report(FILE_PATH)

            print("Tasks completed ✅")

        except Exception as e:
            print("Error:", e)

        time.sleep(60)  # every 60 seconds

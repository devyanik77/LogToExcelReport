# LogToExcelReport
machine-utilisation-reporter or factory-utilisation-dashboard.

**Machine Utilisation Reporter**

A Python-based tool to analyze machine log files, calculate runtime/utilisation, and generate Excel reports with raw data, daily summaries, and a dashboard chart.

✨ **Features**

✅ Parse machine log files (.txt) for StartTime, EndTime, and Runtime
✅ Map logs to machine IDs automatically
✅ Calculate total, average runtime, and job counts per machine
✅ Compute machine utilisation percentage against scheduled time
✅ Generate a multi-sheet Excel report with raw data, summary, and dashboard
✅ Visualise utilisation with an Excel Bar Chart

📂 **Project Structure**
machine-utilisation-reporter/
│── daily_utilisation.py   # Main script
│── requirements.txt       # Python dependencies
│── sample_logs/           # Example log files
│── reports/               # Generated Excel reports
│── README.md              # Project documentation

⚙️ Installation

**Clone the repository:**

git clone https://github.com/<your-username>/machine-utilisation-reporter.git

cd machine-utilisation-reporter


**Create a virtual environment (optional but recommended):**

python -m venv venv
source venv/bin/activate   # Linux/Mac
venv\Scripts\activate      # Windows


**Install dependencies:**

pip install -r requirements.txt


**📝 Usage**

Place all machine log files (.txt) inside a folder (e.g., sample_logs/).

Run the script:

python daily_utilisation.py sample_logs


**The tool will generate an Excel report:**

Daily_Report_YYYY-MM-DD.xlsx

📊 Example Report

Raw Data: Contains extracted log details

Summary: Daily utilisation per machine

Dashboard: Chart showing machine utilisation %

**🛠 Dependencies**

Python 3.8+

pandas

openpyxl

matplotlib

**Install via:**

pip install pandas openpyxl matplotlib

🚀 Roadmap / Future Enhancements

Add support for CSV/JSON log formats

Export PDF summary reports

Integrate with email automation to send reports

Add trend analysis (weekly/monthly)

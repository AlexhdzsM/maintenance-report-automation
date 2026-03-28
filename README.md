# 📊 Maintenance Report Automation

## 🚀 Overview

This project automates the generation of maintenance reports for medical equipment.

It processes raw maintenance data, extracts key metrics, generates visualizations, and exports a structured Excel report automatically.

Additionally, the script is configured to run on a schedule using Windows Task Scheduler, simulating a real-world reporting workflow.

---

## 🎯 Objective

In many environments, maintenance data is either underutilized or not tracked properly, making it difficult to take data-driven decisions.

This project demonstrates how to:

* Transform raw data into actionable insights
* Identify operational inefficiencies
* Automate repetitive reporting tasks

---

## 📂 Project Structure

```
maintenance-report-automation/
│
├── data/
│   └── data.csv
│
├── output/
│   ├── reporte.xlsx
│   ├── duracion_por_equipo.png
│   └── log.txt
│
├── scripts/
│   └── generate_report.py
│
└── README.md
```

---

## ⚙️ Features

* Data processing using Python (pandas)
* Aggregation of key metrics:

  * Total maintenance events
  * Unique equipment count
  * Average maintenance duration
  * Maintenance status distribution
* Automated Excel report generation
* Embedded visualization (bar chart)
* Basic formatting for readability
* Logging system for execution tracking
* Scheduled execution using Windows Task Scheduler

---

## 📊 Key Insight Example

The analysis highlights equipment with higher average maintenance duration, which may indicate:

* Operational bottlenecks
* High repair complexity
* Need for preventive maintenance strategies

---

## 🛠️ Technologies Used

* Python 3.11
* pandas
* matplotlib
* openpyxl

---

## ▶️ How to Run

### 1. Clone repository

```
git clone https://github.com/AlexhdzsM/maintenance-report-automation.git
cd maintenance-report-automation
```

### 2. Install dependencies

```
pip install pandas matplotlib openpyxl
```

### 3. Run script

```
python scripts/generate_report.py
```

---

## ⏱️ Automation

The script can be scheduled using Windows Task Scheduler to run automatically (e.g., daily at 8:00 AM).

This enables a fully automated reporting pipeline without manual intervention.

---

## 📌 Future Improvements

* Email delivery of reports
* Integration with real hospital systems
* Advanced analytics (failure prediction)
* Interactive dashboards

---

## 🧠 What I Learned

* Structuring Python scripts for maintainability
* Handling file paths and execution environments
* Automating processes outside development environments
* Debugging real-world issues (paths, permissions, scheduling)

---

## 📬 Author

Mariano Hernández

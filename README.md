# 📊 RFQ Analyzer

RFQ Analyzer is a data analytics tool built with Python, Power BI, and machine learning to automate and enhance the analysis of Request For Quotation (RFQ) data. It streamlines data ingestion, cleans and transforms data, forecasts win probabilities, and outputs detailed business insights via interactive dashboards and Excel reports.

---

## 🚀 Features

- ✅ **Automated ETL Pipeline**  
  Automates the ingestion and cleaning of over 5,000 RFQ records from a Microsoft Access database using `Python` and `APScheduler`.

- 📈 **Predictive Modeling**  
  Uses a trained Random Forest Regressor to predict RFQ win probabilities, aiding in strategic planning.

- 📊 **Interactive Dashboards (Power BI)**  
  Visualizes trends in customer behavior, factory efficiency, and sales rep performance.

- 📁 **Rich Excel Reporting**  
  Generates formatted Excel reports by customer, factory, and sales rep.

---

## 🛠 Tech Stack

- **Python**: Data processing, modeling, and scheduling
- **Power BI**: KPI and trend visualization
- **Random Forest (scikit-learn)**: Predictive modeling
- **Power Query**: Data transformation in Power BI
- **APScheduler**: Scheduling periodic ETL jobs

---

## 📂 Project Structure
rfq_analyzer/
├─ main.py # Main script to run the pipeline
├─ data/
│ ├─ database_loader.py # Loads data from Access DB
│ ├─ mock_generator.py # For testing/demo purposes
├─ model/
│ ├─ best_model.pkl # Trained Random Forest model
├─ outputs/
│ ├─ format_customer_excel.py # Formats customer reports
│ ├─ format_factory_excel.py # Formats factory reports
│ ├─ formatted/ # Final output Excel files
│ ├─ raw/ # Intermediate analytics reports
├─ processing/
│ ├─ calculate_won_amount.py # Calculates awarded amounts
│ ├─ clean_rfq.py # Data cleaning routines
│ ├─ customer_analysis.py # Customer behavior analysis
│ ├─ customer_factory_analysis.py # Combined customer-factory metrics
│ ├─ factory_analysis.py # Factory performance evaluation
│ ├─ ml_model.py # Model training & prediction
│ ├─ sales_rep_analysis.py # Sales rep effectiveness


## 🤖 Machine Learning
The tool uses a Random Forest Regression model (model/best_model.pkl) trained on historical RFQ outcomes to estimate the probability of winning new RFQs. Features include customer, factory, pricing, and historical win rates.

## 📈 Dashboards
Power BI dashboards built on the output data include:
- RFQ Overview
- Factory Performance by Year/Factory
- Sales Analysis by Year/Sales Rep
- Customer Insights by Year/Customer
- 2025 RFQ Overview
- Historical Average Win Rate
- Predicted Win Rate

## 📅 Scheduling
The ETL pipeline is scheduled via APScheduler to run at regular intervals.

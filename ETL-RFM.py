import dlt
import pandas as pd
import numpy as np
from datetime import datetime

# 1️⃣ Load Data (Extract)
file_path = "Online Retail.xlsx"
data = pd.read_excel(file_path)

# 2️⃣ Data Cleaning
def clean_data(data):
    # Handle missing values
    data = data.dropna()
    
    # Remove duplicates
    data = data.drop_duplicates()

    # Convert InvoiceDate to datetime
    data["InvoiceDate"] = pd.to_datetime(data["InvoiceDate"])
    
    # Detect and remove outliers in 'Quantity' and 'UnitPrice'
    def detect_outliers(series):
        Q1 = series.quantile(0.25)
        Q3 = series.quantile(0.75)
        IQR = Q3 - Q1
        lower_bound = Q1 - 1.5 * IQR
        upper_bound = Q3 + 1.5 * IQR
        return lower_bound, upper_bound

    for column in ["Quantity", "UnitPrice"]:
        if column in data.columns:
            lower, upper = detect_outliers(data[column])
            data = data[(data[column] >= lower) & (data[column] <= upper)]
    
    return data

cleaned_data = clean_data(data)

# 3️⃣ Load Cleaned Data into Database
pipeline = dlt.pipeline(
    pipeline_name="retail_etl",
    destination="duckdb",  # Change to "postgres", "bigquery", etc.
    dataset_name="retail_data"
)

pipeline.run(cleaned_data, table_name="cleaned_online_retail")

print("✅ Data cleaning completed and stored in database.")

# 4️⃣ RFM Segmentation
def rfm_analysis(data):
    reference_date = data["InvoiceDate"].max()
    
    rfm = data.groupby("CustomerID").agg({
        "InvoiceDate": lambda x: (reference_date - x.max()).days,  # Recency
        "InvoiceNo": "nunique",  # Frequency
        "Revenue": "sum"  # Monetary
    }).reset_index()

    # Rename columns
    rfm.columns = ["CustomerID", "Recency", "Frequency", "Monetary"]

    # Assign RFM scores
    rfm["R_Quartile"] = pd.qcut(rfm["Recency"].rank(method="first"), 4, labels=[4, 3, 2, 1])  
    rfm["F_Quartile"] = pd.qcut(rfm["Frequency"].rank(method="first"), 4, labels=[1, 2, 3, 4])  
    rfm["M_Quartile"] = pd.qcut(rfm["Monetary"].rank(method="first"), 4, labels=[1, 2, 3, 4])  

    # Combine RFM scores
    rfm["RFM_Score"] = (
        rfm["R_Quartile"].astype(int) +
        rfm["F_Quartile"].astype(int) +
        rfm["M_Quartile"].astype(int)
    )

    # Define customer segments
    def segment_customer(score):
        if score >= 10:
            return "High Value"
        elif score >= 7:
            return "Loyal Customers"
        elif score >= 4:
            return "At Risk"
        else:
            return "Lost Customers"

    rfm["Segment"] = rfm["RFM_Score"].apply(segment_customer)

    return rfm

rfm_segmented_data = rfm_analysis(cleaned_data)

# 5️⃣ Load RFM Segmented Data into Database
pipeline.run(rfm_segmented_data, table_name="rfm_segmented_customers")

print("✅ RFM segmentation completed and stored in database.")

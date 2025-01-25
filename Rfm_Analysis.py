import pandas as pd
from datetime import datetime

file_path = 'Post_EDA_Dataset.csv' 
data = pd.read_csv(file_path)

#Convert InvoiceDate to datetime
data['InvoiceDate'] = pd.to_datetime(data['InvoiceDate'])
reference_date = data['InvoiceDate'].max()

#Calculate Recency, Frequency & Monetary metrics
rfm = data.groupby('CustomerID').agg({
    'InvoiceDate': lambda x: (reference_date - x.max()).days,  
    'InvoiceNo': 'nunique',                                   
    'Revenue': 'sum'                                          
}).reset_index()

# Rename columns for clarity
rfm.columns = ['CustomerID', 'Recency', 'Frequency', 'Monetary']

#Assign RFM scores using percentiles
rfm['R_Quartile'] = pd.qcut(rfm['Recency'].rank(method='first'), 4, labels=[4, 3, 2, 1])  
rfm['F_Quartile'] = pd.qcut(rfm['Frequency'].rank(method='first'), 4, labels=[1, 2, 3, 4])  
rfm['M_Quartile'] = pd.qcut(rfm['Monetary'].rank(method='first'), 4, labels=[1, 2, 3, 4])  

#Combine RFM scores
rfm['RFM_Score'] = (
    rfm['R_Quartile'].astype(int) +
    rfm['F_Quartile'].astype(int) +
    rfm['M_Quartile'].astype(int)
)

#Define customer segments based on RFM score ranges
def segment_customer(score):
    if score >= 10:
        return 'High Value'
    elif score >= 7:
        return 'Loyal Customers'
    elif score >= 4:
        return 'At Risk'
    else:
        return 'Lost Customers'

rfm['Segment'] = rfm['RFM_Score'].apply(segment_customer)


output_file = 'Segmented_RFM_Metrics.csv'
rfm.to_csv(output_file, index=False)

print(f"RFM analysis and segmentation completed. Results saved to {output_file}.")

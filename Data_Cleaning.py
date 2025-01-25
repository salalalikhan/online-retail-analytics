import pandas as pd
import numpy as np

# Load the dataset
file_path = 'Online Retail.xlsx' 
data = pd.read_excel(file_path)

#Task 1: Explore the dataset structure
print("Dataset Overview:")
print(data.info())
print("\nSample Data:")
print(data.head())

#Task 2: Handling missing values
missing_values = data.isnull().sum()
print("\nMissing Values:\n", missing_values)

#Drop rows with significant missing values
cleaned_data = data.dropna()

#Task 3: Removing duplicates
cleaned_data = cleaned_data.drop_duplicates()
print("\nDataset after cleaning missing values and duplicates:")
print(cleaned_data.info())

#Task 4: Identify outliers in numeric columns
#Defining a function to calculate outlier thresholds
def detect_outliers(series):
    Q1 = series.quantile(0.25)
    Q3 = series.quantile(0.75)
    IQR = Q3 - Q1
    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR
    return lower_bound, upper_bound

#Check for outliers in features 'Quantity' and 'UnitPrice'
for column in ['Quantity', 'UnitPrice']:
    if column in cleaned_data.columns:
        lower, upper = detect_outliers(cleaned_data[column])
        outliers = cleaned_data[(cleaned_data[column] < lower) | (cleaned_data[column] > upper)]
        print(f"\nOutliers in {column}:")
        print(outliers)

#Final cleaned dataset
print("\nFinal cleaned dataset shape:", cleaned_data.shape)

cleaned_data.to_csv('Cleaned_Online_Retail.csv', index=False)
print("\nCleaned dataset saved as 'Cleaned_Online_Retail.csv'.")
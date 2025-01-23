import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go
import bokeh.plotting as bp
import bokeh.io as bio
from bokeh.models import ColumnDataSource, HoverTool
from mpl_toolkits.mplot3d import Axes3D
from tabulate import tabulate

# Load and understand the dataset
file_path = 'Online Retail.xlsx'  # Update the path if needed
print("Loading the dataset...")
data = pd.read_excel(file_path)
print("Dataset loaded successfully! Here's an overview:")
print(tabulate(data.info(), headers='keys', tablefmt='grid'))
print("\nSample Data:")
print(tabulate(data.head(), headers='keys', tablefmt='grid'))

# Task 1: Understand key metrics
print("\nAnalyzing the distribution of key metrics (Quantity, Unit Price):")
print(tabulate(data[['Quantity', 'UnitPrice']].describe(), headers='keys', tablefmt='grid'))

# Visualize Quantity distribution using Bokeh
print("\nCreating an interactive visualization for Quantity distribution...")
quantity_hist, edges = np.histogram(data['Quantity'], bins=50)
source = ColumnDataSource(data=dict(quantities=quantity_hist, left=edges[:-1], right=edges[1:]))
p = bp.figure(title="Distribution of Quantity", x_axis_label="Quantity", y_axis_label="Frequency", 
              width=900, height=500, background_fill_color="#0A192F", border_fill_color="#0A192F")
p.quad(top='quantities', bottom=0, left='left', right='right', source=source, fill_color="#1F77B4", line_color="#1F77B4")
p.add_tools(HoverTool(tooltips=[("Quantity", "@left"), ("Frequency", "@quantities")]))
p.xgrid.grid_line_color = None
p.ygrid.grid_line_color = None
p.xaxis.axis_label_text_color = "#E5E7EB"
p.yaxis.axis_label_text_color = "#E5E7EB"
p.title.text_color = "#FFFFFF"
p.xaxis.major_label_text_color = "#E5E7EB"
p.yaxis.major_label_text_color = "#E5E7EB"
bio.show(p)

# Calculate Revenue if not already present
if 'Revenue' not in data.columns:
    print("\nCalculating revenue for each transaction...")
    data['Revenue'] = data['Quantity'] * data['UnitPrice']

# Advanced 3D scatter plot for Quantity, Revenue, and Unit Price
print("\nCreating an advanced 3D scatter plot...")
fig = px.scatter_3d(data, x='Quantity', y='Revenue', z='UnitPrice', color='Revenue', 
                    title='3D Scatter: Quantity vs Revenue vs UnitPrice', template='plotly_dark', 
                    color_continuous_scale=px.colors.sequential.Plasma)
fig.update_layout(title_font_size=20, font_color="#FFFFFF", scene=dict(
    xaxis=dict(backgroundcolor="#0A192F", gridcolor=None, title_font=dict(color="#E5E7EB")),
    yaxis=dict(backgroundcolor="#0A192F", gridcolor=None, title_font=dict(color="#E5E7EB")),
    zaxis=dict(backgroundcolor="#0A192F", gridcolor=None, title_font=dict(color="#E5E7EB"))))
fig.show()

# Visualize Revenue distribution by Country
if 'Country' in data.columns:
    print("\nVisualizing revenue distribution by country...")
    country_revenue = data.groupby('Country')['Revenue'].sum().reset_index()
    source = ColumnDataSource(country_revenue)
    p = bp.figure(x_range=country_revenue['Country'], title="Revenue Distribution by Country", 
                  x_axis_label="Country", y_axis_label="Revenue", width=900, height=500, background_fill_color="#0A192F", border_fill_color="#0A192F")
    p.vbar(x='Country', top='Revenue', width=0.8, source=source, color="#FF7F0E")
    p.xgrid.grid_line_color = None
    p.ygrid.grid_line_color = None
    p.xaxis.major_label_orientation = "vertical"
    p.add_tools(HoverTool(tooltips=[("Country", "@Country"), ("Revenue", "@Revenue")]))
    p.xaxis.axis_label_text_color = "#E5E7EB"
    p.yaxis.axis_label_text_color = "#E5E7EB"
    p.title.text_color = "#FFFFFF"
    p.xaxis.major_label_text_color = "#E5E7EB"
    p.yaxis.major_label_text_color = "#E5E7EB"
    bio.show(p)

# Task 3: Highlight top-performing products
print("\nIdentifying top-performing products...")
top_products = data.groupby('Description')['Revenue'].sum().sort_values(ascending=False).head(10)
print("\nTop 10 Products by Revenue:")
print(tabulate(top_products.reset_index(), headers='keys', tablefmt='grid'))

fig = px.bar(top_products, x=top_products.values, y=top_products.index, orientation='h', 
             title='Top 10 Products by Revenue', template='plotly_dark', color=top_products.values, 
             color_continuous_scale=px.colors.diverging.Spectral)
fig.update_layout(title_font_size=20, xaxis_title='Total Revenue', yaxis_title='Product Description', font_color="#FFFFFF")
fig.show()

# Highlight top-performing countries
if 'Country' in data.columns:
    print("\nIdentifying top-performing countries...")
    top_countries = data.groupby('Country')['Revenue'].sum().sort_values(ascending=False).head(10)
    print("\nTop 10 Countries by Revenue:")
    print(tabulate(top_countries.reset_index(), headers='keys', tablefmt='grid'))

    fig = px.bar(top_countries, x=top_countries.values, y=top_countries.index, orientation='h', 
                 title='Top 10 Countries by Revenue', template='plotly_dark', color=top_countries.values, 
                 color_continuous_scale=px.colors.sequential.Viridis)
    fig.update_layout(title_font_size=20, xaxis_title='Total Revenue', yaxis_title='Country', font_color="#FFFFFF")
    fig.show()

# Save results
print("\nSaving the cleaned and updated dataset...")
data.to_csv('Post_EDA_Dataset', index=False)
print("Dataset saved successfully as 'Post_EDA_Dataset'.")
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import pandas as pd


output_path='\\'
# Load the Excel file
file_path = 'C:\\Users\\RaphaelThiney\\OneDrive - Unison Energy\\Documents - O&M\\Prognostics/LivePrognosticsBook.xlsx'
excel_data = pd.ExcelFile(file_path)

# Extract sheet names
sheet_names = excel_data.sheet_names
 
# Set plot style
sns.set(style="whitegrid")

prognostics_df = pd.read_excel(file_path, sheet_name='Prognostics')
budget_df = pd.read_excel(file_path, sheet_name='Budget')
engines_df = pd.read_excel(file_path, sheet_name='Engines')
sites_df = pd.read_excel(file_path, sheet_name='Sites')

# Display the first few rows of each dataframe
prognostics_preview = prognostics_df.head()
budget_preview = budget_df.head()
engines_preview = engines_df.head()
sites_preview = sites_df.head()


# Convert DateTime to datetime format
prognostics_df['DateTime'] = pd.to_datetime(prognostics_df['DateTime'])

# List of metrics to plot
metrics = [
    'Total CHP Production',
    'Building Load Current (kW)',
    'Utility Import Current (kW)',
    'CHP1 Electricity Production Current (kW)',
    'CHP1 Gas Consumption Current (m3h)'
]

# Create a time series plot for each metric
for metric in metrics:
    plt.figure(figsize=(10, 6))
    for site in prognostics_df['Site'].unique():
        site_data = prognostics_df[prognostics_df['Site'] == site]
        plt.plot(site_data['DateTime'], site_data[metric], label=site)
    
    plt.title(f'{metric} Over Time by Site')
    plt.xlabel('DateTime')
    plt.ylabel(metric)
    plt.legend()
    plt.grid(True)
    temp_file = f'{output_path}{site}_{metric} Over Time by Site.png'
    plt.savefig(temp_file)
    #plt.show()

 
# Create bar charts for comparing average values of key metrics across sites
for metric in metrics:
    plt.figure(figsize=(12, 6))
    sns.barplot(x='Site', y=metric, data=prognostics_df, estimator=np.mean, ci=None)
    plt.title(f'Average {metric} by Site')
    plt.xlabel('Site')
    plt.ylabel(f'Average {metric}')
    plt.xticks(rotation=90)
    temp_file = f'{output_path}{site}_Average {metric} by Site.png'
    plt.savefig(temp_file)
    #plt.show()

# Create box plots for showing the distribution and variability of key metrics for each site
for metric in metrics:
    plt.figure(figsize=(12, 6))
    sns.boxplot(x='Site', y=metric, data=prognostics_df)
    plt.title(f'Distribution of {metric} by Site')
    plt.xlabel('Site')
    plt.ylabel(metric)
    plt.xticks(rotation=90)
    temp_file = f'{output_path}{site}_Distribution of {metric} by Site.png'
    plt.savefig(temp_file)
    #plt.show()
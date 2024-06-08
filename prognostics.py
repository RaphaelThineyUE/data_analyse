import os 
import pandas as pd
import seaborn as sns 
import matplotlib.pyplot as plt

# Load the data
output_path='c:/source/data_analyse/output/'
file_path = 'C:/Users/RaphaelThiney/OneDrive - Unison Energy/Documents - O&M/Prognostics/LivePrognosticsBook.xlsx'
prognostics_df = pd.read_excel(file_path, sheet_name='Prognostics')

# Convert DateTime to datetime format
prognostics_df['DateTime'] = pd.to_datetime(prognostics_df['DateTime'])

# Function to calculate and plot correlation matrix for each site
def plot_correlation_matrix(site_data, site_name):
    numeric_data = site_data.select_dtypes(include=[float, int])  # Select only numeric columns
    correlation_matrix = numeric_data.corr()
    plt.figure(figsize=(12, 10))
    sns.heatmap(correlation_matrix, annot=True, cmap='coolwarm', vmin=-1, vmax=1)
    plt.title(f'Correlation Matrix for {site_name}')
    plt.show()

# List of sites
sites = prognostics_df['Site'].unique()

# Plot correlation matrices for each site and create HTML report with all plot.
# Create HTML report
html_report = '<html><body>'
for site in sites:
    try:
        site_data = prognostics_df[prognostics_df['Site'] == site]
        plot_correlation_matrix(site_data, site)
        # Save the plot as a temporary image file
        temp_file = f'{output_path}{site}_correlation_matrix.png'
        plt.savefig(temp_file)
        # Embed the image in the HTML report
        html_report += f'<h2>{site}</h2>'
        html_report += f'<img src="{temp_file}" alt="Correlation Matrix">'
        # Remove the temporary image file
        #os.remove(temp_file)
    except Exception as e:
        print(f"Error processing site {site}: {str(e)}")
html_report += '</body></html>'

    # Save the HTML report to a file
report_file = f'{output_path}correlation_report.html'
with open(report_file, 'w') as f:
    f.write(html_report)
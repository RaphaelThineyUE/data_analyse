import pyodbc
import pandas as pd

# Connect to the database
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=PSQLE01.unisonenergy.local;DATABASE=DBUEWH;UID=power;PWD=rP96345*&')
 
# Define SQL query
sql_query = "SELECT * FROM Dim_InstalledProducts"

# Execute SQL query and fetch data into a DataFrame
data = pd.read_sql(sql_query, conn)

# Close the connection
conn.close()

# Perform data analysis
# For example, calculate descriptive statistics

# Generate visualizations
# For example, create a histogram
import matplotlib.pyplot as plt
plt.hist(data['column_name'])
plt.title('Histogram of Column')
plt.xlabel('Values')
plt.ylabel('Frequency')
plt.show()

# Discuss insights with ChatGPT

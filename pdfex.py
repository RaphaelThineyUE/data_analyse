import camelot
import os 
import pandas as pd

def extract_tables_from_pdfs(directory):
    # Define the tables of interest
    tables_of_interest = [
        "Transmitted data point assignment table",
        "Transmitted data assignment",
        "Message texts 0-2999"
    ]

    # Create an empty dictionary to store dataframes
    tables_dict = {table: [] for table in tables_of_interest}

    # Iterate over all PDF files in the directory
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            file_path = os.path.join(directory, filename)
            print(f"Processing {file_path}...")
            
            # Read the PDF file with Camelot
   
            tables = camelot.io.read_pdf(file_path, pages='all', flavor='stream')

            # Extract the tables of interest
            for table in tables:
                for table_name in tables_of_interest:
                    if table_name in table.df.values:
                        tables_dict[table_name].append(table.df)

    # Create an Excel writer
    output_file = os.path.join(directory, "extracted_tables.xlsx")
    with pd.ExcelWriter(output_file) as writer:
        for table_name, dfs in tables_dict.items():
            if dfs:
                combined_df = pd.concat(dfs, ignore_index=True)
                combined_df.to_excel(writer, sheet_name=table_name, index=False)

    print(f"Tables extracted and saved to {output_file}")

# Set the directory containing the PDF files
pdf_directory = f'C:\\Users\\RaphaelThiney\\OneDrive - Unison Energy\\Cyxterra BOS 1A\\'

# Run the extraction function
extract_tables_from_pdfs(pdf_directory)

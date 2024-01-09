# Import necessary libraries
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans

# Load the Excel file into a DataFrame
excel_file_path = r'C:\Users\vinit\Desktop\problem_statement_1_and_2\CategoryCode_Mapping.xlsx'
excel_file = pd.ExcelFile(excel_file_path)

# Print all sheet names in the Excel file
print("Sheet names in the Excel file:", excel_file.sheet_names)

# Check for the 'DisplayLable' column in each sheet
found_display_label = False
for sheet_name in excel_file.sheet_names:
    df_sheet = pd.read_excel(excel_file, sheet_name)
    if 'DisplayLable' in df_sheet.columns:
        print(f"\nFound 'DisplayLable' column in sheet: {sheet_name}")
        found_display_label = True
        break

if not found_display_label:
    print("\nError: 'DisplayLable' column not found in any sheet.")
else:
    # Load the specific sheet with 'DisplayLable' into df
    df = pd.read_excel(excel_file_path, sheet_name)

    # Display the first few rows of the DataFrame to understand the structure
    print(df.head())

    # Ensure that the Excel file has a column named 'DisplayLable' containing the text data
    # Print the column names and the first few rows for verification
    print(df.columns)
    print(df.head())

    # Check if 'DisplayLable' is in the DataFrame columns
    if 'DisplayLable' not in df.columns:
        print("Error: 'DisplayLable' column not found in the DataFrame.")
    else:
        # Text Vectorization using TF-IDF
        vectorizer = TfidfVectorizer(stop_words='english')
        
        # Ensure 'DisplayLable' column is not empty or contains NaN values
        if df['DisplayLable'].isnull().any():
            print("Error: 'DisplayLable' column contains NaN values.")
        else:
            # Perform vectorization
            X = vectorizer.fit_transform(df['DisplayLable'].astype(str))

            # Topic Modeling using K-Means Clustering
            num_clusters = 3  # You may adjust the number of clusters based on your needs
            kmeans = KMeans(n_clusters=num_clusters, random_state=42)
            df['Cluster'] = kmeans.fit_predict(X)

            # Print the result
            # Print the result
# Print the result
print(df[['Id', 'DisplayLable', 'Cluster']])



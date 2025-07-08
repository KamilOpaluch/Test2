import pandas as pd
import re

# Load Excel file and specific sheet
file_path = r'C:\Users\<YourUsername>\Documents\BRD_Queries_Appendix.xlsx'  # Update path
sheet_name = 'CGML_Backtesting'

# Load sheet
df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')

# Clean all cells to string
df_cleaned = df.applymap(lambda x: str(x).strip() if pd.notnull(x) else "")

# Detect the column with most entries starting with 'SELECT' (ignore case, allow whitespace)
def starts_with_select(val):
    return bool(re.match(r'^\s*select\b', val, re.IGNORECASE))

select_counts = df_cleaned.apply(lambda col: col.apply(starts_with_select).sum())
target_column = select_counts.idxmax()

# Filter rows in that column that start with SELECT
query_series = df_cleaned[target_column]
query_series = query_series[query_series.apply(starts_with_select)].reset_index(drop=True)

# Extract [columns] and table names after FROM (including joins, subqueries simplified)
def extract_info(query):
    columns = re.findall(r'\[(.*?)\]', query)
    from_clauses = re.findall(r'\bFROM\s+([^\s;\n\r]+)', query, re.IGNORECASE)
    return columns, from_clauses

# Apply extraction
results = query_series.apply(lambda q: pd.Series(extract_info(q), index=['Columns', 'Tables']))

# Final DataFrame
final_df = pd.concat([query_series, results], axis=1)
final_df.columns = ['Query', 'Columns', 'Tables']

# Display or export
print(final_df)
# Optionally: final_df.to_excel("parsed_queries.xlsx", index=False)

import pandas as pd

# Load the sheets from the Excel file
Sheet1 = pd.read_excel('C:/Projects Sara/Comparison/Book.xlsx', sheet_name='Sheet1')
Sheet2 = pd.read_excel('C:/Projects Sara/Comparison/Book.xlsx', sheet_name='Sheet2')

# Assuming there is a unique identifier column called 'CA Number'
unique_identifier = 'CA Number'

# Merge sheets on the unique identifier
comparison = pd.merge(Sheet1, Sheet2, on=unique_identifier, how='outer', suffixes=('_Sheet1', '_Sheet2'))

# Identify differences in each column (assuming both sheets have the same columns)
for column in Sheet1.columns:
    if column != unique_identifier:
        comparison[f'Difference_in_{column}'] = comparison[f'{column}_Sheet1'] != comparison[f'{column}_Sheet2']

# Identify missing data
comparison['Missing_in_Sheet_one'] = comparison.apply(lambda row: pd.isna(row).any() and pd.notna(row).all(), axis=1)
comparison['Missing_in_Sheet_two'] = comparison.apply(lambda row: pd.notna(row).any() and pd.isna(row).all(), axis=1)

# Save comparison result to a new Excel file
comparison.to_excel('comparison_result.xlsx', index=False)

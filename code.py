import pandas as pd
import os
import matplotlib.pyplot as plt
import seaborn as sns

file_path = '/Users/amelia/Documents/python_project/python_project_final_edit.xlsx'

try:
    # Excel file
    df = pd.read_excel(file_path, sheet_name='Sheet1')
    display(df.head()) 
except Exception as e:
    print(f"Error reading the Excel file: {e}")

##Data cleaning

try:
    # Fill NaN values for categorical columns with 'Unknown' and for numerical columns with the median
    for column in df.columns:
        if df[column].dtype == 'object':
            df[column].fillna('Unknown', inplace=True)
        else:
            df[column].fillna(df[column].median(), inplace=True)

    print("Cleaned Data:")
    display(df.head())

    # Save the cleaned data to a new Excel file
    cleaned_file_path = '/Users/amelia/Documents/python_project/cleaned_python_project_final_edit.xlsx'
    df.to_excel(cleaned_file_path, index=False)
    print(f"Cleaned data saved to {cleaned_file_path}")

except FileNotFoundError:
    print(f"The file at path {file_path} was not found. Please check the file path and try again.")
except Exception as e:
    print(f"Error cleaning the data: {e}")

##Data Analysis

#Count Plot - Gender vs. Shopping Venues
plt.figure(figsize=(10, 6))
sns.countplot(data=df, x='gender', hue='shopping_venues', palette='viridis')
plt.title('Gender vs. Shopping Venues')
plt.xlabel('Gender')
plt.ylabel('Count')
plt.tight_layout()
plt.show()

#Bar Plot - Influence of Sustainability Labels by Age Group
plt.figure(figsize=(12, 6))
sns.barplot(data=df, x='age', y='influence_of_sustainability_labels', palette='viridis')
plt.title('Influence of Sustainability Labels by Age Group')
plt.xlabel('Age Group')
plt.ylabel('Influence of Sustainability Labels')
plt.tight_layout()
plt.show()

# Bar plot - Occupation by Importance Of Sustainability
plt.figure(figsize=(12, 6))
sns.barplot(data=df, x='occupation', y='importance_of_sustainability', palette='viridis')
plt.title('Importance of Sustainability by Occupation')
plt.xlabel('Occupation')
plt.ylabel('Importance of Sustainability')
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

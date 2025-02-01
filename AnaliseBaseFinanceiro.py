# Script convertido do Jupyter Notebook

# Packages
import pandas as pd
import seaborn as srn
import statistics as sts
import os
import shutil  

# Checking if the data path exists 
file_path = '.\\BaseFinanceiro\\Financeiro.xlsx'
print(os.path.exists(file_path))

# Importing the entire dataset
dataset = pd.read_excel(file_path, sheet_name=None)

# Checking all sheets
print(dataset.keys())
print(len(dataset.keys()))  # 0-4

# Defining variables for each sheet
df_cliente = pd.read_excel(file_path, sheet_name=0)
df_fornecedor = pd.read_excel(file_path, sheet_name=1)
df_banco = pd.read_excel(file_path, sheet_name=2)
df_pagamentos = pd.read_excel(file_path, sheet_name=3)
df_recebimentos = pd.read_excel(file_path, sheet_name=4)

## Verifying one variable at a time
#df_cliente
#df_fornecedor
#df_banco
#df_pagamentos
#df_recebimentos

## Summary of dataset information to understand columns, data types, and missing values
print("\nDataset Information:")
df_cliente.info()
#df_banco.info()
#df_fornecedor.info()   
#df_recebimentos.info()
#df_pagamentos.info()

# Creating a calendar
start_date = pd.Timestamp('2000-01-01')  # Fixed start date
end_date = pd.Timestamp.today()  # End date as the current date

calendar = pd.DataFrame({
    'Date': pd.date_range(start=start_date, end=end_date, freq='D')  # Daily frequency
})

# Adding columns for Year, Month, Day, and Year_Month
calendar['Year'] = calendar['Date'].dt.year
calendar['Month'] = calendar['Date'].dt.month
calendar['Day'] = calendar['Date'].dt.day
calendar['Year_Month'] = calendar['Date'].dt.to_period('M')

# Displaying the first few rows of the calendar
calendar.head()

# Saving the calendar base to Excel (optional)
# output_path = './Base_Calendario.xlsx'
# calendar.to_excel(output_path, index=False)
# print(f"Calendar base saved at: {output_path}")


dataset.keys()

# Checking the dataset
df_recebimentos

# Checking for null values
df_recebimentos.isnull().sum()

# Copying the DataFrame for formatted display
df_recebimentos_display = df_recebimentos.copy()

# Formatting the values in the copy
df_recebimentos_display['Valor da Movimentação'] = df_recebimentos_display['Valor da Movimentação'].map(
    lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
)

# Converting dates
columns_to_convert = df_recebimentos.columns[2:5]
df_recebimentos[columns_to_convert] = df_recebimentos[columns_to_convert].apply(
    pd.to_datetime, format='%d/%m/%Y', errors='coerce'
)

# Displaying the Receipts table
df_recebimentos_display

# Grouping by 'Issue Date' (uncomemment the variables to group by other columns)
grouped = df_recebimentos.groupby(['Data de Emissao']).size()
#grouped = df_recebimentos.groupby(['Data da Movimentação']).size()
#grouped = df_recebimentos.groupby(['Data de Vencimento']).size()

grouped

# Descriptive statistics of numerical features
print('\nDescriptive statistics of numerical features:')
df_recebimentos[['Id Cliente', 'Valor da Movimentação']].describe()

# Boxplot of Transaction Value
srn.boxplot(df_recebimentos['Valor da Movimentação'], color='gray').set_title('Valor da Movimentação')

# Histogram of Transaction Value
srn.histplot(df_recebimentos['Valor da Movimentação'])

# Calculating Q1, Q3, and IQR
Q1 = df_recebimentos['Valor da Movimentação'].quantile(0.25) # First quartile 
Q3 = df_recebimentos['Valor da Movimentação'].quantile(0.75) # Third quartile

IQR = Q3 - Q1 # Interquartile range 
print(f'IQR = {IQR}')

# Defining limits for outliers
upper_limit_1 = Q3 + 1.5 * IQR
lower_limit_1 = Q1 - 1.5 * IQR

print(f'limite superior= {upper_limit_1} e limite inferior = {lower_limit_1}.')

# Identifying outliers
df_recebimentos['Eh_outlier'] = df_recebimentos['Valor da Movimentação'].apply(
    lambda x: 'Sim' if (x < lower_limit_1 or x > upper_limit_1) else 'Não'
)
df_recebimentos[['Valor da Movimentação', 'Eh_outlier']]

# Counting outliers
df_recebimentos.groupby(['Eh_outlier']).size()

# Histrogram of outliers
srn.histplot(df_recebimentos['Eh_outlier'])

# Filtering data within the interval for processing without outliers
values_without_outliers_1 = df_recebimentos['Valor da Movimentação'][
    (df_recebimentos['Valor da Movimentação'] >= lower_limit_1) |
    (df_recebimentos['Valor da Movimentação'] <= upper_limit_1)
]

# Statistical calculations
mean_1 = sts.mean(df_recebimentos['Valor da Movimentação'])
median_1 = sts.median(df_recebimentos['Valor da Movimentação'])
std_dev_1 = sts.stdev(df_recebimentos['Valor da Movimentação'])
diff_std_median_1 = abs(std_dev_1 - median_1)

# Statistical calculations without outliers
mean_without_outrs_1 = sts.mean(values_without_outliers_1)
median_without_outrs_1 = sts.median(values_without_outliers_1)
std_dev_without_outrs_1 = sts.stdev(values_without_outliers_1)
diff_std_median_without_outrs_1 = abs(std_dev_without_outrs_1 - median_without_outrs_1)

# Displaying results with outliers
print(f'Mean: {mean_1:.2f}')
print(f'Median: {median_1:.2f}')
print(f'Standard Deviation: {std_dev_1:.2f}')
print(f'Absolute difference between standard deviation and median: {diff_std_median_1:.2f}')
print()  # Empty line

# Displaying results without outliers
print(f"Mean without outliers: {mean_without_outrs_1:.2f}")
print(f"Median without outliers: {median_without_outrs_1:.2f}")
print(f"Standard Deviation without outliers: {std_dev_without_outrs_1:.2f}")
print(
    f'Absolute difference between standard deviation and median: '
    f'{diff_std_median_without_outrs_1:.2f}'
)

# Checking if standard deviation is less than IQR
if std_dev_1 < IQR:
    print('It is less')
else:
    print('It is not less')

if std_dev_1 < (1.5 * median_1):
    print('It is less')
else:
    print('It is not less')

# Coefficient of Variation (CV)
coef_variation_1 = std_dev_1 / mean_1
print(f"Coefficient of Variation (CV): {coef_variation_1:.2f}")

if coef_variation_1 > 0.5:
    print("Outliers may be distorting the data.")
else:
    print("The data is consistent.")

print()  # Empty line

if std_dev_without_outrs_1 < IQR:
    print('It is less')
else:
    print('It is not less')

if std_dev_without_outrs_1 < (1.5 * mean_without_outrs_1):
    print('It is less')
else:
    print('It is not less')

# Coefficient of Variation (CV)
coef_variation_2 = std_dev_without_outrs_1 / mean_without_outrs_1
print(f"Coefficient of Variation (CV): {coef_variation_2:.2f}")
if coef_variation_2 > 0.5:
    print("Outliers may be distorting the data.")
else:
    print("The data is consistent.")


# Identifying outlier clients
outlier_clients_id = df_recebimentos[df_recebimentos['Eh_outlier'] == 'Sim']['Id Cliente'].value_counts()
print(f"Outlier clients:\n{outlier_clients_id}")

# Total received by outlier clients
total_received_clients = df_recebimentos.groupby('Id Cliente')['Valor da Movimentação'].sum()

# Sort values ​​in descending order
total_received_clients = total_received_clients.sort_values(ascending=False)

# Display the first 10 values ​​of the total received by the customer
total_received_clients.head(10)

# Frequency percentage of outliers and "liers"
# Total received by outlier customers
freq_outrs_1 = len(df_recebimentos[df_recebimentos['Eh_outlier'] == 'Sim'])

# Calculate the percentage of outliers
percent_outrs_1 = (freq_outrs_1 / len(df_recebimentos)) * 100

# Display the percentage of outliers
print(f"Percentual de Outliers: {percent_outrs_1:.2f}%")
print(f'Percentual de "liers": {(100 - percent_outrs_1):.2f}%')

# EXPLORING THE CONTRIBUTION OF OUTLIERS TO THE TOTAL RECEIVED
# Total received by outlier customer
otlrs_1 = df_recebimentos[df_recebimentos['Eh_outlier'] == 'Sim'].copy()

# Add columns for Year and Month
otlrs_1['Ano'] = df_recebimentos['Data da Movimentação'].dt.year
otlrs_1['Mês'] = df_recebimentos['Data da Movimentação'].dt.month

# Group and Display total received by year
outliers_by_year = otlrs_1.groupby('Ano')['Valor da Movimentação'].sum()
print(outliers_by_year)
print() # Linha vazia

# Group and Display the total received per month
outliers_by_month = otlrs_1.groupby(['Mês'])['Valor da Movimentação'].sum()
print(f'{outliers_by_month}')

# Variable for customers who are not outliers (the "liers")
normal_clients = df_recebimentos[df_recebimentos['Eh_outlier'] == 'Não'].copy()

# Add columns for Year and Month
normal_clients['Ano'] = df_recebimentos['Data da Movimentação'].dt.year
normal_clients['Mês'] = df_recebimentos['Data da Movimentação'].dt.month

# Group and Display total received by year
normal_clients_year = normal_clients.groupby('Ano')['Valor da Movimentação'].sum()
print(normal_clients_year)
print() # Empty line

# Group and Display the total received per month
normal_clients_month = normal_clients.groupby('Mês')['Valor da Movimentação'].sum()
print(normal_clients_month) 

# Exploring the data frame
df_pagamentos

# Verifying null values
df_pagamentos.isnull().sum()

# Copying the DataFrame for formatted display
df_pagamentos_display = df_pagamentos.copy()
# Formatting the values in the copy
df_pagamentos_display['Valor da Movimentação'] = df_pagamentos_display['Valor da Movimentação'].map(
    lambda x: f'R$ {x:,.2f}'.replace(',', 'X').replace('.', ',').replace('X','.')
)
df_pagamentos_display

# Grouping (uncomemment the variables to group by other columns) 
grouped = df_pagamentos.groupby(['Data de Emissao']).size()
#agrupamento = df_pagamentos.groupby(['Data da Movimentação']).size()
#agrupamento = df_pagamentos.groupby(['Data de Vencimento']).size()

grouped

# Descriptive statistics of numerical features
print('\nDescriptive statistics of numerical features:')
df_pagamentos[['Id Fornecedor', 'Valor da Movimentação']].describe()

# Boxplot of Transaction Value
srn.boxplot(df_pagamentos['Valor da Movimentação'], color='gray').set_title('Valor da Movimentação')

# Histogram of Transaction Value
srn.histplot(df_pagamentos['Valor da Movimentação'])

# Calculating Q1, Q3, and IQR
Q1 = df_pagamentos['Valor da Movimentação'].quantile(0.25) # First quartile
Q3 = df_pagamentos['Valor da Movimentação'].quantile(0.75) # Third quartile

IQR = Q3 - Q1 # Amplitude interquartil
print(f'IQR = {IQR}')

# Calculating limits for outliers
upper_limit_2 = Q3 + 1.5 * IQR 
lower_limit_2 = Q1 - 1.5 * IQR 
print(f'limite superior= {upper_limit_2} e limite inferior = {lower_limit_2}.')

# Defining limits for outliers
df_pagamentos['Valor da Movimentação'][
    (df_pagamentos['Valor da Movimentação'] <= lower_limit_2) |
    (df_pagamentos['Valor da Movimentação'] >= upper_limit_2)
]

# Identifying outliers
df_pagamentos['Eh_outlier'] = df_pagamentos['Valor da Movimentação'].apply(
    lambda x: 'Sim' if (x < lower_limit_2 or x > upper_limit_2) else 'Não'
)
df_pagamentos[['Valor da Movimentação', 'Eh_outlier']]

# Counting outliers
df_pagamentos.groupby(['Eh_outlier']).size()

# Histogram of outliers
srn.histplot(df_pagamentos['Eh_outlier'])

# Statistical calculations
mean_2 = sts.mean(df_pagamentos['Valor da Movimentação'])  # Mean
median_2 = sts.median(df_pagamentos['Valor da Movimentação'])  # Median
std_dev_2 = sts.stdev(df_pagamentos['Valor da Movimentação'])  # Standard deviation
dif_desv_median_2 = abs(std_dev_2 - median_2)  # Difference between standard deviation and median

# Displaying results
print(f'Mean: {mean_2: .2f}')
print(f'Median: {median_2: .2f}')
print(f'Standard Deviation: {std_dev_2: .2f}')
print(f'Absolute difference between standard deviation and median: {dif_desv_median_2: .2f}')
print()  # Empty line

# Statistical calculations without outliers
mean_without_otlrs_2 = sts.median(otlrs_1)  # Mean without outliers
median_without_otlrs_2 = sts.mean(otlrs_1)  # Median without outliers
std_dev_without_otlrs_2 = sts.stdev(otlrs_1)  # Standard deviation without outliers
diff_stdmedian_without_otlrs_2 = abs(std_dev_without_otlrs_2 - median_without_otlrs_2)  # Difference between standard deviation and median

# Displaying results without outliers
print(f"Mean without outliers: {mean_without_otlrs_2: .2f}")
print(f'Median without outliers: {median_without_otlrs_2: .2f}')
print(f'Standard deviation without outliers: {std_dev_without_otlrs_2: .2f}')
print(f'Absolute difference between standard deviation and median: {diff_stdmedian_without_otlrs_2: .2f}')

# Checking if standard deviation is less than IQR
if std_dev_2 < IQR:
    print('It is less')
else:
    print('It is not less')

# Checking if the standard deviation is less than 1.5 times the median
if std_dev_2 < (1.5 * median_2):
    print('It is less')
else:
    print('It is not less')

# Coefficient of Variation (CV)
coef_variation_3 = std_dev_2 / mean_2 
print(f'Coefficient of Variation (CV): {coef_variation_3:.2f}')

if coef_variation_3 > 0.5:
    print("Outliers may be distorting the data.")
else:
    print("The data is consistent.")

print()  # Empty line

# Without outliers
# Checking if standard deviation is less than IQR
if std_dev_without_otlrs_2 < IQR:
    print('It is less')
else:
    print('It is not less')

# Checking if the standard deviation is less than 1.5 times the median
if std_dev_without_otlrs_2 < (1.5 * mean_without_otlrs_2):
    print('It is less')
else:
    print('It is not less')

# Coefficient of Variation (CV)
coef_variacao_4 = std_dev_without_otlrs_2 / mean_without_otlrs_2 
print(f'Coefficient of Variation (CV): {coef_variacao_4:.2f}')

# Checking if the data is consistent or if there are outliers
if coef_variacao_4 > 0.5:
    print('Outliers may be distorting the data.')
else:
    print('The data is consistent.')

# Frequency percentage of outliers and "liers"
# Counting and identifying outlier suppliers
freq_otlrs_pagamentos = len(df_pagamentos[df_pagamentos['Eh_outlier'] == 'Sim'])

# Calculate outlier percentage
percent_otlrs_pagamentos = (freq_otlrs_pagamentos / len(df_pagamentos)) * 100

# Display the percentage of outliers
print(f'Percentage of outliers: {percent_otlrs_pagamentos:.2f}%')
print(f'Percentage of "liers": {(100 - percent_otlrs_pagamentos):.2f}%')

# Identifying outlier suppliers
outliers_fornecedor_id = df_pagamentos[df_pagamentos['Eh_outlier'] == 'Sim']['Id Fornecedor'].value_counts().head(10)
print(f'Outlier suppliers:\n{outliers_fornecedor_id}')

# Group by 'Supplier Id' and add the values
total_supplier_expense = df_pagamentos.groupby('Id Fornecedor')['Valor da Movimentação'].sum()

# Sort from highest to lowest
total_supplier_expense = total_supplier_expense.sort_values(ascending=False)
total_supplier_expense.head(10)

# Filter the data without 'liers'
otlrs_2 = df_pagamentos[df_pagamentos['Eh_outlier'] == 'Sim'].copy()

# Add columns for Year and Month
otlrs_2['Ano'] = df_pagamentos['Data da Movimentação'].dt.year 
otlrs_2['Mês'] = df_pagamentos['Data da Movimentação'].dt.month

# Group by Year and sum the values
otlrs_year_pagamentos = otlrs_2.groupby('Ano')['Valor da Movimentação'].sum()
print(otlrs_year_pagamentos)

print() # Empty line

# Group by Month and sum the values
otlrs_month_pagamentos = otlrs_2.groupby('Mês')['Valor da Movimentação'].sum()
print(otlrs_month_pagamentos)

# Filter the data without 'liers'
normal_clients = df_pagamentos[df_pagamentos['Eh_outlier'] == 'Não'].copy()

# Add columns for Year and Month
normal_clients['Ano'] = df_pagamentos['Data da Movimentação'].dt.year
normal_clients['Mês'] = df_pagamentos['Data da Movimentação'].dt.month

# Group by Year and sum the values
normal_clients_year = normal_clients.groupby('Ano')['Valor da Movimentação'].sum()
print(normal_clients_year)
print()

# Group by Month and sum the values
normal_clients_month = normal_clients.groupby('Mês')['Valor da Movimentação'].sum()
print(normal_clients_month)

# Exploring the data frame
dataset.keys()

# Exploring the customer database
df_cliente

# Checking for null values
df_cliente.isnull().sum()

# Description of the 'company name' column
df_cliente['Razao Social'].describe()

# Description of the Trade Name column
df_cliente['Nome Fantasia'].describe()

# Checking for duplicates
df_cliente['Nome Fantasia'].duplicated().sum()
print('Duplicados:', df_cliente['Nome Fantasia'].duplicated().sum())

# Filter duplicate rows in the 'Trade Name' column
duplicates = df_cliente[df_cliente.duplicated(subset=['Nome Fantasia'], keep=False)]

# Display duplicate values
print(duplicates['Nome Fantasia'].value_counts())
print(f'Total de duplicados: {len(duplicates)}')

# Remove duplicates
df_cliente.drop_duplicates(subset=['Nome Fantasia'], keep='first', inplace=True)
# Check if duplicates have been removed
print('Qtd de valores duplicados:', df_cliente['Nome Fantasia'].duplicated().sum())

# Checking the data frame
df_cliente

# Description of the 'Person Type' column
df_cliente['Tipo Pessoa'].describe()

# Counting values ​​in Person Type
df_cliente['Tipo Pessoa'].value_counts()

# Counting values ​​in Person Type (normalized)
df_cliente['Tipo Pessoa'].value_counts(normalize=True)  

# Bar Chart for Person Type
df_cliente['Tipo Pessoa'].value_counts().plot.bar(color='gray')

# Description of the 'Municipality' column
df_cliente['Municipio'].describe()

# Counting values ​​in Municipality
df_cliente['Municipio'].value_counts()

# Counting values ​​in Municipality (normalized)
df_cliente['Municipio'].value_counts(normalize=True)

# Bar chart for Municipality
df_cliente['Municipio'].value_counts().plot.bar(color='gray')

# UF column description
df_cliente['UF'].describe()

# Counting values ​​in UF
df_cliente['UF'].value_counts()

# Counting values ​​in UF (normalized)
df_cliente['UF'].value_counts(normalize=True)

# Bar chart for UF
df_cliente['UF'].value_counts().plot.bar(color='gray')

# Viewing the Supplier table
df_fornecedor

# Checking for null values
df_fornecedor.isnull().sum()

# Description of the 'company name' column
df_fornecedor['Razao Social'].describe()

# Description of the Trade Name column
df_fornecedor['Nome Fantasia'].describe()

# Checking for duplicates
df_fornecedor['Nome Fantasia'].duplicated().sum()

# Filter duplicate rows in the 'Trade Name' column
duplicates = df_fornecedor[df_fornecedor.duplicated(subset=['Nome Fantasia'], keep=False)]
# Display duplicate values
print(duplicates['Nome Fantasia'].value_counts())

# Remove duplicates
df_fornecedor.drop_duplicates(subset=['Nome Fantasia'], keep='first', inplace=True)
# Check if duplicates have been removed
print('Qtd de valores duplicados:', df_fornecedor['Nome Fantasia'].duplicated().sum())

# Description of the 'Person Type' column
df_fornecedor['Tipo Pessoa'].describe()

# Counting values ​​in Person Type
df_fornecedor['Tipo Pessoa'].value_counts()

# Counting values ​​in Person Type (normalized)
df_fornecedor['Tipo Pessoa'].value_counts(normalize=True)

# Bar Chart for Person Type
df_fornecedor['Tipo Pessoa'].value_counts().plot.bar(color='gray')

# Description of the 'Municipality' column
df_fornecedor['Municipio'].describe()

# Counting values ​​in Municipality
df_fornecedor['Municipio'].value_counts()

# Counting values ​​in Municipality (normalized)
df_fornecedor['Municipio'].value_counts(normalize=True)

# Bar chart for Municipality
df_fornecedor['Municipio'].value_counts().plot.bar(color='gray')

# UF column description
df_fornecedor['UF'].describe()

# Counting values ​​in UF
df_fornecedor['UF'].value_counts()

# Counting values ​​in UF (normalized)
df_fornecedor['UF'].value_counts(normalize=True)

# Bar chart for UF
df_fornecedor['UF'].value_counts().plot.bar(color='gray')

# Viewing the Bank table
df_banco

# Checking for null values
df_banco.isnull().sum()

# Removing the description of the bank name
df_banco['Nome Banco'] = df_banco['Nome Banco'].str.split(' - ').str[0]
# Checking the result
df_banco['Nome Banco'].describe()

# Counting values ​​in the Bank Name column
df_banco['Nome Banco'].value_counts()

# Counting values ​​in the Bank Name column (normalized)
df_banco['Nome Banco'].value_counts(normalize=True)

# Bank description 
df_banco['Município'].describe()

# Counting values ​​in Municipality
df_banco['Município'].value_counts()

# Counting values ​​in Municipality (normalized)
df_banco['Município'].value_counts(normalize=True)

# UF
df_banco['UF'].describe()

# Counting values ​​in UF
df_banco['UF'].value_counts()

# Counting values ​​in UF (normalized)
df_banco['UF'].value_counts(normalize=True)

# Definir o caminho do arquivo de backup
backup_path = '.\\BaseFinanceiro\\Financeiro_backup.xlsx'

# Criar uma cópia de segurança do arquivo original
shutil.copy(file_path, backup_path)
print(f"Cópia de segurança criada em: {backup_path}")

# Update the original file with the modified data
with pd.ExcelWriter(file_path) as writer:
    for sheet_name, df in dataset.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)
print("File updated successfully!")

# Option to restore the original file (uncomment to execute)

# shutil.copy(backup_path, file_path)
# print("Original file restored successfully!")

# # Verify the restored file
# df_restaurado = pd.read_excel(file_path)
# print(df_restaurado.head())


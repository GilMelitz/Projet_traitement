import pandas as pd

######################################################################################
#          EXTRAIRE LES COLONES SOUHAITEES A PARTIR D'UN FICHIER XLSX                #
######################################################################################
def extract_columns_excel(input_excel, output_excel, sheet_name, columns_to_extract):
    # Lit le fichier excel
    df = pd.read_excel(input_excel, sheet_name=sheet_name)

    # Extrait les colonnes spécifiés
    extracted_df = df[columns_to_extract]

    # Ecrit les colonnes dans un nouveau fichier
    extracted_df.to_excel(output_excel, index=False)

# Creation de l'intermediaire

input_excel = 'test.xlsx'  # Replacez avec notre fichier input excel
output_excel = 'inter.xlsx'  # Replacez avec notre fichier intermediaire excel
sheet_name = 'Book1'  # peut etre 'Sheet1'
columns_to_extract = ['vehicule']  # A remplacer avec les colonnes à extraire

extract_columns_excel(input_excel, output_excel, sheet_name, columns_to_extract)

######################################################################################
#          EXTRAIRE LES LIGNES SOUHAITEES A PARTIR D'UN FICHIER XLSX                #
######################################################################################

def extract_rows_by_multiple_values(input_excel, output_excel, sheet_name, column_name, values_to_filter):

    df = pd.read_excel(input_excel, sheet_name=sheet_name)

    # Filtre les lignes qui match les valeurs specifiees
    filtered_df = df[df[column_name].isin(values_to_filter)]

    filtered_df.to_excel(output_excel, index=False)

# Usage
input_excel = 'inter.xlsx'
output_excel = 'output_file.xlsx'
sheet_name = 'Sheet1'
column_name = 'vehicule'  # Replacez with the name of the column you want to filter by
values_to_filter = ['abguytvg', 'abc']  # Replace with the values you're looking for

extract_rows_by_multiple_values(input_excel, output_excel, sheet_name, column_name, values_to_filter)

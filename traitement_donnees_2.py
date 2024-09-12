import pandas as pd

# Usage pour exrtaire colonnes

input_excel = 'essai.xlsx'  # Replacez avec notre fichier input excel
output_excel = 'inter.xlsx'  # Replacez avec notre fichier intermediaire excel
sheet_name = 'Atal'  # peut etre 'Sheet1'
columns_to_extract = ['Véhicule', 'Immatriculation', 'Service daffectation', 'Date', 'Jour', 'heure', 'Compteur', 'Volume', 'Montant TTC']  # A remplacer avec les colonnes à extraire
# Racourci

tt_VL = ['VL', 'VL +', 'VL BLANC', 'VL ELECTRIQUE BLANC', 'VL HYBRIDE', 'VL POOL', 'VL POOL BLANC', 'VL_CdG', 'VL_CdG-HR', 'VL+', 'VL+ HYBRIDE', 'VL+ POOL', 'VLD', 'VLHR', 'VLOG NOVI', 'VLSSSM']
tt_VID = ['VID', 'VID POOL', 'VID POOL BLANC', 'VID XL', 'VIDCYNO']

# Usage pour extraire lignes
output_excel_2 = 'output_file.xlsx'
sheet_name_2 = 'Sheet1'
column_name = 'Véhicule'  # Remplacez avec le nom de la colonne filtre
values_to_filter = tt_VL + tt_VID  # Remplacez avec le nom des lignes à préserver




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

######################################################################################
#          EXTRAIRE LES LIGNES SOUHAITEES A PARTIR D'UN FICHIER XLSX                #
######################################################################################

def extract_rows_by_multiple_values(input_excel, output_excel, sheet_name, column_name, values_to_filter):

    df = pd.read_excel(input_excel, sheet_name=sheet_name)

    # Filtre les lignes qui match les valeurs specifiees
    filtered_df = df[df[column_name].isin(values_to_filter)]

    filtered_df.to_excel(output_excel, index=False)



######################################################################################
#                      LIGNES POUR FAIRE TOURNER LE PROGRAMME                        #
######################################################################################

extract_columns_excel(input_excel, output_excel, sheet_name, columns_to_extract)

extract_rows_by_multiple_values(output_excel, output_excel_2, sheet_name_2, column_name, values_to_filter)

def group_by_immatriculation(input_excel, output_excel, sheet_name, group_column):
    # Read the input Excel file
    df = pd.read_excel(input_excel, sheet_name=sheet_name)

    # Group by the Immatriculation column
    grouped_df = df.groupby(group_column).agg(list)

    # Write the grouped DataFrame to a new Excel file
    grouped_df.to_excel(output_excel)

# Example usage
input_excel = 'inter.xlsx'  # Intermediate file (or your file)
output_excel = 'grouped_by_immatriculation.xlsx'  # Output file
sheet_name = 'Sheet1'  # Sheet name
group_column = 'Immatriculation'  # Column to group by

group_by_immatriculation(input_excel, output_excel, sheet_name, group_column)

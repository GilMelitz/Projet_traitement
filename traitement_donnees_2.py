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

def group_and_aggregate_immatriculation(input_excel, output_excel, sheet_name, immatriculation_column, montant_column, date_column, volume_column, compteur_column, type_column, affectation_column):
    # Read the input Excel file
    df = pd.read_excel(input_excel, sheet_name=sheet_name)

    # Convert 'Date' column to datetime if not already
    df[date_column] = pd.to_datetime(df[date_column], errors='coerce')

    # Create a new column to check if the day is Saturday (5) or Sunday (6)
    df['IsWeekend'] = df[date_column].dt.dayofweek.isin([5, 6])

    # Group by the Immatriculation column and perform standard aggregations
    grouped_df = df.groupby(immatriculation_column).agg(
        Total_Montant_TTC=(montant_column, 'sum'),  # Sum Montant TTC
        Weekend_Count=('IsWeekend', 'sum'),  # Count weekends
        Total_Volume=(volume_column, 'sum'),  # Sum Volume
        First_Type_Vehicule=(type_column, 'first'),  # First Type de véhicule
        Unique_Affectations=(affectation_column, lambda x: list(x.unique()))  # List of unique affectations
    ).reset_index()

    # Now handle the Total Kilometers calculation separately
    # Filter out zero values from 'Compteur' column
    non_zero_compteur_df = df[df[compteur_column] != 0]

    # Calculate the total kilometers traveled (max - min) for each Immatriculation
    kilometers_df = non_zero_compteur_df.groupby(immatriculation_column).agg(
        Max_Compteur=(compteur_column, 'max'),
        Min_Compteur=(compteur_column, 'min')
    ).reset_index()

    # Calculate the total kilometers (Max - Min)
    kilometers_df['Total_Kilometers'] = kilometers_df['Max_Compteur'] - kilometers_df['Min_Compteur']

    # Merge the grouped_df with kilometers_df to include the Total_Kilometers
    final_df = pd.merge(grouped_df, kilometers_df[[immatriculation_column, 'Total_Kilometers']], on=immatriculation_column, how='left')

    # Add Consommation (en L/km) column, dividing Total_Volume by Total_Kilometers
    final_df['Consommation (en L/km)'] = final_df.apply(
        lambda row: row['Total_Volume'] / row['Total_Kilometers'] if row['Total_Kilometers'] > 0 else 0,
        axis=1
    )

    # Write the result to a new Excel file
    final_df.to_excel(output_excel, index=False)

# Example usage
input_excel = 'inter.xlsx'  # Intermediate file or your actual file
output_excel = 'grouped_immatriculation_with_all_aggregates.xlsx'  # Output file
sheet_name = 'Sheet1'  # Sheet name
immatriculation_column = 'Immatriculation'  # Column to group by
montant_column = 'Montant TTC'  # Column to sum for amounts
date_column = 'Date'  # Date column to check weekends
volume_column = 'Volume'  # Column to sum for volume
compteur_column = 'Compteur'  # Column to calculate total kilometers
type_column = 'Type de véhicule'  # Column for the first vehicle type
affectation_column = 'Service daffectation'  # Column for unique affectations

group_and_aggregate_immatriculation(input_excel, output_excel, sheet_name, immatriculation_column, montant_column, date_column, volume_column, compteur_column, type_column, affectation_column)

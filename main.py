import pandas as pd
import re


def dataframe_to_excel(dataframe_new):
    writer = pd.ExcelWriter('Gillette_Envios_Final.xlsx', engine='xlsxwriter')
    dataframe_new.to_excel(writer, 'Hoja1')
    writer.save()


if __name__ == '__main__':
    dataframe_base = pd.read_excel("Gillette_Base_Envios.xls")
    dataframe_codigos_postales = pd.read_excel("codigos-postales-de-mexico.xlsx")
    # Rename Table Columns
    dataframe_base.rename(columns={'Nombre - Lead Capture Data': 'Nombre_Cliente'}, inplace=True)
    dataframe_codigos_postales.rename(columns={'Código': 'Código_Postal'}, inplace=True)

    #Filters Table Columns
    regex_number = r"([0-9]{5})"
    dataframe_base['Código_Postal'] = dataframe_base['Direccion - Lead Capture Data'].str.extract(regex_number)
    dataframe_base.dropna(subset=['Código_Postal'], how='all', inplace=True)

    #Drop Table Columns
    dataframe_base.drop('Direccion - Lead Capture Data', axis=1, inplace=True)
    dataframe_base.drop('Ciudad ', axis=1, inplace=True)

    #Transform Colum Type
    dataframe_base['Código_Postal'] = dataframe_base['Código_Postal'].astype(int)


    dataframe_final = pd.merge(dataframe_base, dataframe_codigos_postales, on='Código_Postal')
    dataframe_final.drop_duplicates(subset='Nombre_Cliente', inplace=True)
    dataframe_to_excel(dataframe_final)





import pandas as pd
import sys

def filt_nan(fila: pd.Series):
    """
        Esta funcion elimina los espacios en blanco de una fila 

        Parametros:
        - `fila` - una fila del archivo excel

        Retorna:
        - devuelve la fila con los valores existentes
    """
    return fila[~pd.isna(fila)]


def get_individual_sites_by_row(site: str, row: pd.Series, direct_connected_sites: pd.DataFrame):
    """
        Esta funcion obtiene los recintos.lineas a partir de una fila dada como argumento y devuelve los recintos.lineas sin repetir.

        Parametros:
        - `site` - el recinto.linea que se quiere analizar
        - `row` - la fila correspondiente al recinto.linea que se quiere analizar
        - `direct_connected_sites` - basicamente todo el archivo excel 

        Retorna:
        - `pd.Series` con los valores recinto.linea que pertenecen a site
    """
    
    individual_sites = [site] # lista en la cual se almacenan los recintos.linea obtenidos

    for f in row: # iteramos por cada recinto.linea contenido en la fila
        individual_sites.append(f) # se agrega el recinto.linea a la lista

        new_row = filt_nan(direct_connected_sites.loc[f]) # se eliminan los espacios en blanco de la fila correspondiente al recinto.linea `f`
        
        if site in new_row.values: # verificamos si `site` ya exciste en la lista `individual_sites`, si existe se pasa al siguiente recinto.linea
            continue 

        if len(new_row) > 0: # si la fila filtrada no esta vacia aplicamos nuevamente la funcion `get_induvidual_sites_by_row` a cada recinto.linea existente 
            r = get_individual_sites_by_row(f, new_row, direct_connected_sites)
            individual_sites.extend(r) # agregamos los recintos.lineas obtenidos a la lista `individual_sites`
        else: # si la fila filtrada no contiiene ningun elemento se pasa al proximo recinto.linea
            continue
    return pd.unique(individual_sites[1:])


def get_individual_sites(direct_conected_sites: pd.DataFrame):
    """
        Esta funcion obtiene los recintos.lineas a partir de una fila dada como argumento y devuelve los recintos.lineas sin repetir.

        Parametros:
        - `direct_connected_sites` - basicamente todo el archivo excel 

        Retorna:
        - `pd.DataFrame` con todos los recintos.lineas conectados de forma indirecta a cada recinto.linea
    """
    individual_connected_sites = pd.DataFrame(index=direct_conected_sites.index, columns=[f'R{i+1}' for i in range(len(direct_conected_sites.index))]) # crea un dataframe con los mismos índices y columnas que rs_d

    for site in direct_conected_sites.index: # iteramos por todos los recintos.lineas
        row = filt_nan(direct_conected_sites.loc[site]) # ontenemos la fila correspondiente al recinto.linea `site`
        
        if len(row) == 0: # si la fila esta vacia pasamos al siguiente recinto.linea
            continue

        sites = get_individual_sites_by_row(site, row, direct_conected_sites) # obtenemos todos los recintos.linea conectados indirectamente a `site`

        individual_connected_sites.loc[site,'R1':f'R{len(sites)}'] = sites # guardamos estos valores en el DataFrame `individual_connected_sites`
    
    return individual_connected_sites


######################################################################################################
excel_file_name = 'INPUT.xlsx' # nombre del archivo excel

rs_d = pd.read_excel(excel_file_name, sheet_name='Hoja1', index_col=0) # carga solo la ventana 'Tabla con aportes directos' del excel.

rs_i = get_individual_sites(rs_d) # obtiene los recintos.lineas individuales para cada recinto.linea

with pd.ExcelWriter(excel_file_name, mode='a') as writer: 
    rs_i.to_excel(writer, sheet_name='Tabla con aportes individuales') # guarda los datos de los recintos individuales en el archivo excel


print('\nAhi tienes tus recintos.linea MOTHER FUCKER!!\n')


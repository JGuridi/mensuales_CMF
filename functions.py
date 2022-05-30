import enum
import sys
import requests
import dateutil.relativedelta
import pandas as pd
from datetime import date
from openpyxl import load_workbook
from openpyxl.utils.cell import rows_from_range


def fecha():
    fecha_mes_pasado = date.today() - dateutil.relativedelta.relativedelta(months=2)    # cambiar a 1
    año, mes = fecha_mes_pasado.year, fecha_mes_pasado.month
    dia = pd.Period(fecha_mes_pasado ,freq='M').end_time.date().day
    return año, mes, dia

def mensual_balance(año, mes, id_b, api_key):
    url = f'https://api.cmfchile.cl/api-sbifv3/recursos_api/balances/{año}/{mes}/instituciones/{id_b}?apikey={api_key}&formato=json'
    response = requests.get(url)
    # print(f'Balance {id_b}: {response}')
    return response.json()
 
def mensual_resultados(año, mes, id_b, api_key):
    url = f'https://api.cmfchile.cl/api-sbifv3/recursos_api/resultados/{año}/{mes}/instituciones/{id_b}?apikey={api_key}&formato=json'
    response = requests.get(url)
    # print(f'Resultado {id_b}: {response}')
    return response.json()

def u_efe(año, mes, dia, api_key):
    url = f'https://api.cmfchile.cl/api-sbifv3/recursos_api/uf/{año}/{mes}/dias/{dia}?apikey={api_key}&formato=json'
    response = requests.get(url)
    # print(f'UF: {response}')
    return response.json()

# def extraccion(eeff):
#     return [k[['CodigoCuenta','MonedaTotal']].rename(columns={'CodigoCuenta': 'cuenta', 'MonedaTotal': j}).sort_values(by=['cuenta']).set_index('cuenta') for j,k in eeff.items()]

def exel(banco):
    wb = load_workbook(filename=f'{banco}')
    ws_balance, ws_resultado = wb['B1'], wb['R1']
    return wb, ws_balance, ws_resultado

# def separar(ids, bj, rj):
#     cuentas_balance, montos_balance = bj['cuenta'].tolist(), bj[ids].tolist()
#     cuentas_resultado, montos_resultado = rj['cuenta'].tolist(), rj[ids].tolist()
#     return cuentas_balance, montos_balance, cuentas_resultado, montos_resultado

def pegar(eeff, cuentas, montos, filas):
    for v,r in zip(cuentas, range(2, filas)): # son 701 cuentas, el 703 no está incluido
        c = eeff.cell(row=r, column=1)
        c.value = v
    for v,r in zip(montos, range(2, filas)): # son 701 cuentas, el 703 no está incluido
        c = eeff.cell(row=r, column=2)
        c.value = v

def sumas(archivo, sheet_name, sum):
    df = pd.read_excel(archivo, sheet_name=sheet_name)
    orden = df.iloc[:, [3]]
    orden.columns = ['cuenta']
    orden['cuenta'] = orden['cuenta'].astype(str)
    todos = df.iloc[:, :2]          # sacamos las 2 primeras columnas (no incluye el 2 (3era columna))
    todos.columns = ['cuenta', 'monto']
    todos['cuenta'] = todos['cuenta'].astype(str)
    merged = orden.merge(todos, how='left', left_on='cuenta', right_on='cuenta').dropna(how='any')
    merged['monto'] = merged['monto'].str.replace(',', '.').astype(float)
    merged.set_index('cuenta', inplace=True)
    for m, n in enumerate(sum):
        merged.loc[f'suma {m + 1}', 'monto'] = merged.loc[n].sum()[0]
    return merged['monto'].tolist()

def pegar_2(ls, rng_p, dest):
    for idx, row in enumerate(rows_from_range(rng_p)):
        for cell in row:
            dest[cell].value = ls[idx]
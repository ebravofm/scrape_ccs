import requests
from bs4 import BeautifulSoup as bs
import os
from selenium import webdriver
import sys
from pyvirtualdisplay import Display
import datetime
from time import sleep
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait # available since 2.4.0
from selenium.webdriver.support import expected_conditions as EC # available since 2.26.0
import pandas as pd
import random
from gspread_pandas import Spread, Client
import gpandas as gpd
from xlrd import XLRDError
from creds import user, psswd

TABS = {'Principal': 'Curriculum Empresarial',
        'DashBoard': 'DashBoard',
        'DSegmentacion': 'Segmentacion',
        'DatosBasicosExterno': 'Datos Generales',
        'Diagnostico': 'Informe RS',
        'InformacionEstandarExterno': 'Inscripción y Servicios',
        'DocumentosDigitalizados': 'Documentos Acreditados',
        'ValorAgregado': 'Informes Especializados',
        'PreviewVitrina': 'Vitrina',
        'ComportamientoProveedor': 'Evaluación Comportamiento'}
GSHEET = '1PRYYRoSDazS_aMZ2He_2Okto73QKP4BEl4JwIX2O5dQ'

PLANILLA = gpd.gExcelFile(GSHEET)
CONTRATISTAS = PLANILLA.parse('Lista Contratistas')
S = Spread(user = 'ebravofm', spread = GSHEET, user_creds_or_client=None)


def browser(func):
    def browser_wrapper(*args, **kwargs):
        
        options = webdriver.ChromeOptions()
        '''if not os.path.exists('browser_data'):
                                    os.makedirs('browser_data')
                                options.add_argument("user-data-dir=browser_data")'''

        if sys.platform == 'linux':
            display = Display(visible=0, size=(800, 600))
            display.start()
            options.add_argument(f"download.default_directory={os.getcwd()}")
            options.add_argument('--no-sandbox')

        d = webdriver.Chrome(options=options)

        
        print('[+] Logged in.')

        try:
            result = func(d, *args, **kwargs)
            
        except Exception as exc:
            d.close()
            if sys.platform == 'linux':
                display.stop()
            raise RuntimeError(str(exc))


        '''d.close()
        if sys.platform == 'linux':
            display.stop()
        print('[+] Logged Out.')'''

        return result
    
    return browser_wrapper


@browser
def ccs_login(d):
    d.get('https://www.rednegociosccs.cl/WebIngresoRPE/Login.aspx')
    
    user_field = d.find_element_by_id('txtCuentaUsuarioMandante')
    psswd_field = d.find_element_by_id('txtPasswordMandante')

    user_field.send_keys(user)
    sleep(1)
    psswd_field.send_keys(psswd)
    sleep(1)
    psswd_field.send_keys(Keys.RETURN)
    
    return d


def extract_contrator_list():
    # Get Page
    d = ccs_login()
    index = 'https://www.rednegociosccs.cl/WebPrivadoMandanteRPE/ConsultarFichaFull/Default.aspx'
    d.get(index)
    Buscar = d.find_element_by_id('UserControlBuscador1_lnkBuscarSegmento')
    Buscar.click()
    WebDriverWait(d, 10).until(EC.presence_of_element_located((By.ID, "UserControlBuscador1_dgrListaBusqueda")))
    print(1)

    # Parse table
    html = d.page_source
    df = pd.read_html(html, attrs={'id': 'UserControlBuscador1_dgrListaBusqueda'})[0]
    df.columns = df.iloc[0]
    df = df[1:]
    df['Link'] = 'https://www.rednegociosccs.cl/WebPrivadoMandanteRPE/ConsultarFichaFull/Principal.aspx?proveedor=' + df['Rut']
    
    d.close()
    
    S.df_to_sheet(df, index=True, replace=True, sheet='Lista Contratistas')

    return df


def scrape_contractors(rut_list):
    tprint(f'[·] Logging in...')
    d = ccs_login()
    n = 1
    for rut in rut_list:
        print()
        tprint(f'[·] Contratista {n}/{len(rut_list)}: {rut}...')
        try:

            link = 'https://www.rednegociosccs.cl/WebPrivadoMandanteRPE/ConsultarFichaFull/Principal.aspx?proveedor=' + rut
            d.get(link)
            sleep(2)

            html = d.page_source
            soup = bs(html, 'lxml').find('ul', id='tabnav')
            tabs = ['https://www.rednegociosccs.cl/WebPrivadoMandanteRPE/ConsultarFichaFull/'+t['href'] for t in soup.find_all('a')]

            for tab in tabs:
                try:
                    tab_ = tab.split('/')[-1].split('.')[0]

                    tprint(f'[·] Working tab {tab_}')
                    d.get(tab)
                    try:
                        alert = d.switch_to.alert
                        alert.accept()
                        #tprint('Alert Accepted')
                    except:
                        pass
                    tprint(f'[·] Extracting Values...')
                    values = extract_values_from_html(d.page_source)
                    append_to_sheet(values, rut, TABS[tab_])
                    tprint(f'[+] Appended to sheet')
                except Exception as exc:
                    tprint(f'[-] Failed on Tab level ({tab_}): {str(exc)[:200]}')
                    
                
        except Exception as exc:
            tprint(f'[-] Failed on Contractor level ({rut}): {str(exc)[:200]}')
        n += 1
        
    d.close()


def append_to_sheet(values, rut, tab):
    df = CONTRATISTAS
    nombre = df['Razon Social'][df.Rut==rut].iloc[0]
    
    try:
        current_sheet = gpd.read_gexcel(GSHEET, sheet_name=tab)
        current_sheet = current_sheet[~(current_sheet['1. Rut']==rut)]
    except XLRDError:
        current_sheet = pd.DataFrame()
        
    values = cleanse_values(values)
    
    row = pd.DataFrame.from_dict(values, orient='index').T
    row.insert(0, '0. Nombre', nombre)
    row.insert(0, '1. Rut', rut)
    
    current_sheet = current_sheet.append(row)
    
    print(tab)
    S.df_to_sheet(current_sheet, index=False, replace=True, sheet=tab)


def cleanse_values(values):
    for key in values.keys():        
        new_key = strip_field_names(key)
        values[new_key] = values.pop(key)

    return values


def strip_field_names(s):
    for prefix in ['wucPrincipal_', 'wucDashBoard_', 'wucSegmentacion_', 'wucDatosBasicos_', 'WucCertificados1_', 'WucDocumentosDigitalizados_', 'WucDatosProveedor1_', 'dgrTotalEvalProv_', 'lbl', 'Lbl']:
        s = s.replace(prefix, '')
    return s


def extract_values_from_html(html):
    values = {}
    soup = bs(html, 'lxml').find('div', id='mainficha')

    tags = ['span', 'a']
    ids = []
    for tag in tags:
        ids += soup.find_all(tag, id=True)
    for i in ids:
        values[i['id']] = i.text
    
    return values
    
    
def tprint(*args):        
    stamp = '[{:%d/%m-%H:%M}]'.format(datetime.datetime.now())
    print(stamp, *args)

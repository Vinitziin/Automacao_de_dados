import requests
import numpy as np
import pandas as pd
from datetime import timedelta,date, datetime
import time
import warnings
import urllib3
import os
import io
import locale
import smtplib
from config import url_arq_mesa, email_1, senha, db, db_connection, basicURL,verifyCertificate,apikey_AEScomercializadora,email_2, password, company_code, destinatarios
from PIL import ImageGrab
import win32com.client as client 
pd.options.mode.chained_assignment = None
warnings.simplefilter(action='ignore', category=UserWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
import pythoncom
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File 
import pandas as pd
pd.set_option('display.max_columns', 500)
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from pathlib import Path
import chromedriver_autoinstaller
from workalendar.america import Brazil
calend = Brazil()

chromedriver_autoinstaller.install()

import mysql_functions
mysql_func = mysql_functions.mysql_func()

credenciais = pd.read_csv(r'C:\acesso_BD\credenciais.csv')
user_banco = credenciais.login.item()
senha_banco= str(credenciais.senha.item())


def marcacao_mesa(user_banco,senha_banco,data_hoje):

    print("Lendo Book Mesa...")
    print()
    # ----- BOOK MESA -----
    url_arq_mesa
    email_1 
    senha 

    #Obtendo Book Mesa
    ctx_auth_arq_mesa = AuthenticationContext(url_arq_mesa)
    if ctx_auth_arq_mesa.acquire_token_for_user(email_1, senha):
        ctx_arq_mesa = ClientContext(url_arq_mesa, ctx_auth_arq_mesa)
        web = ctx_arq_mesa.web
        ctx_arq_mesa.load(web)
        ctx_arq_mesa.execute_query()
        #print("Authentication successful")
    response = File.open_binary(ctx_arq_mesa, url_arq_mesa)
    bytes_file_mesa = io.BytesIO()
    bytes_file_mesa.write(response.content)
    bytes_file_mesa.seek(0)

    arq_mesa = pd.read_excel(bytes_file_mesa, engine='openpyxl',sheet_name='Preços - Mensal')

    df_mesa = arq_mesa.copy()
    df_mesa.columns = df_mesa.loc[3]
    df_mesa = df_mesa.drop([0,1,2,3]).reset_index(drop = True)
    df_mesa = df_mesa[df_mesa.columns[:24]]
    df_mesa['Mês'] = pd.to_datetime(df_mesa['Mês'])
    df_mesa['Ano'],df_mesa['Mes'] = df_mesa['Mês'].dt.year,df_mesa['Mês'].dt.month

    df_mesa.reset_index(drop = True,inplace=True)

    book_mesa = pd.DataFrame(columns=['Data','conv D', 'conv D-1','i5','S','NE','N'])
    book_mesa['Data'] = df_mesa['Mês']
    book_mesa['conv D'] = df_mesa['Px (0)']
    book_mesa['conv D-1'] = df_mesa['Px (-1)']
    book_mesa.S = df_mesa.S + book_mesa['conv D']
    book_mesa.NE = df_mesa.NE + book_mesa['conv D']
    book_mesa.N = df_mesa.N + book_mesa['conv D']

    spread_i5_SE = df_mesa['I5'].T
    spread_i5_SE.reset_index(drop = True,inplace=True)
    spread_i5_SE = list(spread_i5_SE.loc[0])

    book_mesa['i5'] = spread_i5_SE
    book_mesa.Data = pd.to_datetime(book_mesa.Data)

    #---------------MATURIDADES---------------
    # Calcular a maturidade em anos, meses e trimestres para cada data
    locale.setlocale(locale.LC_TIME, 'pt_BR.utf-8')

    # Calcular a diferença em anos, meses e trimestres considerando o ano base como 2024
    book_mesa['ANU'] = (book_mesa['Data'].dt.year - data_hoje.year)
    book_mesa['M'] = ((book_mesa['Data'].dt.year - data_hoje.year) * 12 + (book_mesa['Data'].dt.month - data_hoje.month))
    book_mesa['TRI'] = (book_mesa['Data'].dt.year - data_hoje.year) * 4 + ((book_mesa['Data'].dt.month - 1) // 3) - ((data_hoje.month - 1) // 3)
    book_mesa['SEM'] = (book_mesa['Data'].dt.year - data_hoje.year) * 2 + (book_mesa['Data'].dt.month // 7 - data_hoje.month // 7)
    # Corrigir a maturidade em trimestres para datas no mesmo trimestre
    same_quarter = (book_mesa['Data'].dt.year == data_hoje.year) & ((book_mesa['Data'].dt.month - 1) // 3 == (data_hoje.month - 1) // 3)
    book_mesa.loc[same_quarter, 'TRI'] = 0

    # Adicionar prefixo e sufixo às colunas 'A', 'M' e 'T'
    book_mesa['ANU'] = 'ANU' + book_mesa['ANU'].astype(str) 
    book_mesa['M'] = 'M' + book_mesa['M'].astype(str) 
    book_mesa['TRI'] = 'TRI' + book_mesa['TRI'].astype(str) 
    book_mesa['SEM'] = 'SEM' + book_mesa['SEM'].astype(str)

        #---------------PREÇOS---------------

    maturidades = ['M-1','M0', 'M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'M10', 'M11','M12'
    ,'TRI0', 'TRI1', 'TRI2', 'TRI3', 'TRI4', 'TRI5'
    ,'SEM0','SEM1','SEM2','SEM3', 'SEM4'
    ,'ANU0','ANU1','ANU2','ANU3','ANU4','ANU5','ANU6']

    PREÇOS = pd.DataFrame(columns=['Data', 'Produto', 'Maturidade','Preço', 'Preço Anterior', 'Spread', 'I5', 'S','NE','N'])
    PREÇOS['Maturidade'] = maturidades
    PREÇOS['Data'] = data_hoje.strftime('%d-%m-%Y')

    preco,preco_anterior, i5,S,NE,N, produto = [], [], [], [],[],[],[]
    for maturidade in maturidades:
        if 'M' in maturidade and maturidade in book_mesa['M'].values:
            preco.append(book_mesa.loc[book_mesa.M == maturidade, 'conv D'].mean())
            preco_anterior.append(book_mesa.loc[book_mesa.M == maturidade, 'conv D-1'].mean())
            i5.append(book_mesa.loc[book_mesa.M == maturidade, 'i5'].mean())
            S.append(book_mesa.loc[book_mesa.M == maturidade, 'S'].mean())
            NE.append(book_mesa.loc[book_mesa.M == maturidade, 'NE'].mean())
            N.append(book_mesa.loc[book_mesa.M == maturidade, 'N'].mean())
            produto.append('SE CON MEN ' + book_mesa.loc[book_mesa.M == maturidade, 'Data'].iloc[0].strftime("%b/%y").upper() + ' - Preço Fixo')

        elif 'TRI' in maturidade and maturidade in book_mesa['TRI'].values:
            preco.append(round(book_mesa.loc[book_mesa.TRI == maturidade, 'conv D'].mean()))
            preco_anterior.append(round(book_mesa.loc[book_mesa.TRI == maturidade, 'conv D-1'].mean()))
            i5.append(book_mesa.loc[book_mesa.TRI == maturidade, 'i5'].mean())
            S.append(book_mesa.loc[book_mesa.TRI == maturidade, 'S'].mean())
            NE.append(book_mesa.loc[book_mesa.TRI == maturidade, 'NE'].mean())
            N.append(book_mesa.loc[book_mesa.TRI == maturidade, 'N'].mean())
            produto.append('SE CON TRI ' + (data := book_mesa.loc[book_mesa.TRI == maturidade, 'Data'].iloc[0]).strftime("%b/%y").upper() + ' ' + min(data + pd.DateOffset(months=2), data + pd.tseries.offsets.YearEnd()).strftime("%b/%y").upper() + ' - Preço Fixo')

        elif 'SEM' in maturidade and maturidade in book_mesa['SEM'].values:
            preco.append(round(book_mesa.loc[book_mesa.SEM == maturidade, 'conv D'].mean()))
            preco_anterior.append(round(book_mesa.loc[book_mesa.SEM == maturidade, 'conv D-1'].mean()))
            i5.append(book_mesa.loc[book_mesa.SEM == maturidade, 'i5'].mean())
            S.append(book_mesa.loc[book_mesa.SEM == maturidade, 'S'].mean())
            NE.append(book_mesa.loc[book_mesa.SEM == maturidade, 'NE'].mean())
            N.append(book_mesa.loc[book_mesa.SEM == maturidade, 'N'].mean())
            produto.append('SE CON SEM ' + (data := book_mesa.loc[book_mesa.SEM == maturidade, 'Data'].iloc[0]).strftime("%b/%y").upper() + ' ' + min(data + pd.DateOffset(months=5), data + pd.tseries.offsets.YearEnd()).strftime("%b/%y").upper() + ' - Preço Fixo')
            
        elif 'ANU' in maturidade and maturidade in book_mesa['ANU'].values:
            preco.append(round(book_mesa.loc[book_mesa.ANU == maturidade, 'conv D'].mean()))
            preco_anterior.append(round(book_mesa.loc[book_mesa.ANU == maturidade, 'conv D-1'].mean()))
            i5.append(book_mesa.loc[book_mesa.ANU == maturidade, 'i5'].mean())
            S.append(book_mesa.loc[book_mesa.ANU == maturidade, 'S'].mean())
            NE.append(book_mesa.loc[book_mesa.ANU == maturidade, 'NE'].mean())
            N.append(book_mesa.loc[book_mesa.ANU == maturidade, 'N'].mean())
            produto.append('SE CON ANU '+ (data := book_mesa.loc[book_mesa.ANU == maturidade, 'Data'].iloc[0]).strftime("%b/%y").upper() + ' ' + min(data + pd.DateOffset(months=11), data + pd.tseries.offsets.YearEnd()).strftime("%b/%y").upper() + ' - Preço Fixo')

    PREÇOS['Preço'] = preco
    PREÇOS['Preço Anterior'] = preco_anterior
    PREÇOS['I5'] = i5
    PREÇOS['S'] = S
    PREÇOS['NE'] = NE
    PREÇOS['N'] = N
    PREÇOS['Spread'] = PREÇOS['I5'] - PREÇOS['Preço']
    PREÇOS['Produto'] = produto

    precos = PREÇOS.rename(columns={'Preço':'Preco','Preço Anterior':'Preco_anterior'})
    precos = precos[['Data','Produto','Maturidade','Preco','Preco_anterior','Spread','I5']]

    db 
    db_connection 
    mysql_func.insert(db_connection,precos, 'historico_precos')

    i5_anterior = []

    aux = 1
    precos_anteriores = mysql_func.read_query_table(db_connection,'historico_precos',f"""SELECT * FROM bd_mesa.historico_precos WHERE Data = '{(data_hoje - timedelta(aux)).strftime('%d-%m-%Y')}'""")
    while len(precos_anteriores) == 0:
        aux = aux+1
        precos_anteriores = mysql_func.read_query_table(db_connection,'historico_precos',f"""SELECT * FROM bd_mesa.historico_precos WHERE Data = '{(data_hoje - timedelta(aux)).strftime('%d-%m-%Y')}'""")

    for mat in PREÇOS.Maturidade:
        i5_anterior.append(precos_anteriores.loc[precos_anteriores.Maturidade == mat,'I5'].item())

    PREÇOS['I5 Anterior'] = i5_anterior
    PREÇOS['Spread Anterior'] =   PREÇOS['I5 Anterior'] -  PREÇOS['Preço Anterior']
    PREÇOS = PREÇOS.round(2)
    
    #---------------- Coletando o Token ----------------#
    print("\nBuscando negociações na BBCE...")
    print('=====================================')
    basicURL
    verifyCertificate 
    apikey_AEScomercializadora
    email_2
    password
    company_code

    headers = {}
    headers["Accept"] = "application/json"
    headers["Content-Type"] = "application/json"
    headers["apiKey"] = apikey_AEScomercializadora

    data = {
        'companyExternalCode': company_code,
        'email': email_2,
        'password':password
    }

    url=basicURL+ '/v2/login'
    tokenResponse = requests.post(url, headers=headers, json=data,
                                verify=verifyCertificate)
    token=tokenResponse.json()['idToken']

    #---------------- Coletando as Negociações ----------------#
    headers = {
        'Authorization': 'Bearer ' + token,
        "Accept": "application/json"
    }
    headers["apiKey"] = apikey_AEScomercializadora

    url=basicURL+ '/v1/all-deals/report'
    
    dia_ini = calend.add_working_days(data_hoje, -30).strftime('%Y-%m-%d')
    dia_final = data_hoje.strftime('%Y-%m-%d')

    query={}
    query['initialPeriod'] = dia_ini
    query['finalPeriod'] = dia_final

    response = requests.get(url, headers=headers,params=query, verify=verifyCertificate)
    negocios=response.json()
    dfDeals=pd.DataFrame(negocios)

    if len(dfDeals) == 0:
        print("Não há negociações no intervalo determinado...")
    else:
        dates = []
        for i in range(len(dfDeals)):
            date_i = datetime.fromisoformat(dfDeals['createdAt'][i])
            dates.append(datetime(date_i.year,date_i.month,date_i.day,date_i.hour,date_i.minute,date_i.second).strftime('%Y-%m-%d %H:%M:%S'))
        dfDeals['createdAt'] = dates
        dfDeals['createdAt'] = pd.to_datetime(dfDeals['createdAt'])
        dfDeals = dfDeals.sort_values(by=['createdAt'], ascending=True).reset_index(drop = True)

        #---------------- Adiconando as Descrições dos Produtos ----------------#
        dfNegociacoes=dfDeals.copy()
        dfNegociacoes['description'] = 0
        produtos = mysql_func.read_query_table(db_connection,'produtos_bbce',f"""SELECT * FROM bd_mesa.produtos_bbce """)

        for i in range(len(dfNegociacoes)):
            if dfNegociacoes['description'][i]==0:
                try:
                    try:
                        desc = produtos.loc[produtos.productId == dfNegociacoes['productId'][i],'Produto'].to_list()[0]
                        dfNegociacoes['description'][i] = desc
                    except:
                        url=basicURL+ f"/v2/tickers/{dfNegociacoes['productId'][i]}"
                        response = requests.get(url, headers=headers, verify=verifyCertificate)
                        productDescription=response.json()
                        dfNegociacoes['description'][i] = productDescription['description']
                        novo_produto = [dfNegociacoes['productId'][i],productDescription['description']]
                        linha_nova = pd.DataFrame.from_dict({a: [b] for a,b in zip(list(produtos.columns),novo_produto)})
                        mysql_func.insert(db_connection,linha_nova, 'produtos_bbce')

                        produtos.loc[len(produtos)] = novo_produto
                        produtos.drop_duplicates(inplace=True)
                except:
                    dfNegociacoes = dfNegociacoes.drop(i)
        
        
        dfNegociacoes = dfNegociacoes.rename(columns={'createdAt': 'Data', 'description': 'Produto', 'originOperationType': 'Tipo_Contrato', 'quantity': 'MWm','quantityMeasured': 'MWh', 'unitPrice': 'Preco', 'status': 'Cancelado'})

        dfNegociacoes['Tipo_Contrato'] = dfNegociacoes['Tipo_Contrato'].replace(['Registro','Match'],['Boleta','Balcão'])
        dfNegociacoes['Cancelado'] = dfNegociacoes['Cancelado'].replace(['Cancelado','Ativo'],['Sim','Não'])

        df_clone=dfNegociacoes.copy()
        dfNegociacoes_final=df_clone[['Data','Produto', 'Tipo_Contrato', 'MWm', 'MWh', 'Preco', 'Cancelado']]
        dfNegociacoes_final = dfNegociacoes_final.loc[dfNegociacoes_final.Cancelado == 'Não']

    #---------- VOLUMES 30 DIAS ----------
    print("Apurando Volume...")
    print('=====================================')

    maturidades_volume = ['M-1', 'M-1 I5','M0', 'M0 I5' , 'M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'M10', 'M11','M12'
    ,'TRI0', 'TRI1', 'TRI2', 'TRI3', 'TRI4', 'TRI5'
    ,'SEM0','SEM1','SEM2','SEM3','SEM4'
    ,'ANU0','ANU1','ANU1 I5','ANU2','ANU2 I5','ANU3','ANU4','ANU5'
    ] 
    volumes_final = pd.DataFrame({'Maturidade': maturidades_volume})

    dfNegociacoes_final['data'] = pd.to_datetime(dfNegociacoes_final.Data).dt.date
    negociacoes_hoje = dfNegociacoes_final.loc[dfNegociacoes_final.data == data_hoje]

    contador = 0
    for j in range(30):

        dia = data_hoje - timedelta(j+contador)
        while not calend.is_working_day(dia):
            contador +=1
            dia = data_hoje - timedelta(j+contador)

        x = dfNegociacoes_final.loc[dfNegociacoes_final.data == dia]
        x.reset_index(drop = True, inplace = True)

        #prods = dfNegociacoes_final.Produto.str.split(' - ',expand=True)[0]
        #x.loc[x.Produto.str.contains(prods[0]),'MWm'].sum()

        produtos = x['Produto'].unique().tolist()
        produtos = [x.split(' - ')[0] for x in produtos]

        dict_prod = {}
        for prod in produtos:
            volume = x.loc[x['Produto'].str.startswith(prod), 'MWm'].sum()
            dict_prod[prod] = volume

        volumes = pd.DataFrame().assign(maturidade=maturidades_volume,produto='',volume=0)

        for i in range(len(volumes)):
            maturidade = volumes.loc[i, 'maturidade']
            if ' I5' in maturidade:
                maturidade = maturidade.replace(' I5', '')
                prefixo = 'SE I5'
            else:
                prefixo = 'SE CON'
                
            if 'M' in maturidade and maturidade in book_mesa['M'].values:
                volumes.loc[i, 'produto'] = prefixo + ' MEN ' + book_mesa.loc[book_mesa.M == maturidade, 'Data'].iloc[0].strftime("%b/%y").upper() 
            elif 'TRI' in maturidade and maturidade in book_mesa['TRI'].values:
                volumes.loc[i, 'produto'] = prefixo + ' TRI ' + (data := book_mesa.loc[book_mesa.TRI == maturidade, 'Data'].iloc[0]).strftime("%b/%y").upper() + ' ' + min(data + pd.DateOffset(months=2), data + pd.tseries.offsets.YearEnd()).strftime("%b/%y").upper() 
            elif 'SEM' in maturidade and maturidade in book_mesa['SEM'].values:
                volumes.loc[i, 'produto'] = prefixo + ' SEM ' + (data := book_mesa.loc[book_mesa.SEM == maturidade, 'Data'].iloc[0]).strftime("%b/%y").upper() + ' ' + min(data + pd.DateOffset(months=6), data + pd.tseries.offsets.YearEnd()).strftime("%b/%y").upper() 
            elif 'ANU' in maturidade and maturidade in book_mesa['ANU'].values:
                volumes.loc[i, 'produto'] = prefixo + ' ANU '+ (data := book_mesa.loc[book_mesa.ANU == maturidade, 'Data'].iloc[0]).strftime("%b/%y").upper() + ' ' + min(data + pd.DateOffset(months=12), data + pd.tseries.offsets.YearEnd()).strftime("%b/%y").upper() 
                
        for i in range(len(volumes)):
            for prod in list(dict_prod.keys()):
                if volumes.loc[i, 'produto'] in prod:
                    volumes.loc[i, 'volume'] = dict_prod[prod]
        
        volumes_final['D-' + str(j)] = volumes.volume.tolist() 
    volumes_final['produto'] = volumes.produto.tolist()

    soma = volumes_final.copy()
    for i in range(len(soma)):
        if "TRI" in soma.loc[i, 'Maturidade']:
            for col in soma.columns.tolist():
                if (col != 'Maturidade') & (col != 'produto'):
                    soma.loc[i, col] = soma.loc[i, col] * 3 

        if "SEM" in soma.loc[i, 'Maturidade']:
            for col in soma.columns.tolist():
                if (col != 'Maturidade') & (col != 'produto'):
                    soma.loc[i, col] = soma.loc[i, col] * 6
        
        if "ANU" in soma.loc[i, 'Maturidade']:
            for col in soma.columns.tolist():
                if (col != 'Maturidade') & (col != 'produto'):
                    soma.loc[i, col] = soma.loc[i, col] * 12
            

    volumes_final.loc[len(volumes_final), 'Maturidade'] = "Total" 
    for col in soma.columns.to_list():
        if (col != 'Maturidade') & (col != 'produto'):
            volumes_final.loc[len(volumes_final) - 1, col] = soma[col].sum()

    volumes_final = volumes_final.round(2)
    #---------- VOLUMES HOJE ----------

    for i in range(len(volumes_final)):
        for col in volumes_final.columns:
            if pd.isna(volumes_final.loc[i, col]):
                volumes_final.loc[i, col] = ''

    volumes_final_peq=pd.DataFrame(columns = ['Data','Maturidade','Produto','Volume','Operacoes'])

    volumes_final_peq.Produto = volumes_final.produto
    volumes_final_peq.Maturidade = volumes_final.Maturidade
    volumes_final_peq.Volume = volumes_final['D-0']
    volumes_final_peq.Data = data_hoje

    operacoes = []
    for i in range(len(volumes_final_peq)):
        operacoes.append(negociacoes_hoje.loc[negociacoes_hoje.Produto.str.contains(volumes_final_peq.Produto[i])].shape[0])
    volumes_final_peq.Operacoes = operacoes
    def categorizar_liquidez(liquidez_total):
        if liquidez_total >= 100:
            return 'Very High'
        elif liquidez_total >= 75:
            return 'High'
        elif liquidez_total >= 50:
            return 'Normal'
        elif liquidez_total >= 25:
            return 'Low'
        else:
            return 'Very Low'

    volumes_final_peq['Liquidez'] = '' 
    for maturidade in volumes_final.Maturidade.tolist():
            media = volumes_final.loc[volumes_final.Maturidade == maturidade, 'D-1':'D-14' ].values.mean()
            liquidez = volumes_final.loc[volumes_final.Maturidade == maturidade, 'D-0'] / media
            liquidez = (liquidez * 100).round()
            volumes_final_peq.loc[volumes_final_peq.Maturidade == maturidade, 'Liquidez'] = liquidez.apply(categorizar_liquidez)
    
    curva_teste = pd.DataFrame(columns=['MATURIDADE','EMPRESA', 'TICKER', 'VALOR (R$)'])

    curva_teste['MATURIDADE'] = ['M0', 'M1', 'M2', 'M3', 'TRI2', 'TRI3', 'ANU1', 'ANU2', 'ANU3', 'ANU4', 'ANU5', 'ANU6',
                                'M0 I5', 'M1 I5', 'M2 I5', 'M3 I5','TRI2 I5', 'TRI3 I5', 'ANU1 I5', 'ANU2 I5', 'ANU3 I5', 'ANU4 I5', 'ANU5 I5', 'ANU6 I5',
                                'M0 NO', 'M1 NO', 'M2 NO', 'M3 NO', 'TRI2 NO', 'TRI3 NO', 'ANU1 NO', 'ANU2 NO', 'ANU3 NO', 'ANU4 NO', 'ANU5 NO', 'ANU6 NO',
                                'M0 NE', 'M1 NE', 'M2 NE', 'M3 NE', 'TRI2 NE', 'TRI3 NE', 'ANU1 NE', 'ANU2 NE', 'ANU3 NE', 'ANU4 NE', 'ANU5 NE', 'ANU6 NE',
                                'M0 SU', 'M1 SU', 'M2 SU', 'M3 SU', 'TRI2 SU', 'TRI3 SU', 'ANU1 SU', 'ANU2 SU', 'ANU3 SU', 'ANU4 SU', 'ANU5 SU', 'ANU6 SU']

    for i in range(len(curva_teste)):
        if 'I5' in curva_teste.loc[i, 'MATURIDADE']:
            mat = curva_teste.loc[i, 'MATURIDADE'].replace(' I5', '')
            curva_teste.loc[i, 'TICKER'] = PREÇOS.loc[PREÇOS['Maturidade'] == mat, 'Produto'].values[0].replace('CON', 'I5')
            curva_teste.loc[i, 'VALOR (R$)'] = PREÇOS.loc[PREÇOS['Maturidade'] == mat, 'I5'].values[0]
        elif 'NO' in curva_teste.loc[i, 'MATURIDADE']:
            mat = curva_teste.loc[i, 'MATURIDADE'].replace(' NO', '')
            curva_teste.loc[i, 'TICKER'] = PREÇOS.loc[PREÇOS['Maturidade'] == mat, 'Produto'].values[0].replace('SE ', 'NO ')
            curva_teste.loc[i, 'VALOR (R$)'] = PREÇOS.loc[PREÇOS['Maturidade'] == mat, 'N'].values[0]
        elif 'NE' in curva_teste.loc[i, 'MATURIDADE']:
            mat = curva_teste.loc[i, 'MATURIDADE'].replace(' NE', '')
            curva_teste.loc[i, 'TICKER'] = PREÇOS.loc[PREÇOS['Maturidade'] == mat, 'Produto'].values[0].replace('SE ', 'NE ')
            curva_teste.loc[i, 'VALOR (R$)'] = PREÇOS.loc[PREÇOS['Maturidade'] == mat, 'NE'].values[0]
        elif 'SU' in curva_teste.loc[i, 'MATURIDADE']:
            mat = curva_teste.loc[i, 'MATURIDADE'].replace(' SU', '')
            curva_teste.loc[i, 'TICKER'] = PREÇOS.loc[PREÇOS['Maturidade'] == mat, 'Produto'].values[0].replace('SE ', 'SU ')
            curva_teste.loc[i, 'VALOR (R$)'] = PREÇOS.loc[PREÇOS['Maturidade'] == mat, 'S'].values[0]
        else:
            curva_teste.loc[i, 'TICKER'] = PREÇOS.loc[PREÇOS['Maturidade'] == curva_teste.loc[i, 'MATURIDADE'], 'Produto'].iloc[0]
            curva_teste.loc[i, 'VALOR (R$)'] = PREÇOS.loc[PREÇOS['Maturidade'] == curva_teste.loc[i, 'MATURIDADE'], 'Preço'].iloc[0]

        curva_teste.loc[i,'EMPRESA'] = 'AES COMERCIALIZADORA DE ENERGIA LTDA'

    curva_fwd = pd.DataFrame(columns=['EMPRESA', 'TICKER', 'VALOR (R$)'])

    curva_fwd['EMPRESA'] = curva_teste.EMPRESA
    curva_fwd['TICKER'] = curva_teste.TICKER
    curva_fwd['VALOR (R$)'] = curva_teste['VALOR (R$)']

    ###################################################################################################################################
    print("Subindo a Curva_fwd...")
    print('=====================================')

    headers = {
        'Authorization': 'Bearer ' + token,
        "Accept": "application/json"
    }
    headers["apiKey"] = apikey_AEScomercializadora

    url=basicURL+ '/v1/curve/call'

    response = requests.get(url, headers=headers, verify=verifyCertificate)
    products = pd.DataFrame(response.json())
    products['description'] = 0


    produtos = mysql_func.read_query_table(db_connection,'produtos_bbce',f"""SELECT * FROM bd_mesa.produtos_bbce """)

    for i in range(len(products)):
        if not produtos.loc[produtos.productId == products.loc[i,'tickerId'],'Produto'].empty:
            products.loc[i,'description'] = produtos.loc[produtos.productId == products.loc[i,'tickerId'],'Produto'].item()
        else:
            url=basicURL+ f"/v2/tickers/{products.loc[i,'tickerId']}"
            response = requests.get(url, headers=headers, verify=verifyCertificate)
            productDescription=response.json()
            products.loc[i,'description'] = productDescription['description']
            novo_produto = [products.loc[i,'tickerId'],productDescription['description']]
            linha_nova = pd.DataFrame.from_dict({a: [b] for a,b in zip(list(produtos.columns),novo_produto)})
            mysql_func.insert(db_connection,linha_nova, 'produtos_bbce')

            produtos.loc[len(produtos)] = novo_produto
            produtos.drop_duplicates(inplace=True)

    products_list = []

    for i, row in products.iterrows():
        description = row['description']

        if 'I5' in description:
            modified_description = description.replace('I5', 'CON')
            value = PREÇOS.loc[PREÇOS['Produto'] == modified_description, 'I5'].values

        elif 'NO' in description:
            modified_description = description.replace('NO', 'SE')
            value = PREÇOS.loc[PREÇOS['Produto'] == modified_description, 'N'].values
        elif 'SU' in description:
            modified_description = description.replace('SU', 'SE')
            value = PREÇOS.loc[PREÇOS['Produto'] == modified_description, 'S'].values
        elif 'NE' in description:
            modified_description = description.replace('NE', 'SE')
            value = PREÇOS.loc[PREÇOS['Produto'] == modified_description, 'NE'].values
        elif 'SE' in description:
            value = PREÇOS.loc[PREÇOS['Produto'] == description, 'Preço'].values

        if value.size > 0:
            products.at[i, 'value'] = value[0]
        else:
            products.at[i, 'value'] = None
        products_list.append(products.loc[i][['tickerId','value']].to_dict())

    response = requests.post(url, headers=headers, json=products_list,verify=verifyCertificate)
    ####################################################################################################################################
    excel = pd.ExcelWriter('curva_fwd.xlsx', engine='xlsxwriter')
    curva_fwd.to_excel(excel, sheet_name='curva_fwd', index= False)
    excel.save()
    excel.close()

    #----------------------------------------------------------------------------------------------------
    #PEGAR PLANILHA DO EMAIL
    
    #le a planilha pra ver o tamanho da tabela (quantidade de linhas)
    planilha_email = pd.read_excel("Email.xlsx")
    maturidades = list(planilha_email.Maturity)

    volume_total = volumes_final_peq.loc[volumes_final_peq.Maturidade == 'Total','Volume'].item()
    Liquidity_total = volumes_final_peq.loc[volumes_final_peq.Maturidade == 'Total','Liquidez'].item()
    linhas = []
    for maturidade in maturidades:
        if " I5" in maturidade:
            maturidade = maturidade.replace(' I5', '')
            variacao_abs = round(PREÇOS.loc[PREÇOS.Maturidade == maturidade,'I5'].item() - PREÇOS.loc[PREÇOS.Maturidade == maturidade,'I5 Anterior'].item(),2)
            variacao_per = round(100*(PREÇOS.loc[PREÇOS.Maturidade == maturidade,'I5'].item() - PREÇOS.loc[PREÇOS.Maturidade == maturidade,'I5 Anterior'].item())/PREÇOS.loc[PREÇOS.Maturidade == maturidade,'I5'].item(),2)
            
            if variacao_abs<0:classe = 'negativo'
            elif variacao_abs>0: classe = 'positivo'
            else:classe = 'preto'

            linha = f"""<tr>
                    <td >{maturidade + ' I5'}</td>
                    <td>{" ".join([word.replace('CON', 'I5') for word in PREÇOS.loc[PREÇOS['Maturidade'] == maturidade, 'Produto'].values[0].split()])}</td>
                    <td>{volumes_final_peq.loc[volumes_final_peq.Maturidade == maturidade + ' I5','Operacoes'].item()}</td>
                    <td>{volumes_final_peq.loc[volumes_final_peq.Maturidade == maturidade + ' I5','Volume'].item()}</td>
                    <td>{PREÇOS.loc[PREÇOS.Maturidade == maturidade,'I5'].item()}</td>
                    <td>{PREÇOS.loc[PREÇOS.Maturidade == maturidade,'I5 Anterior'].item()}</td>
                    <td class = {classe}>{variacao_abs}</td>
                    <td class = {classe}>{variacao_per}</td>
                    <td class = 'preto'>{volumes_final_peq.loc[volumes_final_peq.Maturidade == maturidade + ' I5','Liquidez'].item()}</td>
                </tr>"""
        else:
            variacao_abs = round(PREÇOS.loc[PREÇOS.Maturidade == maturidade,'Preço'].item() - PREÇOS.loc[PREÇOS.Maturidade == maturidade,'Preço Anterior'].item(),2)
            variacao_per = round(100*(PREÇOS.loc[PREÇOS.Maturidade == maturidade,'Preço'].item() - PREÇOS.loc[PREÇOS.Maturidade == maturidade,'Preço Anterior'].item())/PREÇOS.loc[PREÇOS.Maturidade == maturidade,'Preço'].item(),2)
            
            if variacao_abs<0:classe = 'negativo'
            elif variacao_abs>0: classe = 'positivo'
            else:classe = 'preto'

            linha = f"""<tr>
                        <td>{maturidade}</td>
                        <td>{PREÇOS.loc[PREÇOS.Maturidade == maturidade,'Produto'].item()}</td>
                        <td>{volumes_final_peq.loc[volumes_final_peq.Maturidade == maturidade,'Operacoes'].item()}</td>
                        <td>{volumes_final_peq.loc[volumes_final_peq.Maturidade == maturidade,'Volume'].item()}</td>
                        <td>{PREÇOS.loc[PREÇOS.Maturidade == maturidade,'Preço'].item()}</td>
                        <td>{PREÇOS.loc[PREÇOS.Maturidade == maturidade,'Preço Anterior'].item()}</td>
                        <td class = {classe}>{variacao_abs}</td>
                        <td class = {classe}>{variacao_per}</td>
                        <td class = 'preto'>{volumes_final_peq.loc[volumes_final_peq.Maturidade == maturidade,'Liquidez'].item()}</td>
                    </tr>"""
        linhas.append(linha)
    style = """
            body {
                padding:1% 3% 1% 1%;
                font-size:14px;
                font-family: Calibri, sans-serif;
            }
            .table-container {
                overflow-x: auto;
            }
            table {
                border-collapse: collapse;
                width: fit-content
                text-align: center;
            }
            .tabela-2 {
                width: 25%;
            }
            th, td {
                border: 1px solid #ddd;
                padding: 4px 10px;
                text-align: center;
                width: fit-content
            }
            
            th {
                background-color: #8C5CF2; /* roxo claro */
                color: white;
                font-size:18px;
                font-style: italic;
                width: fit-content
            }
            tr:nth-child(even) {
                background-color: #f2f2f2; /* cor de fundo alternada */
            }
            .last-row {
                background-color: #16A837; /* verde */
                font-weight:bold;
                font-size:20px;
            }
            .negativo {
                font-weight:bold;
                color:red;
            }
            .positivo {
                font-weight:bold;
                color:green;
            }
            .preto {
                font-weight:bold;
                color:black;
            }
    """


    tabelas_html = f"""
    <html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
            {style}
        </style>
    </head>
    <body>
        <div>
            Prezados,<br><br>
        </div>

        <div>
            Segue a marcação da curva de hoje {data_hoje.strftime('%d/%m')}:<br>
            <br>
        </div>


        <div class="table-container">
            <table>
                <thead>
                    <th>Maturity</th>
                    <th>Product (SE)</th>
                    <th>Operations</th>
                    <th>Volume <br/>(MWm)</th>
                    <th>Price <br/>(Today)</th>
                    <th>Price <br/>(Yesterday)</th>
                    <th>Var. ($)</th>
                    <th>Var. (%)</th>
                    <th>Liquidity <br/>(Today vs Last 15d)</th>
                </thead>
                <tbody>
                    {''.join(linhas)}
                </tbody>
                <tfoot>
                    <!-- Última linha da tabela -->
                <tr class="last-row">
                    <td>Total</td>
                    <td colspan="7">Daily Trade Volume: {volume_total}  aMW-Month</td>
                    <td>{Liquidity_total}</td>
                </tfoot>
            </table>
        </div>


            <strong>OBS:</strong> O campo 'Liquidity' representa a comparação do volume negociado hoje contra a média dos últimos 15 dias.<br>
            <br>
            <br>
            
        </div>

        <div>
            <strong>Exemplo:</strong> (volume M1 de hoje)/(média diária de volume M1 nos últimos 15d) = 130% => Liquidez será 'Muito Alta'.<br>
            <br>
        </div>

        <table class="tabela-2">
            <thead>
                <tr>
                    <th>Liquidity Index Score</th>
                    <th>Perception</th>
                </tr>
            </thead>
            <tbody>
                <!-- Linhas da segunda tabela -->
                <tr>
                    <td>≥ 100</td>
                    <td>Very High</td>
                </tr>
                <tr>
                    <td>≥ 75</td>
                    <td>High</td>
                </tr>
                <tr>
                    <td>≥ 50</td>
                    <td>Normal</td>
                </tr>
                <tr>
                    <td>≥ 25</td>
                    <td>Low</td>
                </tr>
                <tr>
                    <td>≥ 0</td>
                    <td>Very Low</td>
                </tr>
                <!-- Mais linhas da segunda tabela aqui -->
            </tbody>
        </table>

        <div>
            Atenciosamente,<br>
            <br>
        </div>
    </body>
    </html>
    """

    with open("tabela.html", "w", encoding="utf-8") as f:
        f.write(tabelas_html)

    import os
    path = os.getcwd()
    # Defina o caminho para o arquivo HTML

    html_file_path = f"{path}/tabela.html"

    # Iniciar o navegador
    options = Options()
    options.add_argument("--headless")  # Executa o navegador em modo headless
    options.add_argument("window-size=1920x1080")  # Definir tamanho da janela
    driver = webdriver.Chrome(options=options)

    # Abrir o arquivo HTML local
    driver.get(f"file://{html_file_path}")

    # Encontrar o elemento usando By.CSS_SELECTOR
    elemento_precos = driver.find_element(By.CSS_SELECTOR, "div.table-container table")
    elemento_liquidez = driver.find_element(By.CSS_SELECTOR, "table.tabela-2")
    # Capturar o elemento como uma imagem
    elemento_precos.screenshot("Preços.png")
    elemento_liquidez.screenshot("Liquidez.png")
    # Fechar o navegador
    driver.quit()

    # Lista de destinátarios 
    destinatarios 

    outlook = client.Dispatch('Outlook.Application',pythoncom.CoInitialize())
    email_msg = outlook.CreateItem(0)


    email_msg.To = ';'.join(destinatarios)
    email_msg.Subject = f"Preços {data_hoje.strftime('%d/%m')}"

    imagem_path = Path(f"{path}/Preços.png")
    imagem_path2 = Path(f"{path}/Liquidez.png")

    attachment = email_msg.Attachments.Add(str(imagem_path))
    attachment2 = email_msg.Attachments.Add(str(imagem_path2))


    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", "imagecid1")
    attachment2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", "imagecid2")

    mensagem = f"""
    <html>
    <body>
        <div>
            Prezados,<br><br>
        </div>
        <div>
            Segue a marcação da curva de hoje {data_hoje.strftime('%d/%m')}<br>
            <br>
        </div>

        <div>
            <img src="cid:imagecid1"> <br><br>
        </div>

        <div>
            <strong>OBS:</strong> O campo 'Liquidity' representa a comparação do volume negociado hoje contra a média dos últimos 15 dias.<br>
            <br>
            <br>
        </div>

        <div>
            <strong>Exemplo:</strong> (volume M1 de hoje)/(média diária de volume M1 nos últimos 15d) = 130% => Liquidez será 'Muito Alta'.<br>
            <br>
        </div>

        <div>
            <img src="cid:imagecid2"> <br><br>
        </div>

        <div>
            Atenciosamente,<br>
            <br>
        </div>
    </body>
    </html>
    """

    email_msg.HTMLBody = mensagem

    email_msg.Display()

    mysql_func.insert(db_connection,volumes_final_peq, 'historico_volumes')

    print("\nDone!")
    print('=====================')
    return PREÇOS, volumes_final_peq

marcacao_mesa(user_banco,senha_banco,date.today())
import pandas as pd 
from PyPDF2 import PdfFileMerger, PdfFileReader
import os
import pdfkit
import locale

options={'page-size':'A4'}
path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)


teste = pd.read_excel('investimentos.xlsx','Lista Geral',skiprows=6)
teste = teste.iloc[:,1:]
municipios = teste['MUNICÍPIO'].dropna().unique()
filtros = {"concluido":['CONCLUÍDO'],"execucao":["CONTRATAÇÃO","EXECUÇÃO","HOMOLOGAÇÃO","PARALISADO"],"licitacao":["LICITAÇÃO","AÇÕES PREPARATÓRIAS","TRAMITAÇÃO PARA ABERTURA DA LICITAÇÃO","RELICITAÇÃO","PREPARAÇÃO DO EDITAL DE LICITAÇÃO"]}
teste.fillna(pd.np.nan)

def to_html_categoria(territorio,df):
    html = '''
        <table class="paleBlueRows" align='center'>
        <caption>INVESTIMENTOS TOTAIS POR CATEGORIA NO TERRITÓRIO {0}</caption>
        <thead>
        <tr>
        <th style='text-align: left;'>CATEGORIAS</th>
        <th></th></thclass='cifra'> 
        <th class='cifra'>VALOR TOTAL</th>
        </tr>
        </thead>'''.format(territorio)
    valor_total = 0.0
    for row in df.itertuples():
        html += "<tr><td style='text-align: left;'>{0}</td><td class='cifra'>R$</td><td style='text-align: right; padding-right: 15px;'>{1}</td></tr>".format(row[0].upper(),formatacao(row[1]))
        valor_total += row[1]
    #if quantidade_total == 0:
        #html += "<tr><td></td><td></td><td></td><td class='valor'></td></tr>"
    html += "<tfoot><tr><td style='text-align: left;'>Total Geral</td><td class='cifra'>R$</td><td style='text-align: right; padding-right: 15px;'>{0}</td></tr></tfoot>".format(formatacao(valor_total))
    html += "</tbody></table>"
    return html

def to_html_territorio(territorio,df):
    valor = 0.0
    html = '''<table class="paleBlueRows">
        <caption>DETALHAMENTO DOS INVESTIMENTOS DO TERRITÓRIO: {0}</caption>
        <thead>
        <tr>
        <th>Nº</th><th>CATEGORIAS</th><th>INVESTIMENTOS</th><th>ESTABELECIMENTO</th><th>MUNICÍPIO</th><th>FASE</th><th colspan='2' width=150px>VALOR TOTAL</th></tr>
        </thead>'''.format(territorio)
    #quantidade_total, valor_total = 0,0
    for row in df.itertuples():
        html += "<tr><td><b>{0}</b></td><td><b>{1}</b></td><td><b>{2}</b></td><td><b>{3}</b></td><td><b>{4}</b></td><td><b>{5}</b></td><td><b>R$</b></td><td style='text-align: right; padding-right: 15px;'><b>{6}</b></tr>".format(row[0],row[1],row[2],row[3],row[4],row[5],formatacao(row[6]))
        valor += row[6]
        try:
            if df.loc[row[0]+1,'CATEGORIA'] != '':
                html += "<tr bgcolor='Silver'><td colspan='6'></td><td><b>R$</b></td><td style='text-align: right; padding-right: 15px;'><b>"+formatacao(valor)+"</b></td></tr>"
                valor = 0.0
        except:
            html += "<tr bgcolor='Silver'><td colspan='6'></td><td><b>R$</b></td><td style='text-align: right; padding-right: 15px;'><b>"+formatacao(valor)+"</b></td></tr>"
            valor = 0.0

    html += "<tfoot><tr><td></td><td>Total Geral</td><td colspan='4'></td><td><b>R$</b></td><td style='text-align: right; padding-right: 15px;'><b>{0}</b></td></tr></tfoot>".format(formatacao(df['VALOR TOTAL (R$)'].sum()))
    html += "</tbody></table>"
    return html

def fase(fase):
    return fase in list(filtros.values())[0] or fase in list(filtros.values())[1] or fase in list(filtros.values())[2]

def formatacao(num):
    if '0' == str(num).split('.')[0]:
        return '-'
    else:
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
        num = locale.currency(float(num), grouping=True, symbol=None)
        return num[:-3]

territorios = teste['TERRITÓRIO'].dropna().unique()[:len(teste['TERRITÓRIO'].dropna().unique())-2]

for territorio in territorios:
    df_territorio = teste[teste['TERRITÓRIO'] == territorio]
    df_territorio = df_territorio.sort_values(by=['CATEGORIA','INVESTIMENTO','ESTABELECIMENTO','MUNICÍPIO']).groupby(['CATEGORIA','INVESTIMENTO','ESTABELECIMENTO','MUNICÍPIO'])['CATEGORIA','INVESTIMENTO','ESTABELECIMENTO','MUNICÍPIO','FASE','VALOR TOTAL (R$)'].head()#.agg({'FASE':'first','VALOR TOTAL (R$)':'first'})
    df_territorio = df_territorio[df_territorio['FASE'].apply(fase)] #filtro de fases
    df_territorio['Nº'] = range(1,len(df_territorio)+1)
    df_territorio= df_territorio.set_index(['Nº',df_territorio.index])
    df_territorio.reset_index()
    df_territorio = df_territorio.reset_index(level=[None])
    df_territorio = df_territorio.loc[:,['CATEGORIA','INVESTIMENTO','ESTABELECIMENTO','MUNICÍPIO','FASE','VALOR TOTAL (R$)']]
    df_territorio.loc[df_territorio['CATEGORIA'].duplicated(),'CATEGORIA'] = ''
    df_territorio.loc[df_territorio['INVESTIMENTO'].duplicated(),'INVESTIMENTO'] = ''
    df_territorio.loc[df_territorio['ESTABELECIMENTO'].duplicated(),'ESTABELECIMENTO'] = ''
    
    df_categoria = teste[teste['TERRITÓRIO'] == territorio]
    df_categoria.sort_values(by='CATEGORIA')
    df_categoria = df_categoria[df_categoria['FASE'].apply(fase)] #filtro de fases
    df_categoria = df_categoria.groupby('CATEGORIA').agg({'VALOR TOTAL (R$)':'sum'})
    
    tabela_categoria= to_html_categoria(territorio,df_categoria)
    tabela_detalhamento = to_html_territorio(territorio,df_territorio)
    #<meta name="pdfkit-orientation" content="Landscape"/>
    cabeca = '''<!doctype html>
        <html lang="ptbr">
        <head>
            <title>Relatório de Investimentos - Território {0}</title>
            <meta name="description" content="Relatório de investimentos">

            <meta name="author" content="Pedro">
            <meta charset="utf-8">
            <link rel="stylesheet" href="style.css">
        </head>
        <body align='center'>'''.format(territorio)
    html =  cabeca + '''<script src="js/scripts.js"></script>
            <img src={2} id = 'icone_gov' alt='RN'><img src={4} id = 'icone_proj' alt='GOVERNO CIDADÃO'><p id='texto'>{3}</p><br><br><p style='text-align: right;'>Data da última atualização: {5}</p>{0}<br><br>{1}'''.format(tabela_categoria,tabela_detalhamento,"'C:\\Users\\user\\Dropbox\\Notebooks\\gera_pdfs\\icone_gov.png'",
        "GOVERNO DO ESTADO DO RIO GRANDE DO NORTE<br>SECRETARIA DE ESTADO DO PLANEJAMENTO E DAS FINANÇAS<br>PROJETO INTEGRADO DE DESENVOLVIMENTO SUSTENTÁVEL DO RN<br><br><b>INVESTIMENTO NO TERRITÓRIO:<br>"+territorio+"         R$ "+formatacao(df_territorio['VALOR TOTAL (R$)'].sum())+"</b>", "'C:\\Users\\user\\Dropbox\\Notebooks\\gera_pdfs\\icone_proj.png'",'14/01/2020')

    pdfkit.from_string(html,'relatorios/territorios/'+territorio+'.pdf',css='style.css',configuration=config,options=options)




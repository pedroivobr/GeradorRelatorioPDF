import pandas as pd 
from PyPDF2 import PdfFileMerger, PdfFileReader
import os
import pdfkit
import locale

options={'page-size':'A3'}
path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)


teste = pd.read_excel('investimentos.xlsx','Lista Geral',skiprows=6)
teste = teste.iloc[:,1:]
municipios = teste['MUNICÍPIO'].dropna().unique()
filtros = {"concluido":['CONCLUÍDO'],"execucao":["CONTRATAÇÃO","EXECUÇÃO","HOMOLOGAÇÃO","PARALISADO"],"licitacao":["LICITAÇÃO","AÇÕES PREPARATÓRIAS","TRAMITAÇÃO PARA ABERTURA DA LICITAÇÃO","RELICITAÇÃO","PREPARAÇÃO DO EDITAL DE LICITAÇÃO"]}
teste.fillna(pd.np.nan)

## RETORNA UM DATAFRAME COM DADOS DO MUNICIPIO DE ACORDO COM O FILTRO
def filtro(dados,municipio,filtros):
    aux = pd.DataFrame()
    for filtro in filtros:
        aux = aux.append(dados[dados['FASE  '].astype(str) == str(filtro)])
    aux = aux[aux['MUNICÍPIO'] == municipio]
    df = aux.groupby('CATEGORIA').agg(['count','sum'])['VALOR TOTAL (R$)']
    df.columns = ['Quantidade de Investimentos','Valor Total (R$)']
    df.index.names = ['Categorias']
    df_saida = aux.groupby(['CATEGORIA','INVESTIMENTO','ESTABELECIMENTO 2']).agg({'VALOR TOTAL (R$)':'sum', 'VALOR DE PROJETOS':'sum', 
    'VALOR OBRAS':'sum','VALOR SUBPROJETO':'sum','VALOR EQUIPAMENTOS* (Valor Médio)':'sum','VALOR SERVIÇOS':'sum','% DE EXECUÇÃO GERAL':'mean','OBSERVAÇÃO':'first'})
    df_saida.columns = ['VALOR TOTAL (R$)','VALOR DE PROJETOS','VALOR OBRAS ESTRUTURANTES','VALOR SUBPROJETO','VALOR EQUIPAMENTOS','VALOR SERVIÇOS','PORCENTAGEM GERAL','OBSERVAÇÃO']
    pd.to_numeric(df_saida['VALOR SUBPROJETO'], downcast='float')
    df_saida = df_saida.round(2)
    df_saida['PORCENTAGEM GERAL'] = df_saida['PORCENTAGEM GERAL'].astype(int)*100
    df_saida['PORCENTAGEM GERAL'] = df_saida['PORCENTAGEM GERAL'].astype(str) + '%'
    df_saida['Nº'] = range(1,len(df_saida)+1)
    df_saida = df_saida.set_index(['Nº',df_saida.index])
    df_saida = df_saida.reset_index(level=['CATEGORIA','INVESTIMENTO','ESTABELECIMENTO 2'])
    df_saida.loc[df_saida['CATEGORIA'].duplicated(),'CATEGORIA'] = ''
    df_saida.loc[df_saida['INVESTIMENTO'].duplicated(),'INVESTIMENTO'] = ''
    return df, aux,df_saida #df (tabelinha do filtro), aux(tabela completa de acordo com o filtro), df_saida(df formatado para a tabela pivot)

def to_html_pivot(df, investimento,city,total, quantidade):
    valor = [0,0,0,0,0,0]
    html = '''<caption><h1>DETALHAMENTO DOS INVESTIMENTOS {0} DO MUNICÍPIO DE {1}</h1><table><tr><td width='15%'></td><td><h3>R$</h3></td><td><h3>{2}</h3></td><td width='30%'></td><td><h3>{3}</h3></td><td><h3>INVESTIMENTOS</h3></td></tr></table></caption>
        <table class="paleBlueRows">
        <thead>
        <tr>
        <th>Nº</th><th>CATEGORIAS</th><th>INVESTIMENTOS</th><th>ESTABELECIMENTO</th><th colspan='2' width=150px>VALOR TOTAL</th><th colspan='2'>VALOR DE PROJETOS</th><th colspan='2'>VALOR OBRAS ESTRUTURANTES</th><th colspan='2'>VALOR SUBPROJETO</th><th colspan='2' width=200px>VALOR EQUIPAMENTOS</th><th colspan='2'>VALOR SERVIÇOS</th><th>PROCETAGEM GERAL</th><th>OBSERVAÇÕES</th></tr>
        </thead>'''.format(investimento,city,formatacao(total),quantidade)
    #quantidade_total, valor_total = 0,0
    for row in df.itertuples():
        html += "<tr><td><b>{0}</b></td><td><b>{1}</b></td><td><b>{2}</b></td><td><b>{3}</b></td><td><b>R$</b></td><td><b>{4}</b></td><td><b>R$</td></b><td><b>{5}</b></td><td><b>R$</b></td><td><b>{6}</b></td><td><b>R$</b></td><td><b>{7}</b></td><td><b>R$</b></td><td><b>{8}</b></td><td><b>R$</b></td><td><b>{9}</b></td><td><b>{10}</b></td><td>{11}</td></tr>".format(row[0],row[1],row[2],row[3],formatacao(row[4]),formatacao(row[5]),formatacao(row[6]),formatacao(row[7]),formatacao(row[8]),formatacao(row[9]),row[10],row[11])
        
        try:
            if df.loc[row[0]+1,'CATEGORIA'] != '':
                html += "<tr><td colspan='18' bgcolor='grey'><p></p></td></tr>"
        except:
            html += "<tr><td colspan='18' bgcolor='grey'><p></p></td></tr>"
        valor[0] += row[4]
        valor[1] += row[5]
        valor[2] += row[6]
        valor[3] += row[7]
        valor[4] += row[8]
        valor[5] += row[9]
    html += "<tfoot><tr><td></td><td>Total Geral</td><td></td><td></td><td>R$</td><td>{0}</td><td>R$</td><td>{1}</td><td>R$</td><td>{2}</td><td>R$</td><td>{3}</td><td>R$</td><td>{4}</td><td>R$</td><td>{5}</td><td></td><td></td></tr></tfoot>".format(formatacao(valor[0]),formatacao(valor[1]),formatacao(valor[2]),formatacao(valor[3]),formatacao(valor[4]),formatacao(valor[5]))
    return html

def to_html(df,investimento):
    html = '''<b>INVESTIMENTO {0}</b>
        <table class="paleBlueRows" align='center'>
        <thead>
        <tr>
        <th>CATEGORIAS</th>
        <th>QUANTIDADE DE INVESTIMENTOS</th>
        <th></th></thclass='cifra'> 
        <th class='cifra'>VALOR TOTAL (R$)</th>
        </tr>
        </thead>'''.format(investimento)
    quantidade_total, valor_total = 0,0
    for row in df.itertuples():
        html += "<tr><td>{0}</td><td>{1}</td><td class='cifra'>R$</td><td class='valor'>{2}</td></tr>".format(row[0],row[1],formatacao(row[2]))
        quantidade_total += row[1]
        valor_total += row[2]
    #if quantidade_total == 0:
        #html += "<tr><td></td><td></td><td></td><td class='valor'></td></tr>"
    html += "<tfoot><tr><td>Total Geral</td><td>{0}</td><td class='cifra'>R$</td><td class='valor'>{1}</td></tr></tfoot>".format(quantidade_total,formatacao(valor_total))
    html += "</tbody></table>"
    return html, valor_total, quantidade_total

def formatacao(num):
    if '0' == str(num).split('.')[0]:
        return '-'
    else:
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
        num = locale.currency(float(num), grouping=True, symbol=None)
        return num[:-3]

#limpar obras
for i in range(0,len(teste)-1):
    try:
        if str(teste.loc[i,'VALOR OBRAS'])[-3] == ',':
            teste.loc[i,'VALOR OBRAS'] = str(teste.loc[i,'VALOR OBRAS']).replace('.','').replace(',','.')
            print(teste.loc[i,'VALOR OBRAS'])
    except:
        pass
teste['VALOR OBRAS'] = teste['VALOR OBRAS'].astype('float64')

#cidades = pd.Series(['BENTO FERNANDES','CERRO CORÁ','CAICÓ'])#teste['MUNICÍPIO'].dropna().unique()
for cidade in cidades:
    sc_execucao, sc_execucao_saida,pivotada_execucao = filtro(teste,cidade,filtros['execucao'])
    sc_concluido, sc_concluido_saida, pivotada_concluido = filtro(teste,cidade,filtros['concluido'])
    sc_licitacao, sc_licitacao_saida,pivotada_licitacao = filtro(teste,cidade,filtros['licitacao'])

    tabela_concluido, total_concluido, quantidade_concluido = to_html(sc_concluido,'CONCLUÍDOS')
    tabela_execucao, total_execucao, quantidade_execucao = to_html(sc_execucao,'EM EXECUÇÃO')
    tabela_licitacao, total_licitacao, quantidade_licitacao = to_html(sc_licitacao,'EM LICITAÇÃO')
    valor_total = total_licitacao+total_concluido+total_execucao
    quantidade_total = quantidade_execucao + quantidade_licitacao + quantidade_concluido

    pivot_concluido = to_html_pivot(pivotada_concluido,'CONCLUÍDOS',cidade,total_concluido, quantidade_concluido)
    pivot_execucao = to_html_pivot(pivotada_execucao,'EM EXECUÇÃO',cidade,total_execucao, quantidade_execucao)
    pivot_licitacao = to_html_pivot(pivotada_licitacao,'EM LICITAÇÃO',cidade,total_licitacao, quantidade_licitacao)

    #pdfkit.from_string(html,'saida_p.pdf',css='style.css',configuration=config)
    quebra_linha = '<p class="new-page"></p>'
    html = '''<!doctype html>
        <html lang="ptbr">
        <head>
            <title>The HTML5 Herald</title>
            <meta name="description" content="Relatório de investimentos">
            <meta name="pdfkit-orientation" content="Landscape"/>
            <meta name="author" content="Pedro">
            <meta charset="utf-8">
            <link rel="stylesheet" href="style.css">
        </head>
        <body align='center'>'''

    resumo =  '''
            <script src="js/scripts.js"></script>
        <img src={3} id = 'icone_gov' alt='RN'><img src={5} id = 'icone_proj' alt='GOVERNO CIDADÃO'><p id='texto'>{4}</p>
        <table border="0px solid" align="center" width="100%"><tr><td><h1>R$ {7}</h1></td><td></td><td><h1>{6} INVESTIMENTOS</h1></td></tr></table><p align="right">Data da última atualização:{8}</p><table border="0px solid" align="center" width="100%">
        <tr><td style="vertical-align:top">{0}</td><td style="vertical-align:top">{1}</td><td style="vertical-align:top">{2}</td></tr>
    </table>'''.format(tabela_concluido,tabela_execucao,tabela_licitacao,"'C:\\Users\\user\\Dropbox\\Notebooks\\gera_pdfs\\icone_gov.png'",
    "GOVERNO DO ESTADO DO RIO GRANDE DO NORTE<br>SECRETARIA DE ESTADO DO PLANEJAMENTO E DAS FINANÇAS<br>PROJETO INTEGRADO DE DESENVOLVIMENTO SUSTENTÁVEL DO RN<br><br><b>INVESTIMENTO NO MUNICÍPIO DE "+cidade+"</b>",
    "'C:\\Users\\user\\Dropbox\\Notebooks\\gera_pdfs\\icone_proj.png'",quantidade_total,formatacao(valor_total),'14/01/2020')
    #html = html + quebra_linha + pivot_concluido + quebra_linha + pivot_execucao + quebra_linha + pivot_licitacao

    pdfkit.from_string(html+resumo,'relatorios/'+cidade+'_1.pdf',css='style.css',configuration=config,options=options)
    pdfkit.from_string(html+pivot_concluido+'</body></html>','relatorios/'+cidade+'_2.pdf',css='style.css',configuration=config,options=options)
    pdfkit.from_string(html+pivot_execucao+'</body></html>','relatorios/'+cidade+'_3.pdf',css='style.css',configuration=config,options=options)
    pdfkit.from_string(html+pivot_licitacao+'</body></html>','relatorios/'+cidade+'_4.pdf',css='style.css',configuration=config,options=options)

    merger = PdfFileMerger()
    inputFile = open('relatorios/'+cidade+'_1.pdf', 'rb')
    inputFile2 = open('relatorios/'+cidade+'_2.pdf', 'rb')
    inputFile3 = open('relatorios/'+cidade+'_3.pdf', 'rb')
    inputFile4 = open('relatorios/'+cidade+'_4.pdf', 'rb')
    merger.append(inputFile,import_bookmarks=False)
    merger.append(inputFile2,import_bookmarks=False)
    merger.append(inputFile3,import_bookmarks=False)
    merger.append(inputFile4,import_bookmarks=False)

    merger.write('relatorios/'+cidade+".pdf")
    inputFile.close()
    inputFile2.close()
    inputFile3.close()
    inputFile4.close()
    try:
        os.remove('relatorios/'+cidade+'_1.pdf')
        os.remove('relatorios/'+cidade+'_2.pdf')
        os.remove('relatorios/'+cidade+'_3.pdf')
        os.remove('relatorios/'+cidade+'_4.pdf')
    except:
        pass
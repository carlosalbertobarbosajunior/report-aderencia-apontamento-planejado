#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import configparser
import pyodbc
from datetime import datetime
import sys
import win32com.client


# In[ ]:


def extract_planned_codes():
    '''
    Outputs:
      df           - Dataframe contendo todos os códigos de apontamento planejados 
                     para a semana requisitada pelo usuário.
    Ação:
      Armazena as informações de ano e semana da programação nas respectivas
      variáveis de mesmo nome. Em seguida, tenta gerar um dataframe com o endereço
      do arquivo na pasta de planejamento com essas informações. Caso não consiga,
      retorna uma mensagem de erro.
    '''
    # Requisitando ano e semana desejados pelo usuário
    ano = input('Digite o ANO da programação que deseja analisar:\n')
    semana = input('Agora, digite a SEMANA da programação:\n')
    
    # Tenta gerar um dataframe com essas informações
    try:
        df = pd.read_excel(f'K:\\29 - PROGRAMAÇÃO\\Semana_{semana}-{ano}\\Prog_{semana}_{ano}.xlsx', sheet_name='sandbox')
        df = df[['COD_OS_COMPLETO', 'DT_FINALIZACAO', 'CODIGO', 'TIPOSERVICO']]
        df = df.rename(columns={'COD_OS_COMPLETO': 'OS', 'DT_FINALIZACAO':'TITULO'})
        return df, ano, semana
    
    # Retorna uma mensagem de erro caso não consiga
    except:
        print('Não foi possível identificar a semana ou o ano. Por favor, verifique as informações e tente novamente.')
        return None


# In[ ]:


def create_df_from_database(query, config_file='G:\Ciência de Dados\Segurança\config.ini'):
    '''
    Inputs:
      config_file - Caminho do arquivo .ini (string) que possui as 
                    configurações de conexão com o banco de dados 
                    da empresa.
      query       - Instrução em SQL (string) que será transformada
                    em um dataframe do pandas.
                    
    Outputs:
      df          - Dataframe do pandas construído a partir da instrução
                    SQL da variável query.
    '''
    
    # Executando o configparser e extraindo informações do config.ini
    config = configparser.ConfigParser()
    config.read(config_file)
    server = config['database']['server']
    database = config['database']['database']
    username = config['database']['username']
    password = config['database']['password']
    
    # Conectando ao banco de dados e criando o dataframe com a instrução sql
    conn = pyodbc.connect('DRIVER={SQL Server Native Client 10.0};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password)
    df = pd.read_sql(query, conn)
    
    # Fechando a conexão para liberá-la
    conn.close()
    
    return df


# In[ ]:


def date_converter(data):
    '''
    Inputs:
      data            - Informação de data em formato convencional utilizado
                        no Brasil (aa/mm/yyyy).
    
    Outputs:
      data_convertida - Data em formato yyyy-mm-dd.
    '''
    
    data_obj = datetime.strptime(data, '%d/%m/%Y')
    data_convertida = data_obj.strftime('%Y-%m-%d')
    return data_convertida


# In[ ]:


def extract_executed_codes():
    '''
    Outputs:
      df              - dataframe de apenas uma coluna, com informação
                        de todos os códigos únicos de apontamento realizados
                        no intervalo de tempo definido pelo usuário.
    
    Ação:
      Inicialmente, utiliza a função date_converter para converter as informações
      de data inicial e final solicitadas ao usuário. As datas são utilizadas
      para construir uma consulta (variável query) ao banco de dados - acionando a 
      função create_df_from_database - das informações únicas de código neste período. 
      Caso as datas não possam ser convertidas, a função retorna uma mensagem de erro. 
      Caso não hajam apontamentos realizados neste período, a função retorna uma mensagem 
      informando o usuário.
    '''
    
    # Tenta converter as datas fornecidas pelo usuário através da função date_converter
    try:
        data1 = input('Digite a DATA INICIAL dos códigos apontados (formato dd/mm/aaaa): ')
        data_inicial = date_converter(data1)
        data2 = input('Digite a DATA FINAL dos códigos apontados (formato dd/mm/aaaa): ')
        data_final = date_converter(data2)
        
        # Define uma consulta com datas entre as estipuladas pelo usuário
        query = f'''
            select distinct
            cod_barr
        from 
            tctrl_ph
        where 
            data between '{data_inicial}' and '{data_final}'
        '''
        # Cria um dataframe com base na query, usando a função create_df_from_database
        df = create_df_from_database(query=query)
        # Caso não hajam apontamentos no período, informa isso ao usuário
        if len(df) == 0:
            print('Nenhum apontamento encontrado neste período. Por favor, avalie as datas e tente novamente.')
            return None
        # Retorna o dataframe com os códigos únicos de apontamento
        return df['cod_barr'], data1, data2
    # Caso não haja sucesso em converter a data, informa isso ao usuário
    except:
        print('Data inválida. Por favor, verifique as informações e tente novamente.')
        return None


# In[ ]:


def df_to_html_body(df):
    '''
    Input:
      df                                           - Dataframe pandas
      
    Output:
      '<html><body>'+df.to_html()+'</body></html>' - Tabela com as informações do dataframe
                                                     pandas, porém com formatação adaptada
                                                     para html, utilizada para compor corpo
                                                     de e-mail.
    '''
    return '<html><body>'+df.to_html()+'</body></html>'


# In[ ]:


def create_report_by_user_informations():
    '''
    Outputs:
      Retorna None
      
    Ação:
      Função principal, que será acionada. Interage com o usuário, acionando as funções de 
      criação de dataframes a medida que as informações são preenchidas. Posteriormente, 
      analisa as atividades planejadas e executadas, e envia por e-mail duas tabelas HTML: 
      uma com as atividades planejadas E executadas, e outra com as atividades planejadas 
      mas NÃO executadas.
    '''
    
    # Interação com o usuário
    print('ATENÇÃO: ANTES DE PROSSEGUIR GARANTA QUE SEU USUÁRIO POSSUI PERMISSÃO ÀS PASTAS DE PLANEJAMENTO E GESTÃO DE CONTRATOS. ALÉM DISSO, PRECISAM ESTAR MAPEADAS COM AS LETRAS K E G, RESPECTIVAMENTE.\n\n')
    print('Olá! Seja bem-vindo ao assistente de PPs não apontadas por semana.\n')
    print('Primeiro, preciso das informações de programação, que estão dentro da rede da empresa.')
    
    # Extraindo os códigos planejados para um dataframe
    df_codigos_planejados, ano_planejado, semana_planejado = extract_planned_codes()
    print('...')
    print('Ótimo!',df_codigos_planejados.shape[0], 'códigos únicos foram lidos.\n')
    print('Agora preciso que você determine as datas que irei buscar dentro do banco de dados do GRV.\n')
    
    # Extraindo os códigos apontados para um dataframe
    df_codigos_apontados, data_inicial, data_final = extract_executed_codes()
    print('Muito bem, ', df_codigos_apontados.shape[0],'códigos únicos foram lidos.\n')
    print('Muito obrigado! Aguarde enquanto processo as informações e crio o relatório.')
    print('...')
    
    # Inserindo as informações dos dataframes em listas
    lista_codigos_planejados = list(df_codigos_planejados['CODIGO'])
    lista_codigos_apontados = list(map(int, df_codigos_apontados))
    
    # Realizando o comparativo entre as atividades planejadas e executadas, e armazenando
    # em listas distintas
    codigos_planejados_e_apontados = []
    codigos_nao_apontados = []
    for codigo in lista_codigos_planejados:
        if codigo in lista_codigos_apontados:
            codigos_planejados_e_apontados.append(codigo)
        else:
            codigos_nao_apontados.append(codigo)
            
    # Criando dois dataframes, filtrados pelas listas.
    df_codigos_planejados_e_apontados = df_codigos_planejados[df_codigos_planejados['CODIGO'].isin(codigos_planejados_e_apontados)].set_index('CODIGO').sort_index()
    df_codigos_planejados_nao_apontados = df_codigos_planejados[df_codigos_planejados['CODIGO'].isin(codigos_nao_apontados)].set_index('CODIGO').sort_index()    
    
    # Construindo tabelas HTML a partir dos dataframes
    html_codigos_planejados_e_apontados = df_to_html_body(df_codigos_planejados_e_apontados)
    html_codigos_planejados_nao_apontados = df_to_html_body(df_codigos_planejados_nao_apontados)
    
    # Adicionando as tabelas ao corpo do e-mail caso não estejam vazias
    body = ''
    if df_codigos_planejados_e_apontados.shape[0] != 0:
        body += '<br>Códigos apontados de acordo com o planejamento: ('+str(df_codigos_planejados_e_apontados.shape[0])+')<br>'+html_codigos_planejados_e_apontados
    if df_codigos_planejados_nao_apontados.shape[0] != 0:
        body += '<br>Códigos planejados, mas não apontados: ('+str(df_codigos_planejados_nao_apontados.shape[0])+')<br>'+html_codigos_planejados_nao_apontados
        
    # Enviando o e-mail caso exista informação para ser enviada
    if body != '':
        outlook = win32com.client.Dispatch("Outlook.Application")
        Msg = outlook.CreateItem(0)
        Msg.Subject = 'Report: Aderência de apontamentos na semana'
        Msg.HTMLBody = 'Bom dia!<br>Abaixo as informações de apontamento de acordo com o planejado:<br>Planejamento: Semana '+semana_planejado+'/'+ano_planejado+'<br> Apontamento: de '+data_inicial+' a '+data_final+'<br>'+body
        Msg.To = 'raony.silva@hkm.ind.br;victoria.fiorese@hkm.ind.br'
        #Msg.To = 'carlos.junior@hkm.ind.br'
        Msg.Send()
        sucesso = input('E-mail enviado com sucesso!')
        
    else:
        erro = input('Não foi encontrado nenhum apontamento. Por favor, verifique as informações e tente novamente.')  
    return None


# In[ ]:


create_report_by_user_informations()


# report-aderencia-apontamento-planejado
Script capaz de ler as informações de programação do setor do planejamento e compará-los com os códigos de apontamento do ERP. O objetivo é retornar as atividades planejadas que não foram apontadas.

## 1. Contextualização, objetivo e metodologia:
O setor de planejamento da HKM Indústria e Comércio necessita verificar, semanalmente, a aderência entre as atividades planejadas e as executadas pela produção.<br>
Em um processo de reeducação do apontamento de fábrica, é comum verificar apontamentos fora das atividades planejadas, ou ainda sobrecarregar uma atividade de múltiplos apontamentos, enquanto outras tarefas permanecem sem as horas trabalhadas que foram estipuladas.<br>
Para auxiliar na visualização do tamanho da problemática, este algoritmo gera duas tabelas:
    1. Apontamento de acordo com o planejado: Códigos de atividade que, de fato, houveram atividade tal qual o planejado.
    2. Códigos planejados, mas sem apontamento: Códigos que deveriam ter atividades atreladas, mas por qualquer motivo não indicaram registro de trabalho no intervalo definido.    
As informações do planejamento são extraídas de uma tabela em excel, compartilhada na rede da empresa. Os dados de apontamento de fábrica são adquiridos via consulta ao banco de dados do ERP GRV.<br>
O script finalizado foi baixado em formato .py (disponibilizado no repositório), e em seguida utilizou-se a biblioteca pyinstaller para gerar um executável, facilitando a interação com o usuário e eliminando a necessidade de ter python instalado no computador.

## 2. Bibliotecas:
pandas, configparser, pyodbc, datetime, sys, win32com.client

## 3. Funções indiretas:
#### &nbsp;&nbsp;&nbsp;&nbsp;- extract_planned_codes():
    Outputs:
      df           - Dataframe contendo todos os códigos de apontamento planejados 
                     para a semana requisitada pelo usuário.
      ano          - Ano da programação, determinado pelo usuário.
      semana       - Semana da programação, determinado pelo usuário.
      
    Ação:
      Armazena as informações de ano e semana da programação nas respectivas
      variáveis de mesmo nome. Em seguida, tenta gerar um dataframe com o endereço
      do arquivo na pasta de planejamento com essas informações. Caso não consiga,
      retorna uma mensagem de erro.

#### &nbsp;&nbsp;&nbsp;&nbsp;- create_df_from_database(query, config_file):
    Inputs:
      config_file - Caminho do arquivo .ini (string) que possui as 
                    configurações de conexão com o banco de dados 
                    da empresa.
      query       - Instrução em SQL (string) que será transformada
                    em um dataframe do pandas.
                    
    Outputs:
      df          - Dataframe do pandas construído a partir da instrução
                    SQL da variável query.

#### &nbsp;&nbsp;&nbsp;&nbsp;- date_converter(data):
    Inputs:
      data            - Informação de data em formato convencional utilizado
                        no Brasil (dd/mm/yyyy).
    
    Outputs:
      data_convertida - Data em formato yyyy-mm-dd.
      
#### &nbsp;&nbsp;&nbsp;&nbsp;- extract_executed_codes():
    Outputs:
      df['cod_barr']    - Dataframe de apenas uma coluna, com informação
                        de todos os códigos únicos de apontamento realizados
                        no intervalo de tempo definido pelo usuário.
      data1             - Data inicial da análise inserida pelo usuário, no 
                        formato (dd/mm/yyyy).
      data2             - Data final da análise inserida pelo usuário, no formato
                        (dd/mm/yyyy).
    
    Ação:
      Inicialmente, utiliza a função date_converter para converter as informações
      de data inicial e final solicitadas ao usuário. As datas são utilizadas
      para construir uma consulta (variável query) ao banco de dados - acionando a 
      função create_df_from_database - das informações únicas de código neste período. 
      Caso as datas não possam ser convertidas, a função retorna uma mensagem de erro. 
      Caso não hajam apontamentos realizados neste período, a função retorna uma mensagem 
      informando o usuário.
      
#### &nbsp;&nbsp;&nbsp;&nbsp;- df_to_html_body(df): 
    Input:
      df                                           - Dataframe pandas
      
    Output:
      '<html><body>'+df.to_html()+'</body></html>' - Tabela com as informações do dataframe
                                                     pandas, porém com formatação adaptada
                                                     para html, utilizada para compor corpo
                                                     de e-mail.
                                                     
## 4. Função principal:
#### &nbsp;&nbsp;&nbsp;&nbsp;- create_report_by_user_informations():
    Outputs:
      Retorna None
      
    Ação:
      Função principal, que será acionada. Interage com o usuário, acionando as funções de 
      criação de dataframes a medida que as informações são preenchidas. Posteriormente, 
      analisa as atividades planejadas e executadas, e envia por e-mail duas tabelas HTML: 
      uma com as atividades planejadas E executadas, e outra com as atividades planejadas 
      mas NÃO executadas.

Este projeto foi desenvolvido por Glauber Marques, analista de dados da área de planejamento, com o objetivo de transferir o processo de ETL (Extract, Transform, Load) para uma estrutura em Python. Essa transferência visa reduzir as etapas no Power BI, economizando recursos de processamento.

Descrição do Projeto
O projeto consiste em dois principais módulos:

Combinação de Arquivos
O primeiro módulo, combinar_arquivos, é responsável por combinar vários arquivos Excel em uma pasta específica em um único arquivo. Ele realiza as seguintes etapas:

Percorre uma pasta específica em busca de arquivos Excel.
Lê cada arquivo Excel e verifica se é possível acessar a aba "Page 1".
Combina os dados de todos os arquivos válidos em um único DataFrame.
Salva o DataFrame combinado em um novo arquivo Excel.
Pré-processamento de Arquivos do ServiceNow
O segundo módulo, preprocessar_servicenow_file, é responsável pelo pré-processamento de arquivos do ServiceNow. Ele realiza as seguintes etapas:

Lê um arquivo Excel do ServiceNow.
Seleciona apenas as colunas desejadas para análise.
Aplica funções de processamento aos dados, como verificar o estado, prioridade e código de solução.
Salva o DataFrame processado em um novo arquivo Excel.
Pré-requisitos
Para executar este projeto, você precisará das seguintes bibliotecas Python:

pandas
openpyxl
Você pode instalá-las executando o seguinte comando:


pip install pandas openpyxl
Execução do Projeto
Para executar o projeto, basta executar o arquivo Python etl.py. Certifique-se de ter os arquivos necessários no diretório especificado.


Copiar código
python etl.py
Estrutura do Projeto
O projeto possui a seguinte estrutura de arquivos:

etl.py: Arquivo principal contendo os módulos de combinação de arquivos e pré-processamento de arquivos do ServiceNow.
meu_log.log: Arquivo de log onde são registradas as informações sobre a execução do programa.
f_Servicenow_combinado.xlsx: Arquivo resultante da combinação dos arquivos Excel da pasta especificada.
f_Servicenow.xlsx: Arquivo resultante do pré-processamento do arquivo do ServiceNow.
Verificação de Arquivo


Contribuições
Contribuições são bem-vindas! Se você encontrar problemas ou tiver sugestões para melhorias, sinta-se à vontade para abrir uma issue ou enviar um pull request.

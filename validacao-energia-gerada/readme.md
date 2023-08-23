# Analisador de Dados de Medidores

Este é um script Python projetado para analisar dados de medidores armazenados em arquivos Excel. O script realiza as seguintes tarefas:

1. Localiza todos os arquivos Excel (.xlsx) ou planilhas (.xls) no diretório atual, exceto aqueles contendo "Código Medidores" no nome do arquivo.
2. Combina os dados de todas as planilhas em um único DataFrame.
3. Realiza uma série de validações nas colunas do DataFrame de acordo com um esquema específico.
4. Identifica registros inválidos com base em regras específicas e cria um novo arquivo Excel ("Memoria-Invalida.xlsx") contendo esses registros.
5. Caso o arquivo "Código Medidores.xlsx" esteja presente, ele utiliza informações desse arquivo para atualizar os identificadores dos registros inválidos.

## Executando o Script

1. Coloque o script Python na mesma pasta que contém os arquivos Excel (.xlsx) ou planilhas (.xls) que você deseja analisar.
2. Execute o script Python.

O script irá analisar os dados dos arquivos e, se houver ocorrências inválidas, ele as identificará e as armazenará no arquivo "Memoria-Invalida.xlsx".

## Observações

- O script utiliza a biblioteca pandas para manipular os dados dos arquivos Excel.
- Certifique-se de ter o pandas e a biblioteca pandera instalados antes de executar o script.
- O script não faz a análise dos dados em si, mas executa validações com base nas regras especificadas.
- Dados inválidos são aqueles que não atendem às regras definidas nas validações.
- O arquivo "Código Medidores.xlsx" é utilizado para atualizar os identificadores dos registros inválidos, se estiver presente.
- Infelizmente não posso disponibilizar os arquivos que são utilizados😢.

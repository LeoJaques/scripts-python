## Utilizando o Script

1. Coloque o script Python na mesma pasta que contém os arquivos Excel (.xlsx ou .xls) que você deseja analisar e processar.
2. Execute o script Python.
3. O script solicitará que você selecione o arquivo da prévia para processamento.
4. O script irá ler o arquivo selecionado, realizar ajustes nas colunas de datas e filtrar os dados relevantes.
5. Ele irá criar três novos arquivos Excel: "CCEAR - BB.xlsx", "CCEAR - BRA.xlsx" e "PREVIA CCEAR.xlsx".
6. Os arquivos de saída serão formatados com ajustes de cabeçalho, largura de coluna e bordas.

## Observações

- O script utiliza as bibliotecas pandas, openpyxl e openpyxl.styles para manipular e formatar os dados dos arquivos Excel.
- Certifique-se de ter as bibliotecas pandas e openpyxl instaladas antes de executar o script.
- O script permite a seleção de um arquivo da prévia para processamento, o que é útil caso você tenha vários arquivos na pasta.
- Os arquivos de saída são criados com as informações relevantes filtradas e com formatação aprimorada.
- Certifique-se de ajustar as extensões permitidas no código, caso queira adicionar ou remover extensões de arquivo a serem processadas.
- Lembre-se de que os arquivos criados pelo script podem conter informações sensíveis e devem ser tratados com cuidado de acordo com as políticas de privacidade e segurança.

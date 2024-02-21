# Projeto de Automação de Formulários

Este projeto usa Selenium e Python para preencher automaticamente formulários.

## Dependências

O projeto depende das seguintes bibliotecas Python:

- Selenium
- pandas

## Como usar

1. Certifique-se de que todas as dependências estão instaladas.
2. Execute o script Python principal.

O script irá ler os dados de um arquivo Excel, preencher o formulário e salvar os resultados de volta no arquivo Excel.

## Detalhes do Script

O script contém várias funções:

- `preenche_planilha()`: Esta função preenche o formulário com os dados fornecidos.
- `main()`: Esta é a função principal que lê os dados do arquivo Excel, chama a função `preenche_planilha()` para cada linha de dados e salva os resultados.

## Contribuindo

Contribuições para este projeto são bem-vindas. Por favor, abra uma issue ou um pull request se você quiser contribuir.

## Licença

Este projeto é licenciado sob os termos da licença MIT.

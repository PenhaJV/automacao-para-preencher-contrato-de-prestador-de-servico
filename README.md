# Gerador de Contratos de Prestação de Serviços

Este projeto é um script em Python que automatiza a geração de contratos de prestação de serviços para fornecedores, extraindo dados de uma planilha Excel e gerando documentos no formato Word (.docx).

## Funcionalidades

- Carrega uma planilha Excel contendo informações dos fornecedores.
- Gera contratos personalizados para cada fornecedor, utilizando um modelo de texto pré-definido.
- Salva os contratos gerados em um diretório específico.

## Requisitos

- Python 3.x
- Pacotes Python:
  - `openpyxl` para manipulação de arquivos Excel.
  - `python-docx` para criação de documentos Word.

## Instalação

1. Clone este repositório:
   ```bash
   git clone https://github.com/PenhaJV/automacao-para-preencher-contrato-de-prestador-de-servico.git
   ```
2. Navegue até o diretório do projeto:
   ```bash
   cd automacao-para-preencher-contrato-de-prestador-de-servico
   ```
3. Instale as dependências necessárias:
   ```bash
   pip install openpyxl python-docx
   ```

## Uso

1. Coloque sua planilha Excel com os dados dos fornecedores no diretório `data/` e nomeie-a como `fornecedores.xlsx`.
2. Certifique-se de que a planilha tenha uma aba chamada `Sheet1` com as seguintes colunas:
   - Nome do Fornecedor
   - Endereço
   - Cidade
   - Estado
   - CEP
   - E-mail
3. Execute o script:
   ```bash
   python main.py
   ```
4. Os contratos gerados serão salvos no diretório `data/contratos/`.

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir um *pull request* ou relatar problemas.

## Licença

Este projeto está licenciado sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.
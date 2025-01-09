# Automação de Relatorios de Vendas

Este projeto tem como objetivo consolidar dados de vendas, gerar um relatorio em Excel e envia-lo automaticamente por e-mail utilizando Python.

## Funcionalidades
1. **Leitura de Arquivos CSV**: 
   - O script le arquivos de vendas armazenados na pasta `bases/`.
   - Consolida os dados de todos os arquivos em um unico DataFrame.

2. **Processamento e Organizacao dos Dados**:
   - As datas de vendas são convertidas para um formato legivel.
   - Os dados consolidados sao ordenados por data e salvos em um arquivo `Vendas.xlsx`.

3. **Envio Automatico de E-mail**:
   - Utiliza o Outlook para enviar o relatorio por e-mail.
   - O e-mail contem:
     - Destinatario.
     - Assunto dinamico com a data do envio.
     - Corpo do e-mail personalizado.
     - Arquivo `Vendas.xlsx` anexado.

## Tecnologias Utilizadas
- **Python**:
  - `os`: Manipulacao de arquivos e diretorios.
  - `pandas`: Manipulacao e organizacao de dados.
  - `win32com.client`: Automacao do Outlook para envio de e-mails.
  - `datetime`: Manipulacao de datas.

## Como Executar o Projeto
1. Certifique-se de ter os seguintes requisitos instalados:
   - Python 3.x
   - Biblioteca `pandas`
   - Biblioteca `pywin32`

2. Estrutura de Pastas:
   ```
   automacao_planilhas/
   ├── bases/
   │   ├── arquivo1.csv
   │   ├── arquivo2.csv
   └── main.py
   ```

3. Execute o script:
   ```bash
   python main.py
   ```

4. O relatorio sera gerado como `Vendas.xlsx` e enviado automaticamente ao destinatario configurado no codigo.

## Personalizacao
- Para alterar o destinatario, edite o campo `email.To` no codigo.
- Certifique-se de que sua conta do Outlook esta configurada corretamente no computador para o envio automatico.

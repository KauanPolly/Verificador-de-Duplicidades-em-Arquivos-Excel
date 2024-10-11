# Verificador-de-Duplicidades-em-Arquivos-Excel
Verificador de Duplicidades em Arquivos Excel

README (Português)
Verificador de Duplicidades em Arquivos Excel
Este projeto consiste em uma aplicação desenvolvida em Python para verificar duplicidades de datas de vencimento em planilhas Excel. O programa usa uma interface gráfica baseada em Tkinter, permite a conversão de arquivos .xls para .xlsx, e busca por registros duplicados na coluna de datas de vencimento para cada login.

Funcionalidades
Interface gráfica simples usando Tkinter.
Verifica duplicidades de datas de vencimento por login em arquivos Excel.
Suporte a arquivos .xls e .xlsx, com conversão automática de .xls para .xlsx.
Notificação de duplicidades encontradas ou ausência de duplicidades.
Mensagem de carregamento exibida durante o processamento do arquivo.
Requisitos
Python 3.x
Bibliotecas:
pandas
tkinter
openpyxl
xlrd
Você pode instalar as dependências usando o seguinte comando:

bash
Copiar código
pip install pandas tkinter openpyxl xlrd
Como Usar
Clone o repositório:
bash
Copiar código
git clone https://github.com/seu-usuario/nome-do-repositorio.git
Navegue até o diretório do projeto:
bash
Copiar código
cd nome-do-repositorio
Execute o script:
bash
Copiar código
python seu_script.py
Na interface gráfica, clique em "Selecionar Arquivo Excel" e escolha o arquivo Excel desejado.
O programa irá processar o arquivo, exibir uma mensagem de "Carregando..." durante o processamento e informará se duplicidades foram encontradas nas datas de vencimento.
Como Gerar Executável
Para gerar um executável sem console usando PyInstaller e um ícone personalizado, siga estes passos:

Instale o PyInstaller:
bash
Copiar código
pip install pyinstaller
Gere o executável:
bash
Copiar código
pyinstaller --windowed --onefile --icon=caminho/do/seu_icone.ico seu_script.py
O executável será gerado na pasta dist.



Licença

README (English)
Excel File Duplicity Checker
This project consists of a Python application to check for duplicate due dates in Excel spreadsheets. The program uses a Tkinter-based graphical interface, automatically converts .xls files to .xlsx, and searches for duplicate entries in the due date column for each login.

Features
Simple graphical interface using Tkinter.
Checks for duplicate due dates by login in Excel files.
Supports .xls and .xlsx files, with automatic conversion from .xls to .xlsx.
Notifies if duplicates are found or if there are no duplicates.
Loading message displayed during file processing.
Requirements
Python 3.x
Libraries:
pandas
tkinter
openpyxl
xlrd
You can install the dependencies using the following command:

bash
Copiar código
pip install pandas tkinter openpyxl xlrd
How to Use
Clone the repository:
bash
Copiar código
git clone https://github.com/your-username/repository-name.git
Navigate to the project directory:
bash
Copiar código
cd repository-name
Run the script:
bash
Copiar código
python your_script.py
In the graphical interface, click "Select Excel File" and choose the desired Excel file.
The program will process the file, show a "Loading..." message during processing, and notify if duplicates are found in the due dates.
How to Generate Executable
To generate an executable without a console using PyInstaller and a custom icon, follow these steps:

Install PyInstaller:
bash
Copiar código
pip install pyinstaller
Generate the executable:
bash
Copiar código
pyinstaller --windowed --onefile --icon=path/to/your_icon.ico your_script.py
The executable will be generated in the dist folder.




# 🤖 Automação ISS Fortaleza (RPA)

> **Status:** Versão 1.1 (Estável) 🚀

Ferramenta de automação (RPA) desenvolvida em Python para otimizar a rotina fiscal de escritórios de contabilidade. O software realiza o processo completo de extração de dados (ETL) e escrituração no portal da SEFIN Fortaleza.

![Python](https://img.shields.io/badge/Python-3.13-blue?style=for-the-badge&logo=python)
![Selenium](https://img.shields.io/badge/Selenium-Web_Automation-green?style=for-the-badge&logo=selenium)
![Pandas](https://img.shields.io/badge/Pandas-Data_Analysis-150458?style=for-the-badge&logo=pandas)

## 🎯 Funcionalidades

- **Login Múltiplo Automático:** Itera sobre uma lista de empresas via planilha Excel, realizando login seguro.
- **Seletor de Competência (v1.1):** Interface gráfica (GUI) que permite escolher o Mês/Ano alvo, facilitando auditorias e reprocessamentos.
- **Extração de XMLs (ETL):**
  - Download automático de notas fiscais (Serviços Prestados e Tomados).
  - Organização automática de arquivos em pastas padronizadas (`Empresa > Competência`).
- **Paginação Inteligente:** Algoritmo robusto capaz de navegar por múltiplas páginas de notas, detectando automaticamente botões de "Próximo" ou números de página dinâmicos.
- **Escrituração Automática:** Realiza o aceite de notas pendentes e encerra a escrituração do período.
- **Modo Headless:** Opção para executar o robô em segundo plano (invisível), sem ocupar a tela do usuário.

## 🛠️ Tecnologias Utilizadas

- **Linguagem:** Python 3
- **Web Scraping:** Selenium WebDriver (Chrome)
- **Manipulação de Dados:** Pandas
- **Interface Gráfica:** Tkinter (Nativo)
- **Compilação:** PyInstaller (Gerado .exe standalone)

## 📋 Pré-requisitos

1. **Google Chrome** instalado.
2. O software exige uma planilha Excel com as seguintes colunas (exatamente como abaixo):
       NOME        |    CNPJ            |      CPF       | SENHA
Empresa Teste LTDA | 00.000.000/0001-00 | 123.456.789-00 | senha123


## 🚀 Como Usar

### Opção 1: Executável (Usuário Final)
Acesse a aba [Releases](https://github.com/SaulFiuza7/Encerramento_ISSFortaleza/releases) deste repositório e baixe a versão mais recente do `Encerramento ISS For.exe`. 

### Opção 2: Rodando pelo Código (Desenvolvedor)

1. Clone o repositório:
   ```bash
   git clone [https://github.com/SaulFiuza7/Encerramento_ISSFortaleza.git](https://github.com/SaulFiuza7/Encerramento_ISSFortaleza.git)

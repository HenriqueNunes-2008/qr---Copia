#  Projeto QR

Um projeto para geração de QR Codes e registro de horas trabalhadas desenvolvido para a Fleximedical/Kure.

Este repositório contém ferramentas e scripts que permitem gerar QR Codes e registrar informações de horas trabalhadas, com persistência em nuvem, facilitando o uso de QR Codes em sistemas de controle e integração com outros módulos.

---

## 📌 Sobre

Este projeto possui funcionalidades para gerar, registrar e salvar QR Codes, ideal para aplicações como:
* Registro de dados com QR Code
* Integração com sistemas de leitura automática
* Automatização de processos por meio de QR Codes

Este repositório organiza os arquivos de forma simples, utilizando **Python (Flask)** e integração direta com banco de dados.

## 🛠️ Tecnologias e Infraestrutura

O projeto utiliza uma arquitetura moderna para garantir que os dados não sejam perdidos e que a aplicação esteja sempre online:

* **Linguagem:** Python 3.12+
* **Framework Web:** Flask
* **Banco de Dados:** PostgreSQL (gerenciado via SQLAlchemy)
* **Hospedagem:** Magalu Cloud 
* **Planilhas:** Geração de relatórios Excel com `openpyxl`.

## 🧠 Funcionalidades

✔️ Geração dinâmica de QR Codes.
✔️ Registro automático de informações no PostgreSQL.
✔️ Interface web amigável.
✔️ Exportação de dados para Excel.
✔️ Configuração via variáveis de ambiente (Segurança).

## 📁 Estrutura do Repositório

```text
qr/
├── .devcontainer/    # Configurações de desenvolvimento
├── Registro_qr/      # Código principal (app.py, templates, etc.)
├── .gitignore        # Arquivos ignorados pelo Git (como o .env)
├── requirements.txt  # Dependências Python atualizadas
└── README.md         # Documentação do projeto


🌐 Deploy na Magalu Cloud (ou similar)

Para o deploy em um ambiente como a Magalu Cloud, você precisará:
1.  Provisionar uma instância de máquina virtual (VM) com Linux (ex: Ubuntu).
2.  Instalar Python 3, `pip` e as dependências listadas em `requirements.txt`.
3.  Configurar um servidor WSGI (como Gunicorn) para rodar a aplicação Flask.
4.  Configurar um proxy reverso (como Nginx) para gerenciar o tráfego HTTP/HTTPS.
5.  Definir as variáveis de ambiente `DATABASE_URL` e `SECRET_KEY` no ambiente de produção da VM.
6.  Provisionar um serviço de banco de dados PostgreSQL (gerenciado ou instalado na VM) e usar sua URL de conexão em `DATABASE_URL`.


# Changelog do Projeto QR

## Versão 1.2.0 - Reversão para Excel Local (2026-06-11)

Esta versão reverte a integração com o Google Sheets para utilizar arquivos Excel locais, conforme solicitado.

### Principais Mudanças:

-   **Integração de Planilhas:**
    -   Removida a dependência da biblioteca `gspread` e da API do Google Sheets.
    -   Restaurada a funcionalidade de geração e atualização de arquivos Excel (`registros.xlsx`) usando a biblioteca `openpyxl`.
    -   As abas "Registros", "Gráficos" e "Orçamentos" são agora gerenciadas no arquivo Excel local.

## Versão 1.1.0 - Migração para PostgreSQL e Gestão de Usuários (2026-05-21)

Esta versão representa uma refatoração significativa da arquitetura de dados e infraestrutura, movendo a aplicação de um ambiente PaaS (Supabase/Render) para uma solução mais flexível baseada em PostgreSQL e preparada para hospedagem em VM (ex: Magalu Cloud).

### Principais Mudanças:

-   **Migração de Banco de Dados:**
    -   Removida a dependência do Supabase para persistência de dados.
    -   Implementado o uso de **SQLAlchemy** como ORM para interação com um banco de dados **PostgreSQL**.
    -   Criados modelos de banco de dados (`User`, `Area`, `Projeto`, `Funcionario`, `Orcamento`, `Registro`) para gerenciar todas as informações da aplicação.
    -   Os arquivos `employees.json` e `orcamentos.json` foram descontinuados, e seus dados agora são armazenados e gerenciados exclusivamente no banco de dados.
    -   O arquivo `config.json` agora gerencia apenas as configurações de gráficos, enquanto `areas` e `projetos` são gerenciados via banco de dados.

-   **Gestão de Acesso e Usuários:**
    -   Implementado um sistema de cadastro e login para usuários administradores (`User` model).
    -   Novo fluxo de aprovação de usuários: novos cadastros têm status 'pending' e precisam ser ativados por um administrador.
    -   Utilização de `Flask-Bcrypt` para hash seguro de senhas.
    -   Adicionadas rotas para registro de novos usuários (`/register_user`) e para administradores gerenciarem o status dos usuários (`/admin/update_user_status`).

-   **Refatoração do Backend (app.py):**
    -   Todas as operações CRUD (Criar, Ler, Atualizar, Deletar) para áreas, projetos, funcionários e orçamentos foram adaptadas para usar os modelos do SQLAlchemy.
    -   As funções de geração e atualização do arquivo Excel (`criar_planilha_se_nao_existir`, `atualizar_graficos`) agora leem e populam os dados diretamente do banco de dados PostgreSQL.
    -   Configurações sensíveis (`DATABASE_URL`, `SECRET_KEY`) agora são lidas de variáveis de ambiente, conforme boas práticas de segurança para deploy em nuvem.

-   **Preparação para Hospedagem:**
    -   O `README.md` foi atualizado com instruções para deploy em ambientes de VM como a Magalu Cloud, incluindo a necessidade de configurar Gunicorn e Nginx.
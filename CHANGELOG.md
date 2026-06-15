# Changelog do Projeto QR

## Versão 1.4.0 - Edição com Auditoria, Soft Delete e IDs Automáticos (2026-06-15)

Esta versão foca em integridade de dados, rastreabilidade e melhorias na experiência do administrador.

### Principais Mudanças:

-   **Sistema de Soft Delete (Exclusão Lógica):**
    -   Implementação da coluna `ativo` em todas as tabelas principais (`usuarios`, `areas`, `projetos`, `funcionarios`, `orcamentos`).
    -   Registros não são mais apagados fisicamente; a exclusão apenas marca o item como inativo, preservando o histórico de horas e relatórios.
    -   Filtros globais aplicados para garantir que apenas itens ativos apareçam em dropdowns, seletores de QR Code e listas de gerenciamento.
-   **Edição de Registros de Horas com Auditoria:**
    -   Nova funcionalidade que permite ao administrador editar `Data`, `Hora de Início`, `Hora de Fim`, `Área` e `Projeto` de um apontamento existente.
    -   Criação da tabela `historico_alteracoes_horas` para auditoria. Cada edição salva os valores antigos, os novos e o **motivo da alteração** (obrigatório).
    -   Recálculo automático de `total_horas` no servidor após a edição.
-   **Automação de IDs de Funcionários:**
    -   Removida a necessidade de inserir o ID manualmente ao cadastrar funcionários. O sistema agora utiliza o `SERIAL` do PostgreSQL para geração automática.
    -   O Visualizador de QR Codes e o sistema de leitura foram sincronizados para utilizar o ID numérico gerado pelo banco.
-   **Melhorias na Interface Administrativa (UI/UX):**
    -   Remoção da coluna "Ação" na tabela de registros e na exportação Excel, simplificando a visualização.
    -   Ajuste de alinhamento CSS: conteúdos das tabelas agora ficam centralizados sob os títulos das colunas.
    -   Implementação de **Flash Messages (Toasts)** para feedback imediato em ações de sucesso ou erro (ex: "Registro atualizado com sucesso").
-   **Correções de Bugs:**
    -   Correção do erro de JavaScript que impedia a abertura do modal de edição.
    -   Sincronização do modelo SQLAlchemy após a remoção física de colunas no PostgreSQL (remoção do campo `acao`).
    -   Correção na lógica de cálculo de horas para registros que cruzam a meia-noite durante a edição.

---

## Versão 1.3.0 - Filtros Avançados e Refatoração de Registro de Horas (2026-06-12)

Esta versão introduz um fluxo de registro de horas totalmente automatizado via QR Code e aprimora significativamente as funcionalidades de filtragem e visualização de dados nos painéis administrativo e de relatórios.

### Principais Mudanças:

-   **Refatoração do Registro de Horas por QR Code:**
    -   O sistema agora utiliza a leitura de QR Code para iniciar e finalizar atividades automaticamente, eliminando o preenchimento manual de horários.
    -   Registro automático de `ID do funcionário`, `Área`, `Projeto`, `Data atual` e `Timestamp de início` na primeira leitura.
    -   Criação de um `registro aberto` com `status="em_andamento"`.
    -   Na segunda leitura do QR Code, o sistema exibe informações da atividade em andamento e permite a finalização.
    -   Registro automático de `Timestamp de fim` e cálculo de `total_horas` na finalização.
    -   Regras de negócio implementadas: um funcionário só pode ter uma atividade aberta por vez; o cálculo de horas é feito no servidor, garantindo persistência mesmo com o navegador fechado.
-   **Atualização do Modelo de Dados (`Registro`):**
    -   Adição dos campos `status` (VARCHAR) e `total_horas` (FLOAT) ao modelo `Registro`.
    -   O campo `hora_fim` agora é opcional (NULL) para registros em andamento.
    -   *Nota:* Para bancos de dados existentes, foi necessário executar comandos SQL (`ALTER TABLE`) para adicionar as novas colunas e permitir valores nulos em `hora_fim`.
-   **Filtros Avançados no Painel Administrativo:**
    -   Implementação de filtros dinâmicos na aba "Registro de Horas" por `Data` (com presets como "Hoje", "Últimos 7 dias", "Este mês" e "Personalizado"), `Funcionário`, `Área` e `Projeto`.
    -   Suporte a seleção múltipla para `Funcionário`, `Área` e `Projeto`.
    -   Interface de filtros modernizada com dropdowns interativos e contadores de seleção.
    -   A tabela de registros e o total de horas filtradas são atualizados dinamicamente, mantendo o usuário na aba "Registro de Horas" após a aplicação dos filtros.
-   **Filtros Avançados na Área do Relator:**
    -   Adição de filtros por `Área` e `Projeto` (com seleção múltipla) na página de relatórios.
    -   Os gráficos de barras e de rosca, bem como o resumo de horas, são atualizados automaticamente com base nos filtros selecionados.
-   **Melhorias de Código e Correções:**
    -   A função `_get_processed_report_data` foi refatorada para aceitar e aplicar os novos filtros de forma eficiente.
    -   Correções de erros de JavaScript no frontend para garantir a correta renderização da interface após a leitura do QR Code e a aplicação de filtros.
    -   Remoção de seções duplicadas no `admin.html` e correção de pequenos erros de sintaxe CSS/JavaScript.


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
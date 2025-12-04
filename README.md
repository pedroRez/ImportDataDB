# ImportDataDB

Projeto desktop em Python + Qt (PySide6) para importar planilhas Excel (XLSX) e mapear colunas para tabelas de banco de dados, começando por PostgreSQL.

## Preparação do ambiente no Windows
1. **Instalar Python 3.11+** (https://www.python.org/downloads/). Marcar a opção “Add Python to PATH”.
2. **Instalar Git** (https://git-scm.com/download/win) para clonar o repositório.
3. **Instalar PostgreSQL** (https://www.postgresql.org/download/) e criar um usuário com permissões de leitura/escrita. Opcional: instalar o pgAdmin para gerenciar o banco.
4. **Instalar Visual C++ Redistributable** (https://aka.ms/vs/17/release/vc_redist.x64.exe). Necessário para algumas dependências nativas.
5. **(Opcional, recomendado)** Instalar **Microsoft Build Tools** caso seja preciso compilar extensões nativas (https://visualstudio.microsoft.com/visual-cpp-build-tools/).
6. Verificar que o `pip` está atualizado:
   ```powershell
   python -m pip install --upgrade pip
   ```
7. Criar e ativar um ambiente virtual na pasta do projeto:
   ```powershell
   python -m venv .venv
   .\.venv\Scripts\activate
   ```
8. Instalar dependências principais (versões mínimas sugeridas):
   ```powershell
   pip install "PySide6>=6.6" "pandas>=2.2" "SQLAlchemy>=2.0" "openpyxl>=3.1" "psycopg2-binary>=2.9"
   ```
   - Se for necessário suporte a outros bancos no futuro, adicionar os respectivos drivers (ex.: `pyodbc`, `pymysql`).

## Como usar (primeira execução sugerida)
1. Clonar o repositório e entrar na pasta:
   ```powershell
   git clone <URL-do-repositorio>
   cd ImportDataDB
   ```
2. Ativar o ambiente virtual e instalar as dependências (passos acima).
3. Configurar um banco PostgreSQL de teste e criar uma tabela exemplo para mapeamento.
4. Rodar a aplicação (quando a UI estiver implementada, o entrypoint será algo como `python -m importdatadb.app`).
5. Fluxo esperado na UI:
   - Escolher o arquivo Excel (.xlsx) e visualizar as abas.
   - Selecionar a aba e a tabela do banco.
   - Indicar linha de cabeçalho e faixa de dados.
   - Mapear colunas da planilha ↔ colunas da tabela; definir se a PK é autoincrement.
   - Escolher operação (INSERT ou UPDATE) e, para UPDATE, escolher o campo de junção (PK ou outro campo).
   - Pré-visualizar e confirmar a execução.

## Próximos passos sugeridos
- Definir estrutura de pacotes (ex.: `importdatadb/ui`, `importdatadb/db`, `importdatadb/excel`).
- Adicionar scripts de conveniência (`make` ou `Invoke-Task`) para setup, lint e testes.
- Documentar formatos de log e estratégia de rollback em caso de falhas parciais.

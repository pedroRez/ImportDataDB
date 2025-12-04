# ImportDataDB

Projeto desktop em Python + Qt (PySide6) para importar planilhas Excel (XLSX) e mapear colunas para tabelas de banco de dados, começando por PostgreSQL.

## Preparação do ambiente no Windows
1. **Instalar Python 3.11+** (https://www.python.org/downloads/). Marcar a opção “Add Python to PATH”.
2. **Instalar Git** (https://git-scm.com/download/win) para clonar o repositório.
3. **Instalar PostgreSQL** (https://www.postgresql.org/download/) e criar um usuário com permissões de leitura/escrita. Opcional: instalar o pgAdmin para gerenciar o banco.
4. **Instalar Visual C++ Redistributable** (https://aka.ms/vs/17/release/vc_redist.x64.exe). Necessário para algumas dependências nativas.
5. **(Opcional, recomendado)** Instalar **Microsoft Build Tools** caso seja preciso compilar extensões nativas (https://visualstudio.microsoft.com/visual-cpp-build-tools/).
6. Verificar que o `pip` está atualizado (CMD):
   ```cmd
   python -m pip install --upgrade pip
   ```
7. Criar e ativar um ambiente virtual na pasta do projeto (CMD):
   ```cmd
   python -m venv .venv
   .\.venv\Scripts\activate
   ```
8. Instalar dependências principais (versões mínimas sugeridas) a partir do `requirements.txt` (CMD):
   ```cmd
   pip install -r requirements.txt
   ```
   - Se for necessário suporte a outros bancos no futuro, adicionar os respectivos drivers (ex.: `pyodbc`, `pymysql`).

## Como usar (primeira execução sugerida)
1. Clonar o repositório e entrar na pasta (todos os comandos abaixo partem da raiz que contém `LICENSE`, `README.md`, `requirements.txt` e a pasta `src`):
   ```cmd
   git clone <URL-do-repositorio>
   cd ImportDataDB
   ```
2. Ativar o ambiente virtual e instalar as dependências (passos acima).
3. Configurar um banco PostgreSQL de teste e criar uma tabela exemplo para mapeamento.
4. Rodar a aplicação. O código-fonte agora fica diretamente em `src`, então basta apontar o `PYTHONPATH` para essa pasta e executar o módulo `app`:
   ```cmd
   set PYTHONPATH=%CD%\src
   python -m app
   ```
   Em shells Unix-like, use:
   ```bash
   export PYTHONPATH="$(pwd)/src"
   python -m app
   ```
5. Fluxo esperado na UI:
   - Escolher o arquivo Excel (.xlsx ou .xlsm) e visualizar as abas.
   - Selecionar a aba e a tabela do banco.
   - Indicar linha de cabeçalho e faixa de dados.
   - Mapear colunas da planilha ↔ colunas da tabela; definir se a PK é autoincrement.
   - Escolher operação (INSERT ou UPDATE) e, para UPDATE, escolher o campo de junção (PK ou outro campo).
   - Pré-visualizar e confirmar a execução.

## Próximos passos sugeridos
- Consolidar a organização interna iniciada em `src/` (ex.: `src/ui`, `src/db`, `src/excel`).
- Adicionar scripts de conveniência (`make` ou `Invoke-Task`) para setup, lint e testes.
- Documentar formatos de log e estratégia de rollback em caso de falhas parciais.

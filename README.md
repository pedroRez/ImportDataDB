# ImportDataDB

Projeto desktop em Python + Qt (PySide6) para importar planilhas Excel (XLSX) e mapear colunas para tabelas de banco de dados, começando por PostgreSQL.

## Tarefas iniciais do projeto
### 1. Fundamentos e estrutura
- [ ] Preparar ambiente virtual Python e instalar dependências básicas (`PySide6`, `pandas`, `SQLAlchemy`, `openpyxl`, `psycopg2-binary`).
- [ ] Definir a estrutura de pacotes (ex.: `importdatadb/ui`, `importdatadb/db`, `importdatadb/excel`, `importdatadb/core`).
- [ ] Configurar logging centralizado e arquivos de configuração (.env ou dotenv) para credenciais de banco.

### 2. Camada de dados (PostgreSQL primeiro)
- [ ] Criar interface de provedor de banco (ex.: `DatabaseProvider`) para permitir futura troca de SGBD.
- [ ] Implementar provedor PostgreSQL (conexão, listagem de tabelas, leitura de metadados de colunas e PKs).
- [ ] Implementar conversão de tipos básicos entre pandas/Excel e PostgreSQL (datas, números, texto, booleanos).

### 3. Leitura de Excel
- [ ] Implementar leitor de planilhas: listar abas, identificar linhas com valores, detectar linha de cabeçalho.
- [ ] Permitir seleção de faixa (linha inicial/final) e cabeçalho escolhido pelo usuário.
- [ ] Expor lista de colunas da aba selecionada para posterior mapeamento.

### 4. UI (PySide6)
- [ ] Tela inicial com: seleção de arquivo Excel, conexão ao banco (host, porta, database, usuário, senha) e teste de conexão.
- [ ] Passo de seleção: exibir abas do Excel (lado esquerdo) e tabelas do banco (lado direito) para mapear aba ↔ tabela.
- [ ] Passo de cabeçalho/faixa: permitir escolher linha de cabeçalho e intervalo de dados da aba.
- [ ] Passo de mapeamento: mostrar lista de colunas da planilha e colunas da tabela; permitir mapear uma a uma, incluindo opção de autoincremento quando a PK não for mapeada.
- [ ] Passo de operação: escolher entre INSERT ou UPDATE; para UPDATE, selecionar coluna de junção (PK ou outra) para compor o WHERE.
- [ ] Passo de pré-visualização: exibir SQL gerado, amostra de dados e possíveis avisos (tipos, nulos, truncamentos).
- [ ] Passo de execução: rodar em transação, exibir progresso e registrar log/sumário.

### 5. Lógica de importação
- [ ] Implementar geração de comandos INSERT em lote (com parametrização e commit controlado).
- [ ] Implementar geração de UPDATE com cláusula WHERE baseada no campo de junção escolhido.
- [ ] Validar obrigatórios, tamanhos de campo e conversões; oferecer fallback para autoincremento quando PK não vier da planilha.
- [ ] Tratar erros por linha e registrar em log (para reprocessamento ou auditoria).

### 6. Qualidade e automação
- [ ] Adicionar testes unitários para o mapeamento, geração de SQL e conversão de tipos.
- [ ] Incluir ferramentas de lint/format (ex.: `ruff`, `black`, `mypy`) e scripts de automação (Makefile ou `invoke`).
- [ ] Configurar pipeline de CI local (ex.: GitHub Actions) para lint e testes.

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

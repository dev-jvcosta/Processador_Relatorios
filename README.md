# Processador Unificado de Relatórios

Este projeto é uma ferramenta de automação desenvolvida em Python com interface gráfica (Tkinter). O objetivo é processar, consolidar e gerar relatórios gerenciais a partir de diversas fontes de dados (Excel) relacionadas a abastecimento, motoristas, rankings e turnos.

## 📋 Funcionalidades

O script processa e gera os seguintes tipos de relatórios:

* **Abst_Mot_Por_empresa:** Integração de dados de abastecimento e motoristas por empresa, distribuindo proporcionalmente os valores de KM e litros baseado nos horários de trabalho.

* **Ranking_Por_Empresa:** Consolidação de rankings de performance com múltiplas abas organizadas por linha, turno e status.

* **Ranking_Integração:** Relatórios integrados de performance combinando dados de ranking, turnos e abastecimento.

* **Ranking_Ouro_Mediano:** Consolidação específica para faixas de pontuação (Fase: Ouro/Ouro C, Status: Mediano, Ponto: 3.97-3.99).

* **Ranking_Km_Proporcional:** Cálculos de KM distribuídos proporcionalmente baseados nos totais de abastecimento da empresa.

* **Turnos_Integração:** Análise de dados baseada em turnos de trabalho (Madrugada, Manhã, Intervalo, Tarde, Noite).

* **Resumo_Motorista_Cliente:** Métricas consolidadas por cliente e motorista com distribuição proporcional de KM e litros.

* **Relatório de Motoristas Insuficientes (RPP_Insuficientes):** Consolidação de relatórios de ranking por empresa em um único arquivo Excel, cruzando dados de múltiplas empresas para um período específico (Ano/Mês).

## 🛠️ Pré-requisitos

* **Python 3.7 ou superior** instalado (recomendado Python 3.9+)
* **Tkinter** (geralmente incluído com Python, mas pode precisar de instalação separada em alguns sistemas Linux)
* Uma IDE de sua preferência (Cursor, VS Code, PyCharm, etc.) ou terminal/linha de comando

### Verificação do Python

Para verificar se o Python está instalado, execute no terminal:

**Windows:**
```bash
python --version
# ou
py --version
```

**Mac/Linux:**
```bash
python3 --version
```

## 🚀 Instalação e Configuração

Siga os passos abaixo para configurar o ambiente de desenvolvimento em uma nova máquina.

### 1. Clonar ou Baixar o Projeto

Certifique-se de que os seguintes arquivos estejam na pasta raiz do projeto:
- `main.py` (arquivo principal)
- `README.md` (este arquivo)
- `requirements.txt` (será criado no passo 4)

### 2. Configurar o Ambiente Virtual (Venv)

O ambiente virtual isola as dependências do projeto, evitando conflitos com outros projetos Python.

**Método Manual (Recomendado para primeira instalação):**

**Para Windows:**

Se o comando `py -m venv venv` apresentar erro sobre executável não encontrado, use uma das alternativas:

**Opção 1 - Usar o Python diretamente (Recomendado):**
```bash
python -m venv venv
```

**Opção 2 - Usar caminho completo do Python:**
```bash
"C:\Program Files\Python314\python.exe" -m venv venv
```

**Opção 3 - Usar py launcher:**
```bash
py -m venv venv
```

**Para Mac/Linux:**
```bash
python3 -m venv venv
```

> **Nota:** Se durante a criação do ambiente virtual aparecer um aviso do tipo "Could not find platform independent libraries <prefix>" ou "did not find executable", tente usar o caminho completo do Python ou o comando `python` diretamente. O ambiente virtual será criado corretamente.

### 3. Ativar o Ambiente Virtual

**Windows (PowerShell ou CMD):**
```bash
.\venv\Scripts\activate
```

**Windows (Git Bash):**
```bash
source venv/Scripts/activate
```

**Mac / Linux:**
```bash
source venv/bin/activate
```

> **Importante:** Ao ativar, você verá `(venv)` no início da linha do terminal. Isso indica que o ambiente virtual está ativo. **SEMPRE ative o ambiente virtual antes de instalar dependências ou executar o script.**

### 4. Instalar Dependências

Crie um arquivo `requirements.txt` na raiz do projeto com o seguinte conteúdo:

```plaintext
pandas>=1.3.0
openpyxl>=3.0.0
reportlab>=3.6.0
numpy>=1.21.0
sv-ttk>=2.0.0
darkdetect>=0.8.0
```

> **Nota:** As bibliotecas `sv-ttk` e `darkdetect` são necessárias para a interface gráfica moderna com suporte a tema claro/escuro automático.

Em seguida, com o ambiente virtual **ativado**, execute:

```bash
pip install -r requirements.txt
```

**Verificação da Instalação:**

Para verificar se todas as dependências foram instaladas corretamente:

```bash
pip list
```

Você deve ver todas as bibliotecas listadas acima na lista de pacotes instalados.

### 5. Preparar Estrutura de Pastas

Antes de executar o programa, você precisa criar a estrutura de pastas para os arquivos de entrada.

Crie um diretório base (por exemplo: `D:\Scripts\Entrada` ou `C:\Dados\Entrada`) e dentro dele crie as seguintes subpastas:

```
Diretório_Base_Entrada/
│
├── Integração_Abast/          # Arquivos de abastecimento
│   └── Abastecimento_[Empresa]_[Mês]_[Ano].xlsx
│
├── Integração_Mot/             # Arquivos de motoristas
│   └── Motorista_[Empresa]_[Mês]_[Ano].xlsx
│
├── Ranking/                     # Arquivos de ranking
│   └── Ranking_[Empresa]_[Mês]_[Ano].xlsx
│
├── Turnos_128/                  # Arquivos de turnos
│   └── Turnos_128_[Empresa]_[Mês]_[Ano].xlsx
│
└── Resumo_Motorista_Cliente/    # Arquivos de resumo
    └── RMC_[Empresa]_[Mês]_[Ano].xlsx
```

**Formato dos Arquivos:**

- **Abastecimento:** `Abastecimento_[Empresa]_[Mês]_[Ano].xlsx` (ex: `Abastecimento_Amparo_Agosto_2025.xlsx`)
- **Motorista:** `Motorista_[Empresa]_[Mês]_[Ano].xlsx` (ex: `Motorista_Amparo_Agosto_2025.xlsx`)
- **Ranking:** `Ranking_[Empresa]_[Mês]_[Ano].xlsx` (ex: `Ranking_Amparo_Agosto_2025.xlsx`)
- **Turnos:** `Turnos_128_[Empresa]_[Mês]_[Ano].xlsx` (ex: `Turnos_128_Amparo_Agosto_2025.xlsx`)
- **Resumo:** `RMC_[Empresa]_[Mês]_[Ano].xlsx` (ex: `RMC_Amparo_Agosto_2025.xlsx`)

## ▶️ Como Executar

### Execução Básica

Com o ambiente virtual **ativado**, execute o comando abaixo no terminal:

**Windows:**
```bash
python main.py
# ou
py main.py
```

**Mac / Linux:**
```bash
python3 main.py
```

A interface gráfica será aberta automaticamente.

## 📖 Guia de Uso do Usuário

### Primeira Execução

1. **Iniciar o Programa:**
   - Ative o ambiente virtual (veja seção 3 da instalação)
   - Execute `python main.py` (ou `py main.py` no Windows)

2. **Configurar Diretórios:**
   - **Diretório Base dos Arquivos de Entrada:** Clique em "Procurar" e selecione a pasta onde você criou a estrutura de subpastas (ex: `D:\Scripts\Entrada`)
   - **Diretório Base dos Arquivos de Saída:** Clique em "Procurar" e selecione onde deseja salvar os relatórios gerados (ex: `D:\Scripts\Saida`)

3. **Configurar Versão (Opcional):**
   - O campo "Versão" permite adicionar um sufixo aos arquivos gerados
   - Exemplos: `_v1`, `_2.0`, `_teste`
   - Se deixar em branco, os arquivos serão gerados sem sufixo
   - Você pode digitar manualmente ou selecionar uma opção pré-definida no dropdown

4. **Selecionar Tipos de Relatório:**
   - Marque os checkboxes dos tipos de relatório que deseja processar
   - Você pode selecionar múltiplos tipos simultaneamente

5. **⚠️ IMPORTANTE - Atualizar Lista:**
   - **SEMPRE clique no botão "Atualizar"** após:
     - Alterar os diretórios de entrada/saída
     - Alterar a versão
     - Adicionar novos arquivos nas pastas de entrada
     - Alterar a seleção de tipos de relatório
   - O botão "Atualizar" recarrega a lista de empresas e períodos disponíveis
   - **NUNCA use "Processar Tudo" ou "Processar Todos os Períodos" sem antes clicar em "Atualizar"**

6. **Selecionar Empresas e Períodos:**
   - Após clicar em "Atualizar", a lista de empresas será preenchida automaticamente
   - Selecione uma ou mais empresas na lista à esquerda
   - Selecione os anos e meses desejados nas listas à direita
   - Os períodos disponíveis são filtrados automaticamente baseados nas empresas selecionadas

### Processamento

#### Opção 1: Processar Selecionados

1. Selecione empresas, anos e meses específicos
2. Clique em **"Processar Selecionados"**
3. Apenas os períodos selecionados serão processados

#### Opção 2: Processar Todos os Períodos (para empresas selecionadas)

1. **⚠️ IMPORTANTE:** Clique primeiro em **"Atualizar"** para recarregar a lista
2. Selecione uma ou mais empresas
3. (Opcional) Selecione anos/meses para filtrar
4. Clique em **"Processar Todos os Períodos"**
5. Todos os períodos disponíveis para as empresas selecionadas serão processados

#### Opção 3: Processar Todas as Empresas

1. **⚠️ IMPORTANTE:** Clique primeiro em **"Atualizar"** para recarregar a lista
2. (Opcional) Selecione anos/meses para filtrar
3. Clique em **"Processar Todas as Empresas"**
4. Todas as empresas e seus períodos disponíveis serão processados

#### Opção 4: Processar Tudo

1. **⚠️ IMPORTANTE:** Clique primeiro em **"Atualizar"** para recarregar a lista
2. Clique em **"Processar Tudo"** (botão na seção de tipos de relatório)
3. Todos os tipos de relatório selecionados, todas as empresas e todos os períodos serão processados

### Botões Especiais

- **Consolidar Ouro Mediano:** Processa a consolidação específica de registros Ouro Mediano
- **Processar Ranking_Km_Proporcional:** Processa apenas o tipo Ranking_Km_Proporcional
- **Gerar Relatório Insuficientes:** Abre um modal para gerar o relatório consolidado de motoristas insuficientes (ver seção dedicada abaixo)
- **Atualizar:** Recarrega a lista de empresas e períodos disponíveis (use sempre antes de processar em lote)

### Gerar Relatório de Motoristas Insuficientes

Esta funcionalidade permite consolidar os relatórios `Ranking_Por_Empresa` de todas as empresas em um único arquivo Excel.

**Como usar:**

1. Clique no botão **"Gerar Relatório Insuficientes"**
2. Na janela modal que se abre, configure:
   - **Caminho Ranking_Por_Empresa:** Informe o caminho absoluto até a pasta `Ranking_Por_Empresa` (ou clique em "Procurar" para selecionar)
   - **Ano:** Informe o ano desejado (ex: 2025)
   - **Mês:** Selecione o mês no dropdown (ex: Novembro)
3. Clique em **"Gerar Relatório"**

**Estrutura esperada de entrada:**
```
Ranking_Por_Empresa/
├── Alpha/
│   └── 2025/
│       └── Novembro/
│           └── Ranking_Por_Empresa_Alpha_Novembro_2025.xlsx
├── Amparo/
│   └── 2025/
│       └── Novembro/
│           └── Ranking_Por_Empresa_Amparo_Novembro_2025.xlsx
└── [Outras Empresas]/
    └── ...
```

**Arquivo de saída gerado:**
- **Localização:** `RPP_Insuficientes/Relatório_Por_Empresa_Insuficientes.xlsx`
- **Estrutura do Excel:**
  - **Aba "Todas As Empresas":** Consolida os dados de todas as empresas em uma única aba. Os dados de cada empresa são separados por uma linha em branco, com cabeçalhos repetidos.
  - **Abas individuais por empresa:** Uma aba para cada empresa (ex: "Alpha", "Amparo") contendo os dados completos do relatório original.

**Tratamento de erros:**
- Se uma empresa não tiver a pasta do Ano/Mês especificado, ela é ignorada e um aviso é registrado no log
- Se a pasta existir mas não contiver arquivo `.xlsx`, um aviso é exibido
- O processamento continua para as demais empresas mesmo se houver erros em algumas

### Acompanhamento do Processamento

1. **Barra de Progresso:** Mostra o progresso geral do processamento
2. **Status:** Exibe a tarefa atual sendo executada
3. **Log de Processamento:** Mostra mensagens detalhadas sobre cada etapa
   - ✅ Verde: Sucesso
   - ❌ Vermelho: Erro
   - ⚠️ Laranja: Aviso
   - ℹ️ Azul: Informação

### Geração de Relatório PDF

1. Após o processamento, clique em **"Gerar PDF do Relatório"**
2. Selecione onde deseja salvar o PDF
3. O PDF conterá:
   - Informações gerais do processamento
   - Estatísticas (sucessos, erros, avisos)
   - Log completo de processamento

### Limpar Log

- Clique em **"Limpar Log"** para limpar o log de processamento e começar uma nova sessão

## 📂 Estrutura de Pastas de Saída

Os relatórios gerados são organizados automaticamente na seguinte estrutura:

```
Diretório_Saída/
│
├── Abst_Mot_Por_empresa/
│   └── [Empresa]/
│       └── [Ano]/
│           └── [Mês]/
│               ├── Detalhado_[Empresa]_[Mês]_[Ano][Versão].xlsx
│               └── Abst_Mot_Por_empresa_[Empresa]_[Mês]_[Ano][Versão].xlsx
│
├── Ranking_Por_Empresa/
│   └── [Empresa]/
│       └── [Ano]/
│           └── [Mês]/
│               └── Ranking_Por_Empresa_[Empresa]_[Mês]_[Ano][Versão].xlsx
│
├── Ranking_Integração/
│   └── [Empresa]/
│       └── [Ano]/
│           └── [Mês]/
│               └── Ranking_Integração_[Empresa]_[Mês]_[Ano][Versão].xlsx
│
├── Ranking_Ouro_Mediano/
│   └── Ranking_Ouro_Mediano_[Período_Inicial]_a_[Período_Final][Versão].xlsx
│
├── Rankig_Km_Proporcional/
│   └── [Empresa]/
│       └── [Ano]/
│           └── [Mês]/
│               ├── Detalhado_[Empresa]_[Mês]_[Ano][Versão].xlsx
│               ├── Consolidado_[Empresa]_[Mês]_[Ano][Versão].xlsx
│               └── Ranking_Km_Proporcional_[Empresa]_[Mês]_[Ano][Versão].xlsx
│
├── Turnos Integração/
│   └── [Empresa]/
│       └── [Ano]/
│           └── [Mês]/
│               └── Turnos_Integração_[Empresa]_[Mês]_[Ano][Versão].xlsx
│
├── RMC_Destribuida/
│   └── [Empresa]/
│       └── 2025/
│           └── [Mês]/
│               └── RMC_Km_l_Distribuida_[Empresa]_[Mês]_[Ano][Versão].xlsx
│
└── RPP_Insuficientes/
    └── Relatório_Por_Empresa_Insuficientes.xlsx
```

## ⚠️ Solução de Problemas Comuns

### Erro de Permissão (PermissionError)

**Problema:** O script não consegue salvar arquivos Excel.

**Soluções:**
1. Certifique-se de que nenhum arquivo Excel gerado anteriormente esteja aberto no Excel
2. Verifique se você tem permissões de escrita na pasta de saída
3. Feche todas as instâncias do Excel antes de executar o script
4. O script tentará criar arquivos com nomes alternativos se detectar arquivos em uso

### Caminho não encontrado

**Problema:** Mensagem de erro indicando que pastas não foram encontradas.

**Soluções:**
1. Verifique se as pastas de entrada estão nomeadas **exatamente** como descrito:
   - `Integração_Abast` (com acento)
   - `Integração_Mot` (com acento)
   - `Ranking` (sem acento)
   - `Turnos_128` (com underscore e número)
   - `Resumo_Motorista_Cliente` (com underscores)
2. Verifique se os arquivos estão dentro das pastas corretas
3. Verifique se os nomes dos arquivos seguem o padrão esperado

### Erro no pip install

**Problema:** Erro ao instalar dependências.

**Soluções:**
1. Certifique-se de que o ambiente virtual está **ativado** (deve aparecer `(venv)` no terminal)
2. Atualize o pip: `python -m pip install --upgrade pip`
3. Tente instalar as dependências uma por uma para identificar qual está causando problema
4. No Windows, execute o terminal como Administrador

### Tkinter não encontrado

**Problema:** Erro `ModuleNotFoundError: No module named 'tkinter'`

**Soluções:**

**Windows:**
- Reinstale o Python marcando a opção "tcl/tk and IDLE" durante a instalação

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get install python3-tk
```

**Linux (Fedora):**
```bash
sudo dnf install python3-tkinter
```

**Mac:**
- O Tkinter geralmente vem pré-instalado. Se não, reinstale o Python do python.org

### Interface não abre ou aparece em branco

**Problema:** A janela abre mas não mostra conteúdo ou não abre.

**Soluções:**
1. Verifique se todas as dependências foram instaladas, especialmente `sv-ttk` e `darkdetect`
2. Verifique o arquivo de log `unified_processing.log` para mensagens de erro
3. Tente executar o script diretamente no terminal para ver mensagens de erro

### Lista de empresas vazia

**Problema:** Após clicar em "Atualizar", nenhuma empresa aparece na lista.

**Soluções:**
1. Verifique se os arquivos estão nas pastas corretas
2. Verifique se os nomes dos arquivos seguem o padrão esperado
3. Verifique se pelo menos um tipo de relatório está marcado
4. Verifique o log de processamento para mensagens de erro específicas

### Processamento muito lento

**Problema:** O processamento demora muito tempo.

**Soluções:**
1. Processe em lotes menores (selecione menos empresas/períodos por vez)
2. Feche outros programas que possam estar usando recursos do sistema
3. Verifique se há muitos arquivos grandes sendo processados simultaneamente

## 📝 Logs e Arquivos de Log

O programa gera automaticamente um arquivo de log chamado `unified_processing.log` na pasta raiz do projeto. Este arquivo contém:

- Timestamp de cada operação
- Nível de log (INFO, WARNING, ERROR)
- Mensagens detalhadas sobre o processamento
- Erros e exceções

Use este arquivo para diagnosticar problemas quando a interface não fornecer informações suficientes.

## 🔄 Fluxo de Trabalho Recomendado

1. **Preparação:**
   - Organize os arquivos Excel nas pastas corretas
   - Verifique se os nomes dos arquivos seguem o padrão esperado

2. **Configuração Inicial:**
   - Abra o programa
   - Configure os diretórios de entrada e saída
   - Configure a versão (se necessário)
   - Selecione os tipos de relatório desejados

3. **Atualização:**
   - **SEMPRE clique em "Atualizar"** antes de processar em lote

4. **Seleção:**
   - Selecione empresas e períodos específicos

5. **Processamento:**
   - Escolha o método de processamento adequado
   - Acompanhe o progresso pelo log

6. **Verificação:**
   - Verifique os arquivos gerados na pasta de saída
   - Revise o log para identificar possíveis problemas

## 📊 Resumo da Estrutura Final do Projeto

```
Processador_Relatorios/
│
├── venv/                          # Ambiente virtual (gerado automaticamente)
│   ├── Scripts/                   # (Windows) ou bin/ (Linux/Mac)
│   └── ...
│
├── Integração_Abast/              # (Criar manualmente - arquivos de entrada)
├── Integração_Mot/                 # (Criar manualmente - arquivos de entrada)
├── Ranking/                        # (Criar manualmente - arquivos de entrada)
├── Turnos_128/                     # (Criar manualmente - arquivos de entrada)
├── Resumo_Motorista_Cliente/       # (Criar manualmente - arquivos de entrada)
│
├── main.py                         # Script principal
├── requirements.txt                # Lista de dependências
├── README.md                       # Este arquivo
├── unified_processing.log          # Arquivo de log (gerado automaticamente)
│
└── [Diretório de Saída]/          # (Configurado na interface)
    └── [Estrutura de pastas gerada automaticamente]
```

## 🎯 Checklist de Instalação

Use este checklist para garantir que tudo está configurado corretamente:

- [ ] Python 3.7+ instalado e funcionando
- [ ] Ambiente virtual criado (`venv/` existe)
- [ ] Ambiente virtual ativado (aparece `(venv)` no terminal)
- [ ] Todas as dependências instaladas (`pip list` mostra todas as bibliotecas)
- [ ] Estrutura de pastas de entrada criada
- [ ] Arquivos Excel organizados nas pastas corretas
- [ ] Nomes dos arquivos seguem o padrão esperado
- [ ] Diretório de saída configurado e com permissões de escrita
- [ ] Programa executado com sucesso (`python main.py`)
- [ ] Interface gráfica abre corretamente
- [ ] Botão "Atualizar" funciona e lista empresas/períodos

## 📞 Suporte

Se encontrar problemas não listados aqui:

1. Verifique o arquivo `unified_processing.log` para mensagens de erro detalhadas
2. Verifique se todas as dependências estão instaladas: `pip list`
3. Verifique se o Python está na versão correta: `python --version`
4. Tente executar o script diretamente no terminal para ver mensagens de erro

---

**Desenvolvido para automação de processos internos.**

**Versão do Documento:** 2.0  
**Última Atualização:** 2025

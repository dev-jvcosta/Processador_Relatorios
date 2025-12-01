# Processador Unificado de Relatórios

Este projeto é uma ferramenta de automação desenvolvida em Python com interface gráfica (Tkinter). O objetivo é processar, consolidar e gerar relatórios gerenciais a partir de diversas fontes de dados (Excel) relacionadas a abastecimento, motoristas, rankings e turnos.

## 📋 Funcionalidades

O script processa e gera os seguintes tipos de relatórios:
* **Abst_Mot_Por_empresa:** Integração de dados de abastecimento e motoristas.

* **Ranking_Por_Empresa:** Consolidação de rankings de performance.

* **Ranking_Integração:** Relatórios integrados de performance.

* **Ranking_Ouro_Mediano:** Consolidação específica para faixas de pontuação.

* **Ranking_Km_Proporcional:** Cálculos de KM distribuídos proporcionalmente.

* **Turnos_Integração:** Análise de dados baseada em turnos de trabalho.

* **Resumo_Motorista_Cliente:** Métricas consolidadas por cliente e motorista.

## 🛠️ Pré-requisitos

* **Python 3.x** instalado.

* Uma IDE de sua preferência (Cursor, VS Code, PyCharm, etc.).

## 🚀 Instalação e Configuração

Siga os passos abaixo para configurar o ambiente de desenvolvimento.

### 1. Clonar ou Baixar o Projeto
Certifique-se de que o arquivo `main.py` e este `README.md` estejam na pasta raiz do projeto.

### 2. Configurar o Ambiente Virtual (Venv)

Para evitar conflitos de bibliotecas e garantir que o executável correto do Python seja utilizado (especialmente no Windows), você pode usar o script automatizado ou os comandos manuais:

**Método Rápido (Recomendado):**

**Windows:**
```bash
# Execute o arquivo .bat diretamente (duplo clique ou no terminal)
setup_venv.bat

# OU no terminal/PowerShell:
.\setup_venv.bat
```

**Mac/Linux:**
```bash
bash setup_venv.sh
# OU
chmod +x setup_venv.sh && ./setup_venv.sh
```

> **Importante:** No Windows, execute `setup_venv.bat` (não `.sh`). O arquivo `.sh` é apenas para Linux/Mac.

**Método Manual:**

Siga os comandos abaixo no terminal da sua IDE:

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

* Windows
```bash
.\venv\Scripts\activate
```

* Mac / Linux
```bash
source venv/bin/activate
```

> Ao ativar, você verá ```(venv)``` no início da linha do terminal

### 4. Instalar Dependências

Crie um arquivo ```requirements.txt``` (se não existir) com o seguinte conteúdo:

```plaintext
pandas
openpyxl
reportlab
numpy
```

Em seguida, execute:
```plaintext
pip install -r requirements.txt
```

## 📂 Estrutura de Pastas Exigida

Para que o processamento funcione, você deve criar as seguintes subpastas dentro do diretório que será selecionado como "Entrada" na interface gráfica, e colocar os respectivos arquivos Excel nelas:

* ``` Integração_Abast ```
* ``` Integração_Mot ```
* ``` Ranking ```
* ``` Turnos_128 ```
* ``` Resumo_Motorista_Cliente ```

## ▶️ Como Executar

Com o ambiente virtual ativo, execute o comando abaixo no terminal:

**Windows**

```bash
py main.py
```

**Mac / Linux**

```bash
python3 main.py
```

**Configuação na interface**

1. Diretório Base: Selecione a pasta onde você criou a estrutura de subpastas acima.
2. Diretório de Saída: Selecione onde deseja salvar os relatórios gerados.
3. Versão: (Opcional) Adicione um sufixo para os arquivos (ex: _v1).
4. Selecione a Empresa e o Período desejados.
5. Clique em Processar Selecionados ou Processar Tudo.

## ⚠️ Solução de Problemas Comuns
* **Erro de Permissão (PermissionError):** Certifique-se de que nenhum arquivo Excel gerado anteriormente esteja aberto no Excel enquanto o script roda.
* **Caminho não encontrado:** Verifique se as pastas de entrada estão nomeadas exatamente como descrito na seção "Estrutura de Pastas".
* **Erro no** ``` pip install```**:** Verifique se o (venv) está ativo antes de instalar.

---

Desenvolvido para automação de processos internos.

## Resumo da Estrutura Final

```Plaintext
Processador_Relatorios/
│
├── venv/                  # (Pasta do ambiente virtual - gerada automaticamente)
├── Integração_Abast/      # (Pasta criada por você para os arquivos Excel)
├── ... (Outras pastas)
├── main.py                # (Seu script principal)
└── requirements.txt       # (Lista de bibliotecas)
```

# Análise de Relatórios - Processador de Relatórios

## 📋 Lista de Relatórios

| # | Relatório | Descrição |
|---|-----------|-----------|
| 1 | **Abst_Mot_Por_empresa** | Relatório base de Abastecimento por Motorista |
| 2 | **Ranking_km_Proporcional** | Ranking de Km Proporcional por motorista |
| 3 | **Ranking_Integração** | Ranking com dados de integração |
| 4 | **Ranking_Ouro_Mediano** | Ranking consolidado Ouro Mediano |
| 5 | **Ranking_Por_Empresa** | Ranking por empresa |
| 6 | **RMC_Destribuida** | Resumo Motorista Cliente com Km/l distribuído |
| 7 | **Turnos_Integração** | Análise de turnos de integração |

---

## 🔄 Ordem de Processamento (Dependências)

```
┌──────────────────────────────────────────────────────────────────┐
│                    ORDEM DE PROCESSAMENTO                        │
├──────────────────────────────────────────────────────────────────┤
│                                                                  │
│  1️⃣  Abst_Mot_Por_empresa  ─────────────────────────────────────│
│         │                                                        │
│         ├──► 2️⃣  Ranking_Por_Empresa                            │
│         │         │                                              │
│         │         └──► 4️⃣  Ranking_Ouro_Mediano                 │
│         │                                                        │
│         ├──► 3️⃣  Ranking_Integração                             │
│         │                                                        │
│         ├──► 5️⃣  Ranking_Km_Proporcional                        │
│         │                                                        │
│         └──► 6️⃣  Turnos_Integração                              │
│                                                                  │
│  7️⃣  RMC_Destribuida (independente)                             │
│                                                                  │
└──────────────────────────────────────────────────────────────────┘
```

---

## 📁 Estrutura de Arquivos de Entrada

### Diretório Base (Entrada)

```
📂 [BASE_DIR]/
├── 📂 Integração_Abast/
│   └── Abastecimento_{EMPRESA}_{MÊS}_{ANO}.xlsx
│
├── 📂 Integração_Mot/
│   └── Motorista_{EMPRESA}_{MÊS}_{ANO}.xlsx
│
├── 📂 Ranking/
│   └── Ranking_{EMPRESA}_{MÊS}_{ANO}.xlsx
│
├── 📂 Turnos_128/
│   └── Turnos_128_{EMPRESA}_{MÊS}_{ANO}.xlsx
│
└── 📂 Resumo_Motorista_Cliente/
    └── RMC_{EMPRESA}_{MÊS}_{ANO}.xlsx
```

### Diretório de Saída

```
📂 [OUTPUT_DIR]/
├── 📂 Abst_Mot_Por_empresa/
│   └── 📂 {EMPRESA}/
│       └── 📂 {ANO}/
│           └── 📂 {MÊS}/
│               ├── Abst_Mot_Por_empresa_{EMPRESA}_{MÊS}_{ANO}.xlsx
│               └── Detalhado_{EMPRESA}_{MÊS}_{ANO}.xlsx
│
├── 📂 Ranking_Por_Empresa/
│   └── 📂 {EMPRESA}/
│       └── 📂 {ANO}/
│           └── 📂 {MÊS}/
│               └── Ranking_Por_Empresa_{EMPRESA}_{MÊS}_{ANO}.xlsx
│
├── 📂 Ranking_Integração/
│   └── 📂 {EMPRESA}/
│       └── 📂 {ANO}/
│           └── 📂 {MÊS}/
│               └── Ranking_Integração_{EMPRESA}_{MÊS}_{ANO}.xlsx
│
├── 📂 Ranking_Ouro_Mediano/
│   └── Ranking_Ouro_Mediano_{DATA}.xlsx
│
├── 📂 Rankig_Km_Proporcional/  ⚠️ (Typo no código original)
│   └── 📂 {EMPRESA}/
│       └── 📂 {ANO}/
│           └── 📂 {MÊS}/
│               ├── Consolidado_{EMPRESA}_{MÊS}_{ANO}.xlsx
│               ├── Detalhado_{EMPRESA}_{MÊS}_{ANO}.xlsx
│               └── Ranking_Km_Proporcional_{EMPRESA}_{MÊS}_{ANO}.xlsx
│
├── 📂 Turnos Integração/
│   └── 📂 {EMPRESA}/
│       └── 📂 {ANO}/
│           └── 📂 {MÊS}/
│               └── Turnos_Integração_{EMPRESA}_{MÊS}_{ANO}.xlsx
│
└── 📂 RMC_Destribuida/
    └── 📂 {EMPRESA}/
        └── 📂 {ANO}/
            └── 📂 {MÊS}/
                └── RMC_Km_l_Distribuida_{EMPRESA}_{MÊS}_{ANO}.xlsx
```

---

## 📊 Detalhes de Cada Relatório

### 1️⃣ Abst_Mot_Por_empresa

**Classe:** `CompanyProcessor`

**Arquivos de Entrada Necessários:**
| Pasta | Arquivo | Obrigatório |
|-------|---------|-------------|
| `Integração_Abast` | `Abastecimento_{empresa}_{mês}_{ano}.xlsx` | ✅ Sim |
| `Integração_Mot` | `Motorista_{empresa}_{mês}_{ano}.xlsx` | ✅ Sim |

**Arquivos de Saída:**
- `Abst_Mot_Por_empresa_{empresa}_{mês}_{ano}.xlsx` - Consolidado
- `Detalhado_{empresa}_{mês}_{ano}.xlsx` - Detalhado (usado por outros relatórios)

**Erros Comuns:**
| Erro | Causa | Solução |
|------|-------|---------|
| `Supply folder not found` | Pasta `Integração_Abast` não existe | Criar pasta e adicionar arquivos |
| `Driver folder not found` | Pasta `Integração_Mot` não existe | Criar pasta e adicionar arquivos |
| `No matching supply file` | Arquivo de abastecimento não encontrado | Verificar nomenclatura do arquivo |

---

### 2️⃣ Ranking_Por_Empresa

**Classe:** `RankingProcessor`

**Arquivos de Entrada Necessários:**
| Pasta | Arquivo | Obrigatório |
|-------|---------|-------------|
| `Ranking` | `Ranking_{empresa}_{mês}_{ano}.xlsx` | ✅ Sim |
| `Turnos_128` | `Turnos_128_{empresa}_{mês}_{ano}.xlsx` | ✅ Sim |

**Dependências de Relatórios:**
| Relatório | Obrigatório | Uso |
|-----------|-------------|-----|
| `Abst_Mot_Por_empresa` | ❌ Opcional | Enriquece com dados de abastecimento |
| `Ranking_Km_Proporcional` | ❌ Opcional | Adiciona dados proporcionais |

**Arquivos de Saída:**
- `Ranking_Por_Empresa_{empresa}_{mês}_{ano}.xlsx`

**Erros Comuns:**
| Erro | Causa | Solução |
|------|-------|---------|
| `Ranking folder not found` | Pasta `Ranking` não existe | Criar pasta e adicionar arquivos |
| `Arquivo de ranking não encontrado` | Arquivo de ranking faltando | Adicionar arquivo `Ranking_{empresa}_{mês}_{ano}.xlsx` |
| `Arquivo de turnos não encontrado` | Arquivo de turnos faltando | Adicionar arquivo `Turnos_128_{empresa}_{mês}_{ano}.xlsx` |

---

### 3️⃣ Ranking_Integração

**Classe:** `RankingIntegracaoProcessor`

**Arquivos de Entrada Necessários:**
| Pasta | Arquivo | Obrigatório |
|-------|---------|-------------|
| `Ranking` | `Ranking_{empresa}_{mês}_{ano}.xlsx` | ✅ Sim |
| `Turnos_128` | `Turnos_128_{empresa}_{mês}_{ano}.xlsx` | ✅ Sim |

**Dependências de Relatórios:**
| Relatório | Obrigatório | Uso |
|-----------|-------------|-----|
| `Abst_Mot_Por_empresa` | ✅ Sim | Usa arquivos `Abst_Mot_Por_empresa` e `Detalhado` |

**Arquivos de Saída:**
- `Ranking_Integração_{empresa}_{mês}_{ano}.xlsx`

**Erros Comuns:**
| Erro | Causa | Solução |
|------|-------|---------|
| `Arquivo de ranking não encontrado` | Pasta `Ranking` não tem arquivo | Adicionar arquivo de ranking |
| Dados incompletos | `Abst_Mot_Por_empresa` não foi gerado | Gerar `Abst_Mot_Por_empresa` primeiro |

---

### 4️⃣ Ranking_Ouro_Mediano

**Classe:** `RankingOuroMedianoProcessor`

**Arquivos de Entrada:** Nenhum arquivo externo direto

**Dependências de Relatórios:**
| Relatório | Obrigatório | Uso |
|-----------|-------------|-----|
| `Ranking_Por_Empresa` | ✅ Sim | Lê a aba 'Todos' para filtrar motoristas |

**Filtros Aplicados:**
- `fase` = específica (configurável)
- `status` = específico (configurável)
- `ponto acumulado` >= valor mínimo

**Arquivos de Saída:**
- `Ranking_Ouro_Mediano_{data}.xlsx` ou
- `Ranking_Ouro_Mediano_{empresa}_{período_inicial}_a_{período_final}.xlsx`

**Erros Comuns:**
| Erro | Causa | Solução |
|------|-------|---------|
| `Diretório Ranking_Por_Empresa não encontrado` | Nenhum `Ranking_Por_Empresa` foi gerado | Gerar `Ranking_Por_Empresa` primeiro |
| `Nenhum dado encontrado para consolidação` | Nenhum registro atende aos critérios | Verificar filtros ou dados de entrada |

---

### 5️⃣ Ranking_Km_Proporcional

**Classe:** `RankingKmProporcionalProcessor` (referenciado no código)

**Arquivos de Entrada:** Nenhum arquivo externo direto

**Dependências de Relatórios:**
| Relatório | Obrigatório | Uso |
|-----------|-------------|-----|
| `Abst_Mot_Por_empresa` | ✅ Sim | Usa o arquivo `Detalhado_{empresa}_{período}.xlsx` |

**Arquivos de Saída:**
- `Consolidado_{empresa}_{mês}_{ano}.xlsx`
- `Detalhado_{empresa}_{mês}_{ano}.xlsx`
- `Ranking_Km_Proporcional_{empresa}_{mês}_{ano}.xlsx`

**Erros Comuns:**
| Erro | Causa | Solução |
|------|-------|---------|
| `Arquivo detalhado de origem não encontrado` | `Abst_Mot_Por_empresa` não foi gerado | Gerar `Abst_Mot_Por_empresa` primeiro |

---

### 6️⃣ Turnos_Integração

**Classe:** `TurnosIntegracaoProcessor`

**Arquivos de Entrada:** Nenhum arquivo externo direto

**Dependências de Relatórios:**
| Relatório | Obrigatório | Uso |
|-----------|-------------|-----|
| `Abst_Mot_Por_empresa` | ✅ Sim | Usa o arquivo `Detalhado_{empresa}_{período}.xlsx` |

**Definição de Turnos:**
| Turno | Início | Fim |
|-------|--------|-----|
| Madrugada | 00:00 | 05:59 |
| Manhã | 06:00 | 11:59 |
| Intervalo | 12:00 | 13:59 |
| Tarde | 14:00 | 19:59 |
| Noite | 20:00 | 23:59 |

**Arquivos de Saída:**
- `Turnos_Integração_{empresa}_{mês}_{ano}.xlsx`
  - Aba: `Todos_Turnos`
  - Aba: `Consolidado_Motorista_Turno`
  - Aba: `Consolidado_Turno`

**Erros Comuns:**
| Erro | Causa | Solução |
|------|-------|---------|
| `Diretório Abst_Mot_Por_empresa não encontrado` | Nenhum `Abst_Mot_Por_empresa` foi gerado | Gerar `Abst_Mot_Por_empresa` primeiro |
| `Arquivo Detalhado não encontrado` | Arquivo detalhado específico não existe | Verificar se o período está correto |
| `Colunas necessárias não encontradas` | Estrutura do arquivo diferente | Verificar colunas no arquivo Detalhado |

---

### 7️⃣ RMC_Destribuida (Resumo_Motorista_Cliente)

**Classe:** `RMCProcessor`

**Arquivos de Entrada Necessários:**
| Pasta | Arquivo | Obrigatório |
|-------|---------|-------------|
| `Resumo_Motorista_Cliente` | `RMC_{empresa}_{mês}_{ano}.xlsx` | ✅ Sim |
| `Integração_Abast` | `Abastecimento_{empresa}_{mês}_{ano}.xlsx` | ✅ Sim |

**Arquivos de Saída:**
- `RMC_Km_l_Distribuida_{empresa}_{mês}_{ano}.xlsx`

**Erros Comuns:**
| Erro | Causa | Solução |
|------|-------|---------|
| `Pasta de resumo não encontrada` | Pasta `Resumo_Motorista_Cliente` não existe | Criar pasta e adicionar arquivos |
| `Arquivo de resumo não encontrado` | Arquivo RMC específico não existe | Adicionar arquivo `RMC_{empresa}_{mês}_{ano}.xlsx` |
| `Arquivo de abastecimento não encontrado` | Arquivo de abastecimento faltando | Adicionar arquivo de abastecimento |

---

## 🛠️ Guia de Solução de Problemas

### Passo 1: Verificar Estrutura de Pastas

```bash
# Verificar se as pastas de entrada existem
ls -la D:\Scripts\Entrada\Integração_Abast\
ls -la D:\Scripts\Entrada\Integração_Mot\
ls -la D:\Scripts\Entrada\Ranking\
ls -la D:\Scripts\Entrada\Turnos_128\
ls -la D:\Scripts\Entrada\Resumo_Motorista_Cliente\
```

### Passo 2: Verificar Nomenclatura dos Arquivos

Os arquivos devem seguir exatamente o padrão:
- `Abastecimento_{EMPRESA}_{MÊS}_{ANO}.xlsx` (ex: `Abastecimento_Ideal_Novembro_2025.xlsx`)
- `Motorista_{EMPRESA}_{MÊS}_{ANO}.xlsx`
- `Ranking_{EMPRESA}_{MÊS}_{ANO}.xlsx`
- `Turnos_128_{EMPRESA}_{MÊS}_{ANO}.xlsx`
- `RMC_{EMPRESA}_{MÊS}_{ANO}.xlsx`

### Passo 3: Ordem de Geração

1. **Primeiro:** Gerar `Abst_Mot_Por_empresa` para todas as empresas
2. **Depois:** Gerar `Ranking_Por_Empresa` (se tiver arquivos Ranking e Turnos)
3. **Depois:** Gerar os demais relatórios:
   - `Ranking_Integração`
   - `Ranking_Km_Proporcional`
   - `Turnos_Integração`
   - `Ranking_Ouro_Mediano`
4. **Independente:** `RMC_Destribuida` pode ser gerado a qualquer momento

### Passo 4: Executar Script de Teste

```bash
cd D:\Projetos\Processador_Relatorios
python test_reports.py
```

---

## 📈 Empresas Identificadas nos Logs

Baseado no log `unified_processing.log`, as seguintes empresas foram processadas:

| Empresa | Abst_Mot | Ranking | RMC |
|---------|----------|---------|-----|
| Alpha | ✅ | ❌ (falta Ranking folder) | ❌ |
| Amparo | ✅ | ❌ | ❌ |
| Futuro | ✅ | - | - |
| Gracas | ❌ (falta Detalhado) | - | - |
| Ideal | ✅ | ✅ | ✅ |
| Jabour | ✅ | - | - |
| Novacap | ✅ | - | - |
| Nsgloria | ❌ | - | - |
| Pavunense | ✅ | - | - |
| Pendotiba | ✅ | - | - |
| Pontecoberta | ❌ | - | - |
| Recreio | ✅ | - | - |
| Redentor | ✅ | - | - |
| Reginas | ✅ | - | - |
| Transurb | ❌ | - | - |
| Tursan | ❌ | - | - |
| Verdun | ❌ | - | - |
| Vilareal | ❌ | - | - |

**Legenda:**
- ✅ = Processado com sucesso
- ❌ = Erro/Arquivo faltando
- `-` = Não testado/Não aplicável

---

## 🔧 Correções Sugeridas no Código

### 1. Typo na pasta "Rankig_Km_Proporcional"

No código atual, a pasta está escrita como `Rankig_Km_Proporcional` (faltando um 'n'):

```python
# Linha 806-808 em main.py
consolidado_km_prop_file = os.path.join(
    self.OUTPUT_BASE_DIR, 
    'Rankig_Km_Proporcional',  # ← Typo aqui
    ...
)
```

**Correção sugerida:** Manter consistência ou corrigir para `Ranking_Km_Proporcional`

### 2. Pasta "Turnos Integração" com espaço

A pasta de saída usa espaço: `Turnos Integração`, o que pode causar problemas em alguns sistemas.

**Correção sugerida:** Usar `Turnos_Integracao` sem espaços e acentos.

---

## 📝 Checklist de Validação

```
[ ] Diretórios de entrada existem
[ ] Diretórios de saída existem
[ ] Arquivos de Abastecimento presentes
[ ] Arquivos de Motorista presentes
[ ] Arquivos de Ranking presentes (se aplicável)
[ ] Arquivos de Turnos_128 presentes (se aplicável)
[ ] Arquivos RMC presentes (se aplicável)
[ ] Abst_Mot_Por_empresa gerado primeiro
[ ] Relatórios dependentes gerados na ordem correta
```

---

*Documento gerado em: {data_atual}*
*Versão do Processador: Verificar main.py*


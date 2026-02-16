# DD Checklist Generator

**Automated Due Diligence checklist generator for M&A transactions.**

Generates professionally formatted Excel workbooks tailored to the deal type, sector, and jurisdiction — ready for immediate use by legal, financial, and advisory teams.

[Leia em Português](#-versão-portuguesa)

---

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Usage](#usage)
- [Output](#output)
- [API Reference](#api-reference)
- [Project Structure](#project-structure)
- [Roadmap](#roadmap)
- [License](#license)
- [Versão Portuguesa](#-versão-portuguesa)

---

## Features

### Supported Transaction Types

| Type | Description |
|------|-------------|
| **Share Deal** | Acquisition of company shares/equity |
| **Asset Deal** | Acquisition of specific business assets |
| **Merger** | Full corporate merger/combination |

### Sectors

Healthcare, Technology, Industrial, Real Estate, Financial Services, Retail — each with sector-specific document requirements.

### Jurisdictions

Portugal, Spain (Espanha), International — with jurisdiction-aware document lists.

### Bilingual Support

Full support for **English (EN)** and **Portuguese (PT)** — all document names, headers, labels, statuses, and instructions are translated.

### Excel Output (3 Tabs)

1. **Checklist** — Complete document request list with:
   - 8 categories: Legal, Financial, Operational, Tax, HR, Commercial, IP, Compliance
   - Priority levels (High / Medium / Low) with conditional formatting
   - Tracking columns: Received Date, Status, Responsible, Comments
   - Dropdown data validation for Status and Priority
   - Auto-filter and freeze panes

2. **Summary** — Transaction metadata and statistics:
   - Deal parameters (target, type, sector, jurisdiction)
   - Document count by category and priority

3. **Instructions** — Usage guide:
   - Step-by-step workflow
   - Status definitions with color coding
   - Indicative DD timeline (8–10 weeks)
   - Advisor contacts template

---

## Installation

**Requirements:** Python 3.8+

```bash
# Clone the repository
git clone https://github.com/limbus6/dd-checklist-generator.git
cd dd-checklist-generator

# Install dependencies
pip install openpyxl
```

---

## Quick Start

```bash
# Interactive mode — guided prompts
python dd_checklist.py

# Test mode — generates sample files automatically
python dd_checklist.py --test
```

---

## Usage

### Interactive Mode

```bash
python dd_checklist.py
```

The CLI guides you through:
1. Language (EN / PT)
2. Transaction type
3. Sector
4. Jurisdiction
5. Target company name
6. Preview of the document list
7. Optional: add custom documents
8. Confirm and generate

### Programmatic Mode

```python
from dd_checklist import run_automated

path = run_automated(
    company_name="Acme Corp",
    deal_type="Share Deal",
    sector="Technology",
    jurisdiction="Portugal",
    language="EN",
)
print(f"Generated: {path}")
```

#### Parameters

| Parameter | Type | Required | Values |
|-----------|------|----------|--------|
| `target` | `str` | Yes | Company name |
| `deal_type` | `str` | Yes | `"Share Deal"`, `"Asset Deal"`, `"Merger"` |
| `sector` | `str` | Yes | `"Healthcare"`, `"Technology"`, `"Industrial"`, `"Real Estate"`, `"Financial Services"`, `"Retail"` |
| `jurisdiction` | `str` | Yes | `"Portugal"`, `"Espanha"`, `"Internacional"` |
| `lang` | `str` | No | `"EN"` (default), `"PT"` |
| `custom_docs` | `list[tuple]` | No | List of `(category, name, required, priority)` tuples |

### Test Mode

```bash
python dd_checklist.py --test
```

Generates two sample checklists for validation:
- `TechVida_Lda_DD_Checklist_YYYYMMDD.xlsx` — Technology / Share Deal / EN
- `Farma_Saude_SA_DD_Checklist_YYYYMMDD.xlsx` — Healthcare / Merger / PT

---

## Output

**Filename format:** `{CompanyName}_DD_Checklist_{YYYYMMDD}.xlsx`

Each checklist contains **40–50 documents** adapted to the specific deal context, with professional formatting and color-coded priorities:

| Priority | Color |
|----------|-------|
| High | Red |
| Medium | Yellow |
| Low | Green |

| Status | Color |
|--------|-------|
| Pending | Orange |
| Received | Blue |
| Reviewed | Green |
| Missing | Red |

---

## Project Structure

```
dd-checklist-generator/
├── dd_checklist.py    # Main application (single-file)
├── README.md
└── .gitignore
```

---

## Roadmap

- [ ] AI-powered analysis via Claude API
- [ ] Streamlit web interface
- [ ] Expanded language support (ES, FR, DE)
- [ ] PDF export
- [ ] Data room integration (SharePoint, Box)
- [ ] Customizable user templates
- [ ] REST API

---

## Tech Stack

- **Python 3.8+**
- **openpyxl** — Excel file generation and formatting

---

## License

MIT License

---

## Author

**limbus6** — [@limbus6](https://github.com/limbus6)

---

---

# Versão Portuguesa

**Gerador automatizado de checklists de Due Diligence para transações de M&A.**

Gera ficheiros Excel profissionalmente formatados, adaptados ao tipo de deal, sector e jurisdição — prontos para uso imediato por equipas jurídicas, financeiras e de assessoria.

---

## Funcionalidades

### Tipos de Transação Suportados

| Tipo | Descrição |
|------|-----------|
| **Share Deal** | Aquisição de participações sociais / ações |
| **Asset Deal** | Aquisição de ativos específicos do negócio |
| **Merger** | Fusão de sociedades |

### Sectores

Healthcare, Technology, Industrial, Real Estate, Financial Services, Retail — cada um com documentos específicos do sector.

### Jurisdições

Portugal, Espanha, Internacional — com listas de documentos adaptadas à jurisdição.

### Suporte Bilingue

Suporte completo para **Inglês (EN)** e **Português (PT)** — nomes de documentos, cabeçalhos, labels, estados e instruções são traduzidos.

### Output Excel (3 Separadores)

1. **Checklist** — Lista completa de documentos com:
   - 8 categorias: Legal, Financial, Operational, Tax, HR, Commercial, IP, Compliance
   - Prioridades (High / Medium / Low) com formatação condicional
   - Colunas de tracking: Data de Receção, Estado, Responsável, Comentários
   - Validação de dados via dropdowns para Estado e Prioridade
   - Auto-filter e freeze panes

2. **Resumo** — Metadata da transação e estatísticas:
   - Parâmetros do deal (target, tipo, sector, jurisdição)
   - Contagem de documentos por categoria e prioridade

3. **Instruções** — Guia de utilização:
   - Workflow passo a passo
   - Definições de estado com código de cores
   - Timeline indicativo de DD (8–10 semanas)
   - Template de contactos de assessores

---

## Instalação

**Requisitos:** Python 3.8+

```bash
# Clonar o repositório
git clone https://github.com/limbus6/dd-checklist-generator.git
cd dd-checklist-generator

# Instalar dependências
pip install openpyxl
```

---

## Início Rápido

```bash
# Modo interativo — prompts guiados
python dd_checklist.py

# Modo teste — gera ficheiros de exemplo automaticamente
python dd_checklist.py --test
```

---

## Utilização

### Modo Interativo

```bash
python dd_checklist.py
```

O CLI guia-o através de:
1. Idioma (EN / PT)
2. Tipo de transação
3. Sector
4. Jurisdição
5. Nome da empresa target
6. Preview da lista de documentos
7. Opcional: adicionar documentos personalizados
8. Confirmar e gerar

### Modo Programático

```python
from dd_checklist import run_automated

path = run_automated(
    target="Acme Corp",
    deal_type="Share Deal",
    sector="Technology",
    jurisdiction="Portugal",
    lang="EN",
)
print(f"Gerado: {path}")
```

#### Parâmetros

| Parâmetro | Tipo | Obrigatório | Valores |
|-----------|------|-------------|---------|
| `target` | `str` | Sim | Nome da empresa |
| `deal_type` | `str` | Sim | `"Share Deal"`, `"Asset Deal"`, `"Merger"` |
| `sector` | `str` | Sim | `"Healthcare"`, `"Technology"`, `"Industrial"`, `"Real Estate"`, `"Financial Services"`, `"Retail"` |
| `jurisdiction` | `str` | Sim | `"Portugal"`, `"Espanha"`, `"Internacional"` |
| `lang` | `str` | Não | `"EN"` (default), `"PT"` |
| `custom_docs` | `list[tuple]` | Não | Lista de tuplos `(categoria, nome, obrigatório, prioridade)` |

### Modo Teste

```bash
python dd_checklist.py --test
```

Gera duas checklists de exemplo para validação:
- `TechVida_Lda_DD_Checklist_YYYYMMDD.xlsx` — Technology / Share Deal / EN
- `Farma_Saude_SA_DD_Checklist_YYYYMMDD.xlsx` — Healthcare / Merger / PT

---

## Output

**Formato do ficheiro:** `{NomeEmpresa}_DD_Checklist_{YYYYMMDD}.xlsx`

Cada checklist contém **40–50 documentos** adaptados ao contexto específico do deal, com formatação profissional e prioridades codificadas por cores:

| Prioridade | Cor |
|------------|-----|
| High | Vermelho |
| Medium | Amarelo |
| Low | Verde |

| Estado | Cor |
|--------|-----|
| Pendente | Laranja |
| Recebido | Azul |
| Revisto | Verde |
| Em falta | Vermelho |

---

## Roadmap

- [ ] Análise com IA via Claude API
- [ ] Interface web com Streamlit
- [ ] Suporte multi-idioma expandido (ES, FR, DE)
- [ ] Exportação para PDF
- [ ] Integração com data rooms (SharePoint, Box)
- [ ] Templates customizáveis por utilizador
- [ ] API REST

---

## Stack Tecnológica

- **Python 3.8+**
- **openpyxl** — Geração e formatação de ficheiros Excel

---

## Licença

MIT License

---

## Autor

**limbus6** — [@limbus6](https://github.com/limbus6)

# Engineering Hours Consolidation System

![Python](https://img.shields.io/badge/Python-3.8+-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white)
![Excel](https://img.shields.io/badge/Excel%20%2F%20xlwings-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![ETL](https://img.shields.io/badge/ETL-Pipeline-FF6B35?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Production-brightgreen?style=for-the-badge)

> **Automated ETL pipeline** that consolidates weekly work-hour logs from an engineering team (~15 members), validates compliance per member, and generates individual status reports — replacing a fully manual Excel process.

---

## Business Problem

A logistics company's engineering team tracked weekly hours manually across individual `.xlsm` workbooks. The consolidation process was:

- ❌ Done **entirely by hand** every month
- ❌ Error-prone when merging across multiple team members  
- ❌ Had **no automatic alerts** for members falling behind
- ❌ Didn't account for **holidays** or **variable weekly schedules** (e.g. part-time members at 20h vs 45h/week)

---

## Solution

A Python pipeline that runs in one command and:

1. **Reads** every team member's `.xlsm` workbook from a shared folder
2. **Converts** percentage-based entries into actual worked hours (based on each member's individual weekly schedule)
3. **Reshapes** the data from wide format (one column per week) to long format (one row per task/week)
4. **Consolidates** everything into a single timestamped `.xlsx` dataset
5. **Evaluates** each member's monthly compliance vs. expected hours, adjusted for holidays
6. **Generates** individual Excel status reports per member with a 5-tier alert system

---

## Data Schema

### Input: Individual member files (`Integrantes/Nombre_Apellido.xlsm`)

Each team member has their own workbook. The main sheet (named after the member) has a **wide format** with one column per week of the year:

| Rows 1–5 | Header metadata |
|---|---|
| Row 1 | Week start dates (`datetime`) |
| Row 2 | Week number (`1` → `53`) |
| Row 3 | Holiday count per week |
| Row 4 | `% Semana` — adjusted week percentage if there are holidays |
| Row 5 | `Total` hours auto-sum |

| Row 7 | Column headers |
|---|---|
| Cols A–H | Task metadata: `ID`, `Proyecto`, `Usuario`, `Área`, `UUNN`, `CeCo`, `Tipo`, `Facturación` |
| Cols I–BI | Weeks 1–53: member enters a **percentage** (e.g. `0.5` = 50% of weekly hours on this project) |

**Example data row:**

| ID | Proyecto | Usuario | Área | UUNN | CeCo | Tipo | Facturación | Sem 1 | Sem 2 | ... |
|---|---|---|---|---|---|---|---|---|---|---|
| 82 | YPF Aduana | YPF | Planeamiento | UNL | Desarrollo | Implementaciones | — | — | 0.2 | ... |

> The pipeline multiplies each percentage by the member's weekly hours (from `Horas_ING.xlsm`) to get **actual hours worked per task per week**.

---

### Input: Master workbook (`Horas_ING.xlsm`)

Contains three sheets:

**`Integrantes`** — Team member registry:

| Integrantes | Área | Horas semanales [h] | Legajo | Mail |
|---|---|---|---|---|
| García Martín | Excelencia Operativa | 45 | 1001 | mgarcia@empresa.com |
| Ramírez Diego | Planeamiento Estratégico | 20 | 1003 | dramirez@empresa.com |

**`Proyectos`** — Project master list:

| ID | Proyecto | Usuario | Área | UUNN | CeCo | Tipo | Facturación |
|---|---|---|---|---|---|---|---|
| 1 | Optimización rutas | Cliente A | Excelencia Operativa | UNL | Desarrollo | Implementaciones | — |

**`Consolidado Horas`** — Populated automatically by the pipeline.

---

### Output: Consolidated dataset

Long-format table with one row per (member × task × week):

| Column | Description |
|---|---|
| `ID` | Project ID |
| `Proyecto` | Project name |
| `Área` | Engineering area |
| `Semana` | Week number |
| `Horas` | Actual hours (percentage × weekly schedule) |
| `Mes` | Month number |
| `Mes nombre` | Month name (Spanish) |
| `Inicio Semana` | Week start date |
| `Integrante` | Team member name |
| `Legajo` | Employee ID |

---

## Pipeline Flow

```
┌──────────────────────────────┐
│  Integrantes/*.xlsm          │  ← One file per team member (wide format)
│  + Horas_ING.xlsm            │  ← Member registry + project master
└──────────────┬───────────────┘
               │  get_integrantes_dfs()
               ▼
┌──────────────────────────────┐
│  % × horas_semanales         │  ← Convert percentages → actual hours
│  per member per week         │    (e.g. 0.5 × 45h = 22.5h on that project)
└──────────────┬───────────────┘
               │  filter_integrante_df()
               ▼
┌──────────────────────────────┐
│  Wide → Long reshape         │  ← One row per (member, project, week)
│  Drop empty weeks            │
└──────────────┬───────────────┘
               │  general_df()
               ▼
┌──────────────────────────────┐
│  Enrich with month, dates    │  ← Add Mes, Mes nombre, Inicio Semana
│  Save timestamped .xlsx      │  ← Consolidado Horas (dd-mm-yyyy).xlsx
└──────────────┬───────────────┘
               │  get_integrantes_month_status()
               ▼
┌──────────────────────────────┐
│  Holiday-adjusted compliance │  ← Expected hours per week adjusted for
│  check per member/month      │    feriados (e.g. 1 holiday → 80% week)
└──────────────┬───────────────┘
               │  save_status()
               ▼
┌──────────────────────────────┐
│  Per-member status .xlsx     │  ← Individual report for each member
│  with alert level            │    written via xlwings into a template
└──────────────────────────────┘
```

---

## 🔔 Alert Logic

The system evaluates each member's monthly completion percentage and, based on the current date relative to the month-end deadline, assigns one of five alert levels:

| Alert | Condition | Meaning |
|---|---|---|
| `NO AVISAR` | On track | No action needed |
| `AVISAR LEVE` | After 25th, completion 50–60% | Missing ~1 week, soft reminder |
| `AVISAR` | After 25th, completion < 50% | Missing multiple weeks |
| `AVISAR FUERTE` | After 25th, zero hours logged | Member hasn't logged anything |
| `AVISAR MUY FUERTE` | ≤2 days to deadline, still incomplete | Urgent escalation |

> **Deadline** = last day of month + 8 days, giving members a buffer into the following month.  
> **Holiday adjustment**: a week with 1 holiday out of 5 days → expected hours × 80%.

---

## Libraries used

| Library | Use |
|---|---|
| `pandas` | Data ingestion, reshaping (wide→long), groupby aggregations |
| `numpy` | Vectorized operations, NaN handling |
| `xlwings` | Read/write `.xlsm` files with macro support |
| `openpyxl` | Write output `.xlsx` status reports |
| `datetime` / `calendar` | Deadline calculation, week/month mapping |
| `locale` | Spanish date formatting |
| `shutil` / `os` | File system operations, report cleanup |

---

## Getting Started

### 1. Install dependencies

```bash
pip install pandas numpy xlwings openpyxl
```

### 2. Try it with sample data

```bash
git clone https://github.com/your-username/hours-consolidation
cd hours-consolidation
```

Update the paths in `consolidator.py`:

```python
carpeta_ing_path      = "sample_data/input/Integrantes"
integrantes_info_path = "sample_data/input/Horas_ING_sample.xlsx"
dir_consolidado       = "sample_data/output"
```

Then run:

```bash
python consolidator.py
```

### 3. Run for a specific month

```python
# Bottom of consolidator.py — change to target a past month:
status = get_integrantes_month_status(
    df_general, path_p, integrantes_info_path,
    fecha=(6, 2024)   # (month, year)
)
```

---

## Repository Structure

```
hours-consolidation/
│
├── consolidator.py                  # Main ETL pipeline
│
├── sample_data/
│   ├── input/
│   │   ├── Horas_ING_sample.xlsx    # Master workbook (synthetic data)
│   │   └── Integrantes/             # Individual member files (synthetic)
│   │       ├── García_Martín.xlsx
│   │       ├── López_Sofía.xlsx
│   │       └── ...
│   └── output/                      # Generated on run
│       ├── Consolidado Horas (dd-mm-yyyy).xlsx
│       └── Status Integrantes/
│           ├── García Martín.xlsx
│           └── ...
│
└── README.md
```

---

## 📈 Impact

- Eliminated **manual monthly consolidation** across 15+ Excel files
- Enabled **accurate compliance tracking** with holiday-adjusted expected hours
- Replaced manual follow-up with an **automated 5-tier alert system**
- Produced **individual reports** deliverable directly to each team member
- Deployed and used in **production** within a logistics company engineering team

---

## 👤 Author

**Fernando Bonvecchiato**  
Industrial Engineering Student | Data Analyst  

# VBA ETL Accelerator

**Short Description:**
A lightweight, VBA-driven ETL utility that automates extraction from Excel workbooks, data validation and transformation, and loading into Azure SQL, all within a macro-enabled Excel environment.

---

## Table of Contents

1. [Objective](#objective)
2. [Repository Structure](#repository-structure)
3. [Prerequisites](#prerequisites)
4. [Setup & Installation](#setup--installation)
5. [Usage](#usage)
6. [CI & Testing](#ci--testing)
7. [Local Integration Testing](#local-integration-testing)
8. [Contribution Guidelines](#contribution-guidelines)
9. [License](#license)

---

## Objective

Establish the foundational repository and CI pipeline for the VBA ETL Accelerator. This project provides a modular VBA framework to:

* **Extract** raw data from Excel and CSV files.
* **Transform** data through validation, defaults, and column mapping.
* **Load** cleaned data into Azure SQL with upsert semantics.
* **Orchestrate** downstream Azure Data Factory pipelines.

---

## Repository Structure

```
vba_etl_accelerator/
├── workbook/
│   └── ETL_Accelerator.xlsm       # Macro-enabled Excel workbook
├── modules/                       # VBA code modules
│   ├── Extract.bas
│   ├── Transform.bas
│   ├── Load.bas
│   ├── Orchestrate.bas
│   └── Config.bas                 # Configuration loader/validator
├── tests/                         # Rubberduck VBA test classes
│   ├── TestExtract.cls
│   ├── TestTransform.cls
│   ├── TestLoad.cls
│   ├── TestOrchestrate.cls
│   └── TestConfig.cls
├── sample_data/                   # Sample datasets for demos
│   └── sample_input.xlsx
├── docs/
│   └── architecture.png           # ETL flow diagram
├── .github/
│   └── workflows/
│       └── ci.yml                 # GitHub Actions CI workflow
├── docker-compose.yml             # Local SQL Server setup
├── scripts/
│   └── start-local-sql.ps1        # ODBC DSN & container start script
├── README.md                      # This file
├── LICENSE                        # MIT License
└── CONTRIBUTING.md                # Contribution guidelines
```

---

## Prerequisites

* Windows OS with Microsoft Excel (2016 or later)
* Rubberduck VBA add-in (install via Chocolatey or MSI)
* Azure SQL Database credentials (ODBC DSN or connection string)
* Git and GitHub CLI (optional for repo management)

---

## Setup & Installation

1. **Clone the repository**

   ```bash
   git clone https://github.com/your-org/vba_etl_accelerator.git
   cd vba_etl_accelerator
   ```
2. **Open the workbook**

   * Launch Excel and open `workbook/ETL_Accelerator.xlsm`.
   * Enable macros when prompted.
3. **Install Rubberduck** (if not already installed):

   ```powershell
   choco install rubberduck
   ```

---

## Usage

1. **Configure settings**

   * Edit the hidden `Config` worksheet in `ETL_Accelerator.xlsm` to set source folder, file patterns, column mappings, and connection strings.
2. **Run ETL pipeline**

   * In Excel, press `Alt+F8`, select `SampleRun`, and click **Run**.
   * Monitor the **Logs** worksheet for progress and errors.
3. **Verify results**

   * Check the target Azure SQL table for inserted/updated records.
   * Confirm that the ADF pipeline was triggered (via Azure portal or logs).

---

## CI & Testing

Automated tests run on each push via GitHub Actions:

### CI Workflow: `.github/workflows/ci.yml`

```yaml
name: CI
on:
  push:
    branches: [ main ]
jobs:
  test:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v3
      - name: Install Rubberduck
        run: choco install rubberduck
      - name: Run VBA Tests
        run: |
          cd workbook
          "C:\Program Files\Rubberduck\Rubberduck.CLI.exe" -project ETL_Accelerator.xlsm -tests all
```

### Local Test Command

Run all tests locally via Rubberduck CLI:

```powershell
& "C:\Program Files\Rubberduck\Rubberduck.CLI.exe" -project ".\workbook\ETL_Accelerator.xlsm" -tests all
```

---

## Local Integration Testing

Spin up a local SQL Server container and register an ODBC DSN:

```powershell
.
scripts/start-local-sql.ps1
```

* The script brings up the `mssql/server` container and configures an ODBC DSN named `VBA_ETL_Test`.
* Update the `Config` worksheet to point to `VBA_ETL_Test` for local runs.

Tear down the environment:

```powershell
docker-compose down
```

---

## Contribution Guidelines

See [CONTRIBUTING.md](CONTRIBUTING.md) for details on branch naming, commit messages, and pull request process.

---

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.

# Quantify Excel Importer

Excel add-in (VSTO) that imports bill-of-materials and material-list data from Excel into **Avontus Quantify**.

Built for scaffolding and rental teams who already maintain material lists in Excel and need a reliable path into Quantify estimates and inventory.

## Features

- Recognizes common material-list layouts (Cuplok, Haki, Kwikstage, Layher, SMART Scaffolder, tube & fitting, and related formats)
- Converts spreadsheet rows into Quantify import data (`.qimport`)
- Opens results directly in Quantify when the desktop client is installed
- Surfaces validation errors before import so bad rows can be corrected in Excel
- Sample workbooks included under `SampleFiles/`

## Prerequisites

| Requirement | Notes |
|---|---|
| Windows | Desktop development |
| Visual Studio 2017+ | With **Office/SharePoint development** workload (VSTO) |
| .NET Framework 4.6 | Target framework for both projects |
| Microsoft Excel | Host application for the add-in |
| VSTO 4.0 Runtime | Required to run the add-in |
| Avontus Quantify | Desktop client; library DLLs must match your Quantify version |

## Getting started

1. Clone this repository.
2. Create a `binaries/` folder at the solution root and copy these Quantify assemblies from a Quantify install or build output:
   - `Avontus.Core.dll`
   - `Avontus.Rental.Library.dll`
   - `Avontus.Rental.Utility.dll`
   - `SmartAssembly.ReportUsage.dll` (optional; used by utilities / exception reporting)
3. Open `Avontus.Quantify.ExcelImporter.sln` in Visual Studio.
4. Build **Debug** and run (F5). Excel starts with the Quantify Excel Importer ribbon loaded.
5. Open a workbook from `SampleFiles/` (or your own material list) and use the ribbon actions to validate, export, or send data to Quantify.

### Release signing (optional)

Code-signing certificates are **not** included in this repository. For release builds:

1. Place your `.pfx` locally (do not commit it).
2. Set `SignManifests` to `true` and `ManifestKeyFile` in `Avontus.Quantify.ExcelImporter.csproj`.
3. Update Advanced Installer / SmartAssembly / FinalBuilder projects under `installer/` and `Build ExcelImporter.fbp8` with your certificate path and password (use secure storage; never commit passwords).

## Project structure

```
Avontus.Quantify.ExcelImporter.sln
├── Avontus.Quantify.ExcelImporter/     # VSTO Excel add-in (UI, import logic)
├── Avontus.Quantify.ExcelImporter.Utils/
├── SampleFiles/                        # Example material-list workbooks
├── installer/                          # Advanced Installer + SmartAssembly projects
├── binaries/                           # Local Quantify DLLs (not in source control)
├── SharedAssemblyInfo.cs               # Shared version (currently 2.0.79.0)
└── Build ExcelImporter.fbp8            # FinalBuilder release pipeline (credentials redacted)
```

## Configuration

User settings live in `Avontus.Quantify.ExcelImporter/app.config` (for example `SkipBlankPNWithQty`). Username is stored empty by default; credentials are entered at login when connecting to Quantify.

## Support

For product questions and support, contact [Avontus](https://www.avontus.com).

Copyright © Avontus Software Corporation. All rights reserved.

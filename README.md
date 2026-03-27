# ParametricBuilder

A WPF desktop application that automates parametric CAD model generation in **SolidWorks** using Excel-driven configuration. Users define parameters through a structured Excel template, select or create model configurations, and trigger automated SolidWorks model updates — all from a clean Material Design UI.

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Architecture](#architecture)
- [Project Structure](#project-structure)
- [How It Works](#how-it-works)
- [Configuration Data Layout](#configuration-data-layout)
- [Excel Template Format](#excel-template-format)
- [Parameter Control Types](#parameter-control-types)
- [Parameter Validation Rules](#parameter-validation-rules)
- [Parameter Dependencies](#parameter-dependencies)
- [Model Data Management](#model-data-management)
- [App Configuration](#app-configuration)
- [Tech Stack & NuGet Packages](#tech-stack--nuget-packages)
- [Getting Started](#getting-started)
- [Building the Project](#building-the-project)
- [Running Unit Tests](#running-unit-tests)
- [Logging](#logging)
- [Release History](#release-history)

---

## Overview

ParametricBuilder bridges the gap between **Excel-based engineering specifications** and **SolidWorks parametric models**. Instead of manually editing SolidWorks configurations, engineers fill in parameter values through a structured WPF UI, select the desired product configurator, and let the application handle the rest — updating the Excel template and launching the underlying `SolidworksConfigurator.exe` (CadCaster) to drive the CAD automation.

---

## Features

- **Multiple Configurators** — Switch between product types (e.g., different heat exchanger variants) from a dropdown; each configurator has its own Excel template, model data, CAD files, and geometry image.
- **Dynamic Parameter UI** — Parameters are loaded directly from the Excel configuration template at runtime. Control types (TextBox, ComboBox, CheckBoxList) are driven by the Excel definition.
- **Built-in Validation** — Each parameter supports numeric range validation (`min–max`) or allowed-value lists, with real-time error feedback in the UI.
- **Parameter Dependencies** — A parameter can be conditionally visible/enabled based on the value selected in another parameter (controller → dependent relationship).
- **Model Library** — Pre-configured model data stored in a Model Data Excel file. Users can select an existing model to auto-populate all parameter values, or start fresh with a new model.
- **Add / Update / Delete Models** — Persist new or modified model configurations back to the Model Data Excel file directly from the UI.
- **Drag-and-Drop Excel** — Drop an Excel file onto the main window to load it as the active file.
- **Output Path Selection** — Browse and set the SolidWorks output directory before running.
- **CadCaster Integration** — Writes parameter values into the Config Excel Template, updates the CadCaster `.config` file, and launches `SolidworksConfigurator.exe` as administrator.
- **In-App Log Viewer** — A real-time log panel shows all INFO/DEBUG/WARN/ERROR messages, capped at 1000 entries to prevent memory growth.
- **Orphaned Excel Process Cleanup** — Automatically kills background Excel processes left open by automation before and after each run.
- **NLog File Logging** — Daily rolling log files stored in a `logs/` folder next to the executable.

---

## Architecture

The application follows the **MVVM (Model-View-ViewModel)** pattern with WPF data bindings.

```
┌─────────────────────────────────────────────────────┐
│                       View                          │
│              MainWindow.xaml (WPF)                  │
│  Data bindings ↕  Commands ↕  Converters ↕          │
├─────────────────────────────────────────────────────┤
│                    ViewModel                        │
│           MainWindowViewModel.cs                    │
│   - ObservableCollections (Parameters, Models, …)  │
│   - ICommand bindings                               │
│   - Dependency setup & validation orchestration    │
├─────────────────────────────────────────────────────┤
│                    Commands                         │
│  BrowseCommand │ GetExcelParametersCommand           │
│  RunCadCasterCommand │ AddCommand                   │
│  UpdateCommand │ DeleteCommand                      │
│  ChangeOutputPathCommand                            │
├─────────────────────────────────────────────────────┤
│                     Models                          │
│  ParameterModel  │  Configurator  │  ModelData      │
│  ExcelFileEntry  │  DependencyManager               │
├─────────────────────────────────────────────────────┤
│                    Helpers / Services               │
│  ExcelHelper  │  ExcelModelExtractor                │
│  exeHelper    │  UpdateExcelPathInConfig            │
│  ConfiguratorDataService                            │
│  CadAiLogger (NLog + Custom UI Event Target)        │
└─────────────────────────────────────────────────────┘
```

---

## Project Structure

```
ParametricBuilder/
├── ParametricBuilder.sln
│
├── ParametricBuilder/                        # Main WPF application
│   ├── App.xaml / App.xaml.cs
│   ├── App.config                            # Runtime settings (paths, config keys)
│   ├── CadAiLogger.cs                        # NLog setup + UI event target
│   │
│   ├── Views/
│   │   └── MainWindow.xaml / .xaml.cs        # Single-window WPF UI
│   │
│   ├── ViewModels/
│   │   ├── MainWindowViewModel.cs            # Primary ViewModel
│   │   └── ConfiguratorDataService.cs        # Loads configurator list from Mapping Excel
│   │
│   ├── Commands/
│   │   ├── BaseCommand.cs                    # ICommand base implementation
│   │   ├── BrowseCommand.cs                  # File-open dialog for Excel selection
│   │   ├── GetExcelParametersCommand.cs      # Reads parameters + model data from Excel
│   │   ├── RunCadCasterCommand.cs            # Writes Excel, updates config, runs CadCaster
│   │   ├── AddCommand.cs                     # Saves new model to Model Data Excel
│   │   ├── UpdateCommand.cs                  # Updates existing model in Model Data Excel
│   │   ├── DeleteCommand.cs                  # Removes model from Model Data Excel
│   │   └── ChangeOutputPathCommand.cs        # Browse for SolidWorks output folder
│   │
│   ├── Models/
│   │   ├── ParameterModel.cs                 # Bindable parameter with validation & dependency
│   │   ├── Configurator.cs                   # Product configurator definition (paths)
│   │   ├── ModelData.cs                      # Single name-value pair for model parameters
│   │   ├── ExcelFileEntry.cs                 # Display name + full path for Excel file entries
│   │   ├── DependencyManager.cs              # Static registry for controller→dependent relationships
│   │   └── IFileHandler.cs                   # Interface for drag-and-drop file handling
│   │
│   ├── Helpers/
│   │   ├── ExcelHelper.cs                    # Kills orphaned background Excel processes
│   │   ├── ExcelModelExtractor.cs            # Reads model names from Model Data Excel
│   │   ├── exeHelper.cs                      # Runs an .exe as administrator
│   │   ├── UpdateExcelPathInConfig.cs        # Patches CadCaster .exe.config XML
│   │   ├── AutoScrollBehavior.cs             # Attached behavior for auto-scrolling ListBox
│   │   ├── BoolToVisibilityConverter.cs      # IValueConverter: bool → Visibility
│   │   ├── CommaSeparatedValuesConverter.cs  # IValueConverter: List<string> ↔ comma string
│   │   ├── ParameterTemplateSelector.cs      # DataTemplateSelector by ControlType
│   │   └── ParameterToTooltipConverter.cs    # IValueConverter: ParameterModel → tooltip text
│   │
│   ├── Behavior/
│   │   └── ListBoxSelectedItemsBehavior.cs   # Syncs ListBox multi-selection to ViewModel
│   │
│   ├── Properties/
│   │   └── AssemblyInfo.cs
│   │
│   └── Resource/
│       ├── excel_solidworks_icon.ico / .png
│       └── ParametricBuilderV1.1.zip
│
├── UnitTestParametricBuilder/                # MSTest unit test project
│   └── UnitTest1.cs
│
└── packages/                                 # NuGet packages (packages.config style)
```

---

## How It Works

### End-to-End Workflow

```
1. Launch Application
        │
        ▼
2. Select Configurator (dropdown)
   → App reads MappingExcel to populate the list
        │
        ▼
3. Click "Get Parameters"
   → Reads Config Excel Template (Sheet 1)
   → Filters rows where Column C = "INPUT"
   → Builds ParameterModel list with control types, validation rules, and dependencies
   → Reads Model Data Excel to populate the Models dropdown
        │
        ▼
4. Select Existing Model  ─OR─  Enter New Values
   → Auto-fills parameter values from Model Data Excel
        │
        ▼
5. Review / Edit Parameter Values
   → Real-time validation (range checks, allowed values)
   → Dependent parameters show/hide automatically
        │
        ▼
6. Set Output Path (Browse or use default from App.config)
        │
        ▼
7. Click "Run"
   → Closes orphaned Excel processes
   → Opens Config Excel Template, writes parameter values (Column B, starting Row 2)
   → Writes Output Path to Sheet 2, Cell B2
   → Saves workbook
   → Updates SolidworksConfigurator.exe.config with the full Excel path
   → Launches SolidworksConfigurator.exe as Administrator
   → On completion, clears Column B in the Config Excel Template
   → Closes orphaned Excel processes again
```

---

## Configuration Data Layout

The `Config.Data/` folder (relative to the executable) holds all data files:

```
Config.Data/
├── MappingExcel_ParametricBuilder_V1.xlsx    ← Master mapping file (App.config: MappingExcel)
│
├── ConfigExcelTemplates/                     ← One Excel per configurator type
│   ├── ProductA_Config.xlsx
│   └── ProductB_Config.xlsx
│
├── ModelDataExcel/                           ← Saved model parameter sets
│   ├── ProductA_Models.xlsx
│   └── ProductB_Models.xlsx
│
├── MasterCadFiles/                           ← SolidWorks template files (.SLDPRT / .SLDASM)
│   └── ...
│
└── GeometryImages/                           ← Preview images shown in the UI
    └── ...
```

### Mapping Excel Format (`MappingExcel_ParametricBuilder_V1.xlsx`)

| Column A (DisplayName) | Column B (Config Template) | Column C (Model Data Excel) | Column D (Master CAD Folder) | Column E (Geometry Image) |
|---|---|---|---|---|
| ProductA | ProductA_Config.xlsx | ProductA_Models.xlsx | ProductA/ | ProductA.png |
| ProductB | ProductB_Config.xlsx | ProductB_Models.xlsx | ProductB/ | ProductB.png |

---

## Excel Template Format

The **Config Excel Template** (one per configurator) defines all parameters on **Sheet 1**. The application reads every row where **Column C = `INPUT`**.

| Col A (Name) | Col B (Value) | Col C (Type) | Col D (ControlType) | Col E (Validation Rule) | Col F (…) | Col G (IsVisible) | Col H (IsEnabled) | Col I (DependsOn) | Col J (TriggerValue) |
|---|---|---|---|---|---|---|---|---|---|
| Width | *(written at runtime)* | INPUT | TextBox | 100-2000 | | TRUE | TRUE | | |
| ConnectionType | *(written at runtime)* | INPUT | ComboBox | Flanged,Threaded,Welded | | TRUE | TRUE | | |
| FlangeSize | *(written at runtime)* | INPUT | ComboBox | DN50,DN80,DN100 | | FALSE | FALSE | ConnectionType | Flanged |

- **Column B** is left blank in the template — the app fills it in at runtime before calling CadCaster.
- **Column G / H** (`IsVisible` / `IsEnabled`): Set to `FALSE` (case-insensitive) to hide/disable a parameter by default.
- **Column I / J** (`DependsOn` / `TriggerValue`): Conditional visibility — the parameter becomes active only when the named controller parameter equals the trigger value.

---

## Parameter Control Types

| ControlType | UI Rendered | Value stored as |
|---|---|---|
| `TextBox` | Single-line text input | String (numeric or text) |
| `ComboBox` | Dropdown with allowed values | Selected string |
| `CheckBoxList` | Multi-select checkbox list | Comma-separated string |

The **Validation Rule** column (Column E) drives both the `AllowedValues` list (for ComboBox/CheckBoxList) and the min/max range (for TextBox).

---

## Parameter Validation Rules

Validation rules are parsed from **Column E** of the Config Excel Template:

| Rule Format | Example | Behaviour |
|---|---|---|
| `min-max` (numeric range) | `100-2000` | Value must parse as a number between 100 and 2000 (inclusive) |
| `val1,val2,val3` (allowed list) | `Flanged,Threaded,Welded` | Value must exactly match one of the listed options (case-insensitive) |
| *(empty)* | | No validation applied |

Validation errors are surfaced via `INotifyDataErrorInfo` and displayed inline in the WPF UI. The **Run** button is disabled until all visible, enabled parameters pass validation and the output path is set.

---

## Parameter Dependencies

Parameters can conditionally show/hide based on a controlling parameter's value:

- **`DependsOn`** (Column I): Name of the controlling parameter.
- **`TriggerValue`** (Column J): The exact value the controller must have for this parameter to become active.

When a controller parameter changes value:
1. The `DependencyManager` notifies all registered dependents.
2. Each dependent evaluates whether its `TriggerValue` matches; it sets `IsVisible` and `IsEnabled` accordingly.
3. If a dependent becomes inactive, its value is automatically cleared.

**Example**: `FlangeSize` is only visible when `ConnectionType = "Flanged"`.

---

## Model Data Management

The **Model Data Excel** file stores pre-configured parameter sets for a configurator. Its format is:

| Model (Col A) | Param1 (Col B) | Param2 (Col C) | … |
|---|---|---|---|
| Model_Small | 200 | Flanged | … |
| Model_Medium | 500 | Threaded | … |

- **Select model** — auto-populates all matching parameters in the UI by name-matching (case-insensitive).
- **Add** — saves the current parameter values as a new row (uses the first non-empty parameter value as the model name).
- **Update** — overwrites an existing model's row with the current UI values.
- **Delete** — removes the selected model's row from the Excel file.

The special entry `(New Model)` is always present at the top of the dropdown and indicates a fresh configuration.

---

## App Configuration

`ParametricBuilder/App.config` controls all runtime paths:

```xml
<appSettings>
  <!-- Directory containing SolidworksConfigurator.exe -->
  <add key="CadCasterPath" value="Console" />

  <!-- Default SolidWorks output directory -->
  <add key="OutPutDirectory" value="C:\ConfiguratorProject" />

  <!-- Root folder for Config.Data sub-directories -->
  <add key="ExcelFileDirectory" value="Config.Data" />

  <!-- Sub-folder name for config data (used for file discovery) -->
  <add key="ConfigData" value="Config.Data" />

  <!-- Path to the Mapping Excel (relative to executable) -->
  <add key="MappingExcel" value="Config.Data\MappingExcel_ParametricBuilder_V1.xlsx" />
</appSettings>
```

> **Note:** `CadCasterPath` must point to the folder containing `SolidworksConfigurator.exe`. Update this before first use.

---

## Tech Stack & NuGet Packages

| Package | Version | Purpose |
|---|---|---|
| .NET Framework | 4.7.2 | Runtime target |
| WPF | Built-in | Desktop UI framework |
| **ClosedXML** | 0.104.2 | Read/write `.xlsx` files without Excel interop |
| **DocumentFormat.OpenXml** | 3.1.1 | Underlying OpenXml SDK (ClosedXML dependency) |
| **MaterialDesignThemes** | 5.2.1 | Google Material Design controls & styles for WPF |
| **MaterialDesignColors** | 5.2.1 | Material Design color palettes |
| **Fluent.Ribbon** | 10.1.0 | Office-style ribbon toolbar |
| **ControlzEx** | 6.0.0 | Extended WPF controls (Material/Fluent dependency) |
| **NLog** | 5.2.8 | Structured logging with file + custom event targets |
| **Newtonsoft.Json** | 13.0.3 | JSON serialisation utilities |
| **Microsoft.Xaml.Behaviors.Wpf** | 1.1.39 | Attached behaviors (e.g., auto-scroll, multi-select) |
| **MSTest.TestAdapter** | 2.2.10 | MSTest runner integration |
| **MSTest.TestFramework** | 2.2.10 | Unit test framework |

---

## Getting Started

### Prerequisites

- **Windows 10/11** (64-bit recommended)
- **Visual Studio 2019 or later** with the **.NET desktop development** workload
- **.NET Framework 4.7.2** (typically included with Windows)
- **SolidWorks** with `SolidworksConfigurator.exe` installed and accessible
- **Microsoft Excel** is *not* required — Excel files are accessed via ClosedXML

### Setup

1. **Clone the repository**

   ```bash
   git clone https://github.com/Urjitpatel28/ParametricBuilder.git
   cd ParametricBuilder
   ```

2. **Restore NuGet packages**

   Open `ParametricBuilder.sln` in Visual Studio. NuGet packages will be restored automatically on first build. Alternatively:

   ```powershell
   nuget restore ParametricBuilder.sln
   ```

3. **Configure `App.config`**

   Update `ParametricBuilder/App.config` to point to your installation:

   ```xml
   <add key="CadCasterPath" value="C:\Path\To\SolidworksConfigurator" />
   <add key="OutPutDirectory" value="C:\SolidWorks\Output" />
   ```

4. **Set up Config.Data**

   Place your Excel files and CAD templates in the `Config.Data/` folder structure described in [Configuration Data Layout](#configuration-data-layout), relative to the built executable (typically `bin\Debug\` or `bin\Release\`).

5. **Build and Run**

   Press **F5** in Visual Studio or use:

   ```powershell
   msbuild ParametricBuilder.sln /p:Configuration=Release
   ```

---

## Building the Project

```powershell
# Debug build
msbuild ParametricBuilder.sln /p:Configuration=Debug

# Release build
msbuild ParametricBuilder.sln /p:Configuration=Release
```

Output binaries land in `ParametricBuilder\bin\Debug\` or `ParametricBuilder\bin\Release\`.

> Ensure the `Config.Data\` folder and its sub-directories are present alongside the built executable before running.

---

## Running Unit Tests

The `UnitTestParametricBuilder` project contains MSTest-based unit tests.

**Via Visual Studio:**  
Open **Test Explorer** (`Test → Test Explorer`) and click **Run All**.

**Via command line:**

```powershell
dotnet test UnitTestParametricBuilder\UnitTestParametricBuilder.csproj
```

---

## Logging

Log files are written to `logs/` next to the application executable:

```
<exe directory>/
└── logs/
    ├── logger_20260101.log
    ├── logger_20260102.log
    └── ...
```

**Log format:**

```
2026-01-01 09:15:32.4812|INFO|ParametricBuilder.Commands.RunCadCasterCommand|Workbook Saved Successfully.|
```

Log levels written: `DEBUG`, `INFO`, `WARN`, `ERROR`.

Logs are also streamed live to the **in-app log panel** (capped at 1000 entries). Duplicate messages within a single session are suppressed.

---

## Release History

| Version | Notes |
|---|---|
| **V1.2** | Latest release — improved parameter dependency handling, UI refinements |
| **V1.1** | Initial public release — core Excel-to-SolidWorks parametric automation |

Release archives are available in `ParametricBuilder/bin/` and `ParametricBuilder/Resource/`.

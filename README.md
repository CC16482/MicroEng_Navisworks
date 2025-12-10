# MicroEng_Navisworks - Developer Guide

## 1) Overview
- Purpose: custom automation toolkit for Autodesk Navisworks 2025.
- Assembly: `MicroEng.Navisworks.dll`
- Target: .NET Framework 4.8 (`net48`)
- UI: Dockable panel (`MicroEng.DockPane`) hosting the MicroEng buttons (Data Mapper, Reconstruct, Zone Finder, Data Scraper, Data Matrix toggle).

## 2) Repository Layout
```
MicroEng_Navisworks
+---MicroEng_Navisworks_tree.txt
+---MicroEng.Navisworks
|   +---AppendIntegrateDialog.cs
|   +---AppendIntegrateExecutor.cs
|   +---AppendIntegrateModels.cs
|   +---Class1.cs
|   +---MicroEng.Navisworks.csproj
|   +---MicroEngPlugins.cs
|   \---PropertyPickerDialog.cs
\---ReferenceProjects
    +---NavisAddinManager-dev
    \---NavisLookup-dev
```
- `MicroEng.Navisworks/`: Main plugin project loaded by Navisworks.
- `ReferenceProjects/`: Upstream tools kept for debugging/snooping. They are not built into the main DLL.

## 3) Prerequisites
- Autodesk Navisworks Manage 2025 (or Simulate 2025 with path tweaks).
- .NET 8 SDK (for `dotnet build`).
- .NET Framework 4.8 Developer Pack.
- VS Code with C# extensions.
- Permission to write into the Navisworks Plugins folder (default `C:\Program Files\Autodesk\Navisworks Manage 2025\Plugins`).

## 4) Project: MicroEng.Navisworks
- Target framework: `net48`.
- Navisworks references: `Autodesk.Navisworks.Api.dll`, `Autodesk.Navisworks.Api.Interop.ComApi.dll`.
- WinForms reference: `System.Windows.Forms`.
- MSBuild properties (override with `/p:` if needed):
  - `NavisApiDir` (defaults to `C:\Program Files\Autodesk\Navisworks Manage 2025`)
  - `NavisPluginsDir` (defaults to `C:\Program Files\Autodesk\Navisworks Manage 2025\Plugins`)
- AfterBuild copies the DLL to `$(NavisPluginsDir)\MicroEng.Navisworks\`.
- Main code files: `MicroEng.Navisworks/MicroEngPlugins.cs` (panel + add-ins) and `MicroEng.Navisworks/AppendIntegrate*.cs` (Append & Integrate Data dialog, execution, templates).

### Plugin entry points (do not rename IDs)
- `AppendDataAddIn` -> `MicroEng.AppendData` (displayed as **Data Mapper**)
- `ReconstructAddIn` -> `MicroEng.Reconstruct`
- `ZoneFinderAddIn` -> `MicroEng.ZoneFinder`
- `DataMatrixCommand` -> `MicroEng.DataMatrix.Command` (toggles the Data Matrix dock pane)
- `MicroEngPanelCommand` -> `MicroEng.PanelCommand` (toggles the dock pane)
- `MicroEngDockPane` -> `MicroEng.DockPane` (hosts `MicroEngPanelControl`)

## 5) Build and Deploy
```powershell
cd MicroEng.Navisworks
dotnet build
```
- Output: `MicroEng.Navisworks/bin/Debug/net48/MicroEng.Navisworks.dll`.
- The build copies the DLL into `$(NavisPluginsDir)\MicroEng.Navisworks\`. Override paths at build time if Navisworks is installed elsewhere:
```powershell
dotnet build `
  /p:NavisApiDir="D:\Navisworks 2025" `
  /p:NavisPluginsDir="D:\Navisworks 2025\Plugins"
```
- If you see “Could not resolve Autodesk.Navisworks.Api”, pass `NavisApiDir` explicitly as above (quotes required for spaces).

## 6) Running in Navisworks
1. Build the project so the DLL is in the Plugins folder.
2. Start Navisworks Manage 2025 and open a model.
3. In the Add-Ins tab you should see:
   - MicroEng Data Mapper
   - MicroEng Reconstruct
   - MicroEng Zone Finder
   - MicroEng Data Matrix
   - MicroEng Panel
4. Run `MicroEng Panel` to show the dockable panel with the buttons and log box, or `Data Matrix` to open the grid viewer dock pane.

## 7) Tool Behavior (current state)
- Data Mapper (formerly Append & Integrate Data):
  - Opens from the Data Mapper button in the MicroEng panel (and the Add-In command).
  - Template bar: pick template, New/Copy/Rename/Delete/Save. Templates persist to `append_templates.json` beside the DLL.
  - Standard tab: set Target Tab Name (default `MicroEng`); grid of rows (Target Property, Mode = Static/From Property/Expression, Source Property + picker, Static/Expression value, Option, Enabled).
  - Property picker: uses first selected item to browse its categories/properties; writes the chosen path into the row.
  - Advanced tab: apply to Items/Groups/Both; create/update target tab; delete blank properties/empty tab; apply to selection only; toggle internal property names for the picker.
  - Options available (UI + template): None, ConvertToDecimal, ConvertToInteger, ConvertToDate, FormatAsText, SumGroupProperty, UseParentProperty, UseParentRevitProperty, ReadAllPropertiesFromTab, ParseExcelFormula, PerformCalculation, ConvertFromRevit (most are stubs for now; only FormatAsText behavior is applied, the rest are placeholders for future parity with iConstruct).
  - Run: processes selection (or whole model if unchecked), creates/updates properties via COM API, applies simple FormatAsText handling, stubs the rest. Shows status + message box and logs summary to the docked log.
- Data Scraper:
  - Scans a chosen scope (single item, current selection, entire model) and caches results in `DataScraperCache`.
  - Stores `ScrapeSession` objects with a profile name, scope description, property catalog, and raw entries (item key/path, category, property, type, value).
  - UI tabs: Profile (latest run per profile), History (all runs), Properties (distinct property catalog), Raw Data (ungrouped per-item property rows).
  - Logs summary to the MicroEng panel; Append/Data Mapper uses the last session for property suggestions.
- Data Matrix:
  - Dockable grid viewer that **never re-scans** Navisworks; it pivots the cached ScrapeSession from Data Scraper into rows/columns.
  - Top toolbar: select Data Scraper profile, apply view presets, toggle column chooser, sync selection, filter to current selection, and export CSV (filtered/all).
  - Grid supports column sorting (with Shift for multi-sort), per-column filters, and column visibility presets. Selection can sync to Navisworks.
  - Presets are stored in-memory per profile (save, save-as, delete). Exports include profile/scope metadata in the CSV header.
- Reconstruct / Zone Finder: Placeholders; still show messages but ready for future logic.

## 8) Reference Projects (helpers only)
- `ReferenceProjects/NavisAddinManager-dev`: Hot-load/unload add-ins, run commands, view Debug/Trace inside Navisworks.
- `ReferenceProjects/NavisLookup-dev`: Inspect selections, properties, and Navisworks objects while developing.
- Reference documents: `ReferenceDocuments/Append_Data_Instructions.txt` (feature spec) and `ReferenceDocuments/iConstruct_Pro_Append_Data_&_Integrator.docx` (iConstruct behaviors for parity ideas).

## 9) Roadmap / Next Steps
- Zone Finder: add input (textbox/dropdown) on the panel and select items by property (Zone/Area/Level), then zoom to selection and log counts.
- Reconstruct: rebuild search sets or other derived data grouped by `MicroEng.Tag`.
- Architecture: move logic out of `MicroEngActions` into helper classes; consider WPF panel for richer UI; add JSON-based config for property names/zones.

## 10) Working with Codex
- Main file to edit: `MicroEng.Navisworks/MicroEngPlugins.cs`.
- Keep plugin IDs and class names unchanged.
- Prefer incremental changes: add helpers first, then wire into commands, then test via NavisAddinManager and NavisLookup.

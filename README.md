# MicroEng_Navisworks - Developer Guide

## 1) Overview
- Purpose: custom automation toolkit for Autodesk Navisworks 2025.
- Assembly: `MicroEng.Navisworks.dll`
- Target: .NET Framework 4.8 (`net48`)
- UI: Dockable panel (`MicroEng.DockPane`) hosting three tools:
  - `MicroEng.AppendData`
  - `MicroEng.Reconstruct`
  - `MicroEng.ZoneFinder`

## 2) Repository Layout
```
MicroEng_Navisworks
+---MicroEng_Navisworks_tree.txt
+---MicroEng.Navisworks
|   +---Class1.cs
|   +---MicroEng.Navisworks.csproj
|   \---MicroEngPlugins.cs
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
- Main code file: `MicroEng.Navisworks/MicroEngPlugins.cs`.

### Plugin entry points (do not rename IDs)
- `AppendDataAddIn` -> `MicroEng.AppendData`
- `ReconstructAddIn` -> `MicroEng.Reconstruct`
- `ZoneFinderAddIn` -> `MicroEng.ZoneFinder`
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
   - MicroEng Append Data
   - MicroEng Reconstruct
   - MicroEng Zone Finder
   - MicroEng Panel
4. Run `MicroEng Panel` to show the dockable panel with the three buttons and log box.

## 7) Tool Behavior (current state)
- Append Data: Checks current selection; if empty shows a friendly message. For each selected item it creates/reuses the `MicroEng` category and writes/updates a `Tag` string using the COM API (`ME-AUTO-{timestamp}`). Logs results to the docked log textbox (and Debug/Trace).
- Reconstruct / Zone Finder: Placeholders; still show messages but ready for future logic.

## 8) Reference Projects (helpers only)
- `ReferenceProjects/NavisAddinManager-dev`: Hot-load/unload add-ins, run commands, view Debug/Trace inside Navisworks.
- `ReferenceProjects/NavisLookup-dev`: Inspect selections, properties, and Navisworks objects while developing.

## 9) Roadmap / Next Steps
- Zone Finder: add input (textbox/dropdown) on the panel and select items by property (Zone/Area/Level), then zoom to selection and log counts.
- Reconstruct: rebuild search sets or other derived data grouped by `MicroEng.Tag`.
- Architecture: move logic out of `MicroEngActions` into helper classes; consider WPF panel for richer UI; add JSON-based config for property names/zones.

## 10) Working with Codex
- Main file to edit: `MicroEng.Navisworks/MicroEngPlugins.cs`.
- Keep plugin IDs and class names unchanged.
- Prefer incremental changes: add helpers first, then wire into commands, then test via NavisAddinManager and NavisLookup.

# MicroEng_Navisworks - Developer Guide

## Overview
- Navisworks 2025 automation toolkit written in C# targeting .NET Framework 4.8 (`net48`).
- Main assembly: `MicroEng.Navisworks.dll` with WPF UI (`UseWPF` + `UseWindowsForms` for ElementHost).
- Dockable MicroEng panel (`MicroEng.DockPane`) hosts buttons for Data Scraper, Data Mapper (Append Data), Data Matrix, and Space Mapper. Reconstruct / Zone Finder remain add-in stubs only.
- Data Matrix and Space Mapper are WPF dock panes; Data Scraper opens its own window; the panel log captures messages raised via `MicroEngActions.Log`.
- Branding/colours come from `MICROENG_THEME_GUIDE.md` and the `Logos` folder (`microeng_logotray.png`, `microeng-logo2.png`).

## Repository Layout (trimmed)
```
MicroEng_Navisworks
├─ MicroEng.Navisworks/              # Main plugin project
│  ├─ MicroEng.Navisworks.csproj     # net48, UseWPF, Navisworks refs
│  ├─ MicroEngPlugins.cs             # Plugin registrations, shared actions
│  ├─ MicroEngPanelControl.xaml(.cs) # Docked launcher + log
│  ├─ DataMatrixControl.xaml(.cs)    # WPF Data Matrix dock pane
│  ├─ SpaceMapperControl.xaml(.cs)   # WPF Space Mapper dock pane
│  ├─ AppendIntegrate*.cs            # Data Mapper (Append & Integrate) engine/templates
│  ├─ DataScraper*.cs                # Data Scraper window & services
│  ├─ SpaceMapper*.cs                # Space Mapper services/models/engines
│  └─ Logos\ (linked from repo root)
├─ Logos/                            # PNG assets for ribbon/panels
├─ ReferenceDocuments/               # Specs (Data Matrix, Space Mapper, etc.)
├─ ReferenceProjects/                # Sample/utility projects (not built)
└─ MICROENG_THEME_GUIDE.md
```

## Prerequisites
- Autodesk Navisworks Manage 2025 (or Simulate with path tweaks).
- .NET Framework 4.8 Developer Pack.
- .NET 8+ SDK for `dotnet build`.
- Permission to copy into `C:\Program Files\Autodesk\Navisworks Manage 2025\Plugins`.

## Build & Deploy
```powershell
cd C:\Users\Chris\Documents\GitHub\MicroEng_Navisworks
dotnet build MicroEng.Navisworks/MicroEng.Navisworks.csproj `
  /p:NavisApiDir="C:\Program Files\Autodesk\Navisworks Manage 2025"
# Optional if your plugins folder is different
# /p:NavisPluginsDir="C:\Program Files\Autodesk\Navisworks Manage 2025\Plugins"
```
- The build copies `MicroEng.Navisworks.dll` (and Logos) to `$(NavisPluginsDir)\MicroEng.Navisworks\`.
- `ResolveAssemblyReferenceAdditionalPaths` uses `NavisApiDir`, so point it to the folder that contains `Autodesk.Navisworks.Api.dll` and friends.

## Running in Navisworks
1. Build so the DLL is present in the Plugins folder.
2. Launch Navisworks Manage 2025 and load a model.
3. Add-Ins tab shows:
   - MicroEng Data Mapper (Append Data)
   - MicroEng Data Matrix (dockable)
   - MicroEng Space Mapper (dockable)
   - MicroEng Data Scraper
   - MicroEng Panel (toggles the docked launcher/log)
4. Reconstruct and Zone Finder are add-in placeholders only; they are not shown on the panel.

## Tool Behavior (current)
- **Append Data (Data Mapper button / `MicroEng.AppendData`)**: Tags the current selection. Creates/reuses the `MicroEng` custom tab and writes/updates a `Tag` string set to `ME-AUTO-<timestamp>` for each selected item. Logs counts to the docked log; shows a friendly message if nothing is selected.
- **Data Scraper**: Scans model properties into `DataScraperCache` (profiles, distinct properties, raw entries). Source of truth for Data Matrix/Space Mapper metadata.
- **Data Matrix** (dock pane `MicroEng.DataMatrix.DockPane`): WPF grid built from the latest ScrapeSession. Supports column chooser (toggle properties on/off), presets per profile, selection sync back to Navisworks, and CSV export (filtered/all).
- **Space Mapper** (dock pane `MicroEng.SpaceMapper.DockPane`): WPF UI per `ReferenceDocuments/Space_Mapper_Instructions.txt` with zones/targets, processing settings, attribute mapping, and results. Uses Data Scraper metadata and CPU intersection engine (GPU stubbed for now).
- **MicroEng Panel**: WPF launcher with branding/logo and log textbox that listens to `MicroEngActions.LogMessage`.

## Notes
- Theme/branding: follow `MICROENG_THEME_GUIDE.md` and use logos in `Logos/`.
- Data Matrix & Space Mapper use WPF hosted in ElementHost for dock panes; the main panel is also WPF.
- If Navisworks assemblies are not found at build time, set `NavisApiDir` to the root install folder containing `Autodesk.Navisworks.Api.dll`.

## Known issues / tips
- Space Mapper UI clipping: keep `SpaceMapperControl.xaml` rooted in a simple Grid (header row + content row) and ensure the dock pane is `[DockPanePlugin(..., FixedSize=false, AutoScroll=true)]` with `ElementHost` docked fill. If the pane looks truncated, delete old plugin copies from `C:\Program Files\Autodesk\Navisworks Manage 2025\Plugins\MicroEng.Navisworks\` (and any AppData bundle), rebuild, redeploy, then dock the pane before resizing.
 - CPU Normal mode only: Space Mapper currently forces CPU Normal processing. Geometry extraction uses bbox fallback; COM triangle extraction TODO remains for Navis main thread only.

## Working with Codex
- Primary entry points live in `MicroEngPlugins.cs`.
- UI work: `MicroEngPanelControl.xaml`, `DataMatrixControl.xaml`, `SpaceMapperControl.xaml`.
- Do not rename plugin IDs/classes (`MicroEng.AppendData`, `MicroEng.DockPane`, `MicroEng.DataMatrix.DockPane`, `MicroEng.SpaceMapper.DockPane`, etc.).

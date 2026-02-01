# MicroEng_Navisworks - Developer Guide

## Human Notes (Start Here)
- Build: `dotnet build MicroEng.Navisworks/MicroEng.Navisworks.csproj /p:NavisApiDir="C:\Program Files\Autodesk\Navisworks Manage 2025"` (default deploy goes to `C:\ProgramData\Autodesk\Navisworks Manage 2025\Plugins\MicroEng.Navisworks\` and writes `MicroEng.Navisworks.addin` to the **Plugins root**).
- If ProgramData deploy is locked, close Navisworks or add `/p:DeployToProgramData=false` to the build.
- Runtime: launch Navisworks and use Add-Ins tab (MicroEng Panel opens the launcher/log; Data Scraper/Mapper/Matrix/Space Mapper are separate windows/panes).
- WPF-UI is active (4.1.0). Theme is per-root only (no global App.xaml). Defaults: Dark theme, Black/White accent, DataGrid gridlines #C0C0C0.
- WPF-UI is the primary UI system. If something looks off, fix WPF-UI usage/styles first instead of falling back to non-WPF-UI controls.
- Settings: open via the gear button in the panel. Theme toggle + accent mode (System, Custom, Black/White) + DataGrid gridline color apply live to all open MicroEng windows.
- Step 3 Benchmark & Testing: Benchmark button runs multi-preset comparisons (Compute only / Simulate writeback / Full writeback). Options include writeback strategy, skip unchanged (signature), pack outputs, show internal properties, and close Navisworks panes (restore after). Reports save to `C:\ProgramData\Autodesk\Navisworks Manage 2025\Plugins\MicroEng.Navisworks\Reports\`.
- Advanced Performance: Fast Traversal (Auto/Zone-major/Target-major). Target-major is only valid when partial options are off.
- Smart Set Generator: dockable panel for Search/Selection Sets with Quick Builder, Smart Grouping, From Selection, and Packs. Uses Data Scraper cache for pickers/fast preview; recipes save to `C:\ProgramData\Autodesk\Navisworks Manage 2025\Plugins\MicroEng.Navisworks\SmartSets\Recipes\`.
- Quick Colour profiles are stored in `%APPDATA%\MicroEng\ColourProfiles\` (JSON schema v1). Palettes include Default (Deep), Custom, Shades, and fixed palettes; Hue Groups UI is removed (legacy discipline map remains at `%APPDATA%\MicroEng\Navisworks\QuickColour\CategoryDisciplineMap.json`).
- Data Matrix view presets persist in `%APPDATA%\MicroEng\Navisworks\DataMatrix\ViewPresets.json`.
- Log file: `%LOCALAPPDATA%\MicroEng.Navisworks\NavisErrors\MicroEng.log` (fallback: `%TEMP%\MicroEng.log`).
- Crash reports: `%LOCALAPPDATA%\MicroEng.Navisworks\NavisErrors\CrashReports\` (auto-written on unhandled/dispatcher exceptions with recent log tail).
- If UI is blank or a window fails to open, check the log for XAML resource errors (missing resource keys or Wpf.Ui.dll not found).
- Known UI issue to verify: some labels/text still render black in Dark mode; fix by removing local `Foreground` overrides or ensuring `MicroEngWpfUiTheme.ApplyTextResources` updates `TextFillColor*` brushes.
- If the plugin does not load: verify `C:\ProgramData\Autodesk\Navisworks Manage 2025\Plugins\MicroEng.Navisworks.addin` points to `.\MicroEng.Navisworks\MicroEng.Navisworks.dll`, and that DLL actually exists in `C:\ProgramData\Autodesk\Navisworks Manage 2025\Plugins\MicroEng.Navisworks\`. Do not mix ProgramData and Program Files plugin roots.
## AI Notes (Context for Codex)
- Priority #1: WPF-UI is the primary UI system. Do not introduce non-WPF-UI elements as a fallback unless absolutely required for stability; instead, fix/adjust WPF-UI usage to match Gallery patterns when possible.
- WPF-UI Gallery reference is local: `ReferenceProjects\wpfui-main`. Use patterns from:
  - `src\Wpf.Ui.Gallery\Views\Pages\BasicInput\CheckBoxPage.xaml`
  - `src\Wpf.Ui.Gallery\Views\Pages\BasicInput\ComboBoxPage.xaml`
  - `src\Wpf.Ui.Gallery\Views\Pages\BasicInput\ToggleSwitchPage.xaml`
  - `src\Wpf.Ui.Gallery\Views\Pages\Collections\TabControlPage.xaml`
  - `src\Wpf.Ui.Gallery\Views\Pages\DesignGuidance\ColorsPage.xaml`
- WPF-UI theme bootstrapping is per-root via `MicroEngWpfUiTheme.ApplyTo(root)` (no global App.xaml). `MicroEngWpfUiRoot.xaml` merges WPF-UI dictionaries + `MicroEngUiKit.xaml`.
- `MicroEngUiKit.xaml` provides shared spacing/typography/card/button styles but should not override WPF-UI implicit styles for `ComboBox`, `CheckBox`, `RadioButton`, etc.
- `MicroEngWpfUiTheme.cs` handles theme/accents/gridlines and broadcasts changes to registered roots; uses `TextFillColor*` resources to fix dark-mode text.
- Primary UI files: `MicroEngPanelControl.xaml`, `DataScraperWindow.xaml`, `AppendIntegrateDialog.xaml`, `DataMatrixControl.xaml`, `SpaceMapperControl.xaml`.
- Data Matrix Column Builder window: `DataMatrixColumnBuilderWindow.xaml(.cs)` (tri-state category tree + search + chosen list).
- Quick Colour UI files: `QuickColour/QuickColourControl.xaml`, `QuickColour/QuickColourWindow.xaml`, `QuickColour/Profiles/QuickColourProfilesPage.xaml`.
- Smart Set UI files: `SmartSets/SmartSetGeneratorControl.xaml(.cs)` and `SmartSets/PropertyPickerWindow.xaml(.cs)`.

## Overview
- Navisworks 2025 automation toolkit written in C# targeting .NET Framework 4.8 (`net48`).
- Main assembly: `MicroEng.Navisworks.dll` with WPF UI (`UseWPF` + `UseWindowsForms` for ElementHost).
- Dockable MicroEng panel (`MicroEng.DockPane`) hosts buttons for Data Scraper, Data Mapper (Append Data), Data Matrix, Quick Colour, Smart Set Generator, Space Mapper, and 4D Sequence. Reconstruct / Zone Finder remain add-in stubs only.
- Data Matrix, Smart Set Generator, and Space Mapper are WPF dock panes; Data Scraper opens its own window; the panel log captures messages raised via `MicroEngActions.Log`.
- Branding/colours come from `MICROENG_THEME_GUIDE.md` and the `Logos` folder (`microeng_logotray.png`, `microeng-logo2.png`).

## Repository Layout (trimmed)
```
MicroEng_Navisworks/
- MicroEng.Navisworks/              # Main plugin project
  - MicroEng.Navisworks.csproj     # net48, UseWPF, Navisworks refs
  - MicroEngPlugins.cs             # Plugin registrations, shared actions
  - MicroEngPanelControl.xaml(.cs) # Docked launcher + log
  - QuickColour/                   # Quick Colour window, palettes, profiles
  - SmartSets/                     # Smart Set Generator (dock pane + services)
  - DataMatrixControl.xaml(.cs)    # WPF Data Matrix dock pane
  - SpaceMapperControl.xaml(.cs)   # WPF Space Mapper dock pane
  - AppendIntegrate*.cs            # Data Mapper (Append & Integrate) engine/templates
  - DataScraper*.cs                # Data Scraper window & services
  - SpaceMapper*.cs                # Space Mapper services/models/engines
  - Logos/ (linked from repo root)
- Logos/                            # PNG assets for ribbon/panels
- ReferenceDocuments/               # Specs (Data Matrix, Space Mapper, etc.)
- ReferenceProjects/                # Sample/utility projects (not built)
- MICROENG_THEME_GUIDE.md
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
- The build copies `MicroEng.Navisworks.dll` (and dependencies + Logos) to `$(NavisPluginsDir)\MicroEng.Navisworks\`.
- The `.addin` manifest must be located in the **Plugins root**: `$(NavisPluginsDir)\MicroEng.Navisworks.addin`. Navisworks will not discover add-ins if the `.addin` lives only in the subfolder.
- `ResolveAssemblyReferenceAdditionalPaths` uses `NavisApiDir`, so point it to the folder that contains `Autodesk.Navisworks.Api.dll` and friends.

## Running in Navisworks
1. Build so the DLL is present in the Plugins folder.
2. Launch Navisworks Manage 2025 and load a model.
3. Add-Ins tab shows:
   - MicroEng Data Mapper (Append Data)
   - MicroEng Data Matrix (dockable)
   - MicroEng Quick Colour (window)
   - MicroEng Smart Set Generator (dockable)
   - MicroEng Space Mapper (dockable)
   - MicroEng 4D Sequence (dockable)
   - MicroEng Data Scraper
   - MicroEng Panel (toggles the docked launcher/log)
4. Reconstruct and Zone Finder are add-in placeholders only; they are not shown on the panel.

## Tool Behavior (current)
- **Data Mapper (`MicroEng.AppendData`)**: Opens the Data Mapper UI (`AppendIntegrateDialog`). Uses Data Scraper cache for property pickers/type hints; no auto-tagging logic remains.
- **Data Scraper**: Raw capture only (no JSONL export); scans model properties into `DataScraperCache` (profiles, distinct properties, raw entries). Selection/Search set scopes supported via `NavisworksSelectionSetUtils`.
- **Data Matrix** (dock pane `MicroEng.DataMatrix.DockPane`): Schedule-style builder. Columns are selected via **Column Builder**; default view opens with no columns until you choose or load a preset. Supports scope (Entire/Current/Set/Search/Single), persistent presets, selection sync back to Navisworks, CSV export, and JSONL export (Item documents or Raw rows, optional `.jsonl.gz`).
- **Quick Colour** (window): Three tabs (Profiles, Quick Colour, Hierarchy Builder). Profiles tab applies saved MicroEng Colour Profiles (JSON in AppData). Quick Colour builds a profile from Category+Property using Data Scraper cache, palette selection (Default/Custom/Shades + fixed palettes), stable colors, and scope options that mirror Smart Sets. Hierarchy Builder generates base+type colors with Shade Groups, single-hue mode, and type shade controls. Apply supports temp/permanent overrides by Category + Type (dual-property AND search). Values table supports per-row color picking.
- **Smart Set Generator** (dock pane `MicroEng.SmartSetGenerator.DockPane`): Quick Builder rules, smart grouping, selection-driven suggestions, and pack presets. Property picker and fast preview use Data Scraper cache; recipes persist to ProgramData.
- **Space Mapper** (dock pane `MicroEng.SpaceMapper.DockPane`): WPF UI per `ReferenceDocuments/Space_Mapper_Instructions.txt` with zones/targets, processing settings, attribute mapping, and results. Uses Data Scraper metadata with CPU + GPU backends (D3D11/CUDA) and a GPU sampling option; falls back to CPU when GPU is unavailable. Presets Fast/Normal/Accurate; Fast uses origin-point containment (bbox center in zone AABB). Advanced Performance includes Fast Traversal (Auto/Zone-major/Target-major). Step 3 includes Benchmark & Testing (compute/simulate/full), writeback strategy, skip-unchanged signature, packed writeback option, and optional pane-closing during runs.
- **4D Sequence** (dock pane `MicroEng.Sequence4D.DockPane`): WPF UI that captures selection, orders targets, and generates Timeliner task sequences; includes delete-by-name for regeneration.
- **MicroEng Panel**: WPF launcher with branding/logo and log textbox that listens to `MicroEngActions.LogMessage`.

## Notes
- Theme/branding: follow `MICROENG_THEME_GUIDE.md` and use logos in `Logos/`.
- Data Matrix & Space Mapper use WPF hosted in ElementHost for dock panes; the main panel is also WPF.
- If Navisworks assemblies are not found at build time, set `NavisApiDir` to the root install folder containing `Autodesk.Navisworks.Api.dll`.

## Known issues / tips
- Space Mapper UI clipping: keep `SpaceMapperControl.xaml` rooted in a simple Grid (header row + content row) and ensure the dock pane is `[DockPanePlugin(..., FixedSize=false, AutoScroll=true)]` with `ElementHost` docked fill. If the pane looks truncated, delete old plugin copies from `C:\Program Files\Autodesk\Navisworks Manage 2025\Plugins\MicroEng.Navisworks\` (and any AppData bundle), rebuild, redeploy, then dock the pane before resizing.
- Target-major traversal only applies when partial options are off (Fast origin-point mode). When partial tagging is enabled, Fast traversal falls back to Zone-major.
- Data Scraper: verify Selection/Search set scopes populate and resolve items correctly.
- Smart Grouping: preview uses Data Scraper cache and does not apply grouping scope (generation does apply scope).
- Quick Colour: scope filtering + profile scope encoding need verification across selection set/tree/property filter.
- Quick Colour: verify palette ordering/labels, shade expansion, and Shade Groups toggle behavior.

## Working with Codex
- Primary entry points live in `MicroEngPlugins.cs`.
- UI work: `MicroEngPanelControl.xaml`, `DataMatrixControl.xaml`, `SpaceMapperControl.xaml`, `SmartSets/SmartSetGeneratorControl.xaml`.
- Do not rename plugin IDs/classes (`MicroEng.AppendData`, `MicroEng.DockPane`, `MicroEng.DataMatrix.DockPane`, `MicroEng.SpaceMapper.DockPane`, etc.).

## Codex Handoff (Smart Set Generator)
- Scope + rules grid fixes from `ReferenceDocuments/Codex_Instructions_02.txt` are implemented: scope UI (Search in + Scope mode + saved selection set + use current selection + clear), scope applied in search creation, fast preview disabled when scope constrained, rules grid combo columns with single-click open, and Value disabled for Defined/Undefined.
- Scope picker window is implemented with WPF-UI styling; model-tree scopes use tri-state checkbox tree selection with a scope summary.
- Smart Grouping has its own output/scope settings and uses WPF-UI dropdowns; Include Blanks toggles empty-set generation for Quick Builder + Smart Grouping.
- Property picker dialog is implemented (category/property search + samples); fast preview uses Data Scraper cache.
- Recipe save/load persists JSON to `C:\ProgramData\Autodesk\Navisworks Manage 2025\Plugins\MicroEng.Navisworks\SmartSets\Recipes\`.

## Codex Handoff (Quick Colour)
- Profiles tab uses MicroEng Colour Profiles (schema v1) saved under `%APPDATA%\MicroEng\ColourProfiles\` and supports Apply Temporary/Permanent, Import/Export/Delete.
- Quick Colour tool includes Hierarchy Builder (Shade Groups toggle, single-hue mode, type shade controls). Hue Groups UI removed (discipline map file remains at `%APPDATA%\MicroEng\Navisworks\QuickColour\CategoryDisciplineMap.json`).
- Quick Colour values table supports per-row color picking; Profiles preview shows color swatches.
- Custom Hue palette keeps the Deep hue assignment but shifts lightness toward the picked base color.
- Scope options mirror Smart Sets (All model, Current selection, Saved selection set, Tree selection, Property filter); Load Values/Hierarchy and Apply respect scope; scope kind/path persist in profiles.
- Profiles tab apply uses saved category/property (Quick Colour) or saved Level 1/2 (Hierarchy) when applying a profile.

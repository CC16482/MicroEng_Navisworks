# MicroEng Handoff

## Status
- Last local build: succeeded on **February 9, 2026** (`dotnet build "MicroEng_Navisworks.sln" /p:DeployToProgramData=false`).
- Last ProgramData deploy build: succeeded on **February 9, 2026** (`dotnet build "MicroEng_Navisworks.sln"`).
- Navisworks restart after deploy confirmed add-in startup (see `%LOCALAPPDATA%\\MicroEng.Navisworks\\NavisErrors\\MicroEng.log` around `2026-02-09 17:24:04`).
- Recommended: close Navisworks or use `dotnet build /p:DeployToProgramData=false` if ProgramData deploy is locked.
- Working tree is dirty (build outputs in bin/obj + Data Matrix refactor/column builder + crash-report logging + Quick Colour profile apply fix + Data Scraper cache refactor/perf fixes).
- Storage is now unified under `%APPDATA%\\MicroEng\\Navisworks\\DataStore\\` (legacy locations auto-migrate on load).
- Changing PCs: copy `%APPDATA%\\MicroEng\\Navisworks\\DataStore\\` and re-apply TreeSpec/COM registry fixes (see TreeMapper notes below).
- Benchmark reports are saved under `C:\ProgramData\Autodesk\Navisworks Manage 2025\Plugins\MicroEng.Navisworks\Reports\`.
- Smart Set recipes saved under `%APPDATA%\\MicroEng\\Navisworks\\DataStore\\SmartSetRecipes\\` (legacy ProgramData recipe folder auto-migrates).
- Quick Colour profiles saved under `%APPDATA%\\MicroEng\\Navisworks\\DataStore\\QuickColourProfiles.json` (legacy `%APPDATA%\\MicroEng\\ColourProfiles\\` auto-migrates).
- Data Scraper metadata saved under `%APPDATA%\\MicroEng\\Navisworks\\DataStore\\ScrapeSessions.json` (v2 metadata-only store).
- Data Scraper raw rows saved per session under `%APPDATA%\\MicroEng\\Navisworks\\DataStore\\DataScraperRaw\\*.json`.
- Crash reports saved under `%LOCALAPPDATA%\MicroEng.Navisworks\NavisErrors\CrashReports\`.

## Recent changes
- TreeMapper Selection Tree dropdown now appears; breakthrough was setting COM Plugin License to DWORD (HKCU 22.0 was REG_SZ) and using the Roaming TreeSpec path (%APPDATA%\\Autodesk\\Navisworks Manage 2025\\TreeSpec). No need to rerun the installer once these are set.
- TreeMapper tool UI implemented with Data Matrix-style header and profile editor + preview; profiles persist and reload (Save/Save As/Reload).
- TreeMapper publish writes `PublishedTree.json` snapshot; COM Selection Tree renders from snapshot and supports selection.
- TreeMapper publish now stamps DocumentFile/DocumentFileKey (fallback to active document) so trees are model-bound.
- Selection Tree COM plugin strips `Int32:`/type prefixes from display values (display-only).
- Selection Tree COM plugin reloads snapshot when file length or write time changes to avoid stale root-only view (fixes “only Floors” after publish).
- TreeMapper UI shifted to vertical flow (profile -> explanation -> builder -> preview) with SizeToContent height.
- TreeMapper node type labels aligned to Navisworks terms (File, Composite Object, Insert, Instance).
- TreeMapper node type icons added in builder grid and preview tree (WPF UI symbols).
- TreeMapper Collection icon updated to Connected20 (matches Navisworks “Collection” symbol).
- Column Builder: Filter-by-selection now **scans only selected keys** (no full item→property index build) to avoid long GC/UI stalls; scan cancels/restarts on selection changes.
- Column Builder: When selection filter is active, only **expanded categories** refresh their property view; collapsed categories are just hidden/shown.
- Column Builder: When selection is empty, the tree is hidden and **no processing** should occur (fast path).
- Column Builder (Data Matrix): left-side TreeView scrolling tuned with deferred scrolling + extra virtualization to reduce layout stalls on large categories.
- Data Matrix: schedule-style refactor with scope culling, persistent view presets (AppData), and JSONL export (Item docs/Raw rows, optional .jsonl.gz).
- Data Matrix: default view opens with no columns; "Columns..." replaced by "Column Builder".
- Data Matrix: column header beaker filter implemented with per-column filter builder (AND/OR joins, text/numeric/date/regex operators, per-rule case/trim).
- Data Matrix: filter rules persist in View presets; beaker icons reflect active filters.
- Data Matrix: filter evaluation updated to honor per-rule join, case sensitivity, and trim settings.
- Data Matrix: column header hover/pressed styling forced to dark mode; filter prompt dialog forced to dark theme.
- Data Matrix Filter Builder window added (WPF Window): compact layout, Add Rule button, inline rule editing, delete (Dismiss12), join disabled when no enabled rule above.
- Data Matrix Column Builder: split window with tri-state category tree + search on left, chosen columns list with two-state toggles on right.
- Data Matrix Column Builder: right list uses ItemsControl + plain CheckBox to avoid WPF-UI animation crashes when toggling filtered items.
- Data Matrix: dedupe attribute IDs during build to avoid duplicate Category|Property collisions.
- Data Scraper: Selection/Search set scopes implemented via NavisworksSelectionSetUtils (raw capture only; JSONL stays in Data Matrix).
- Crash reporting: unhandled/dispatcher exceptions write crash files with log tail + dock pane visibility.
- Quick Colour: Profiles tab apply uses saved category/property (Quick Colour) or Level props (Hierarchy) when applying saved profiles.
- Data Scraper UI rebuilt into step cards (Profile, Scope, History, Data View, Export) with WPF-UI styling, scroll, and tighter spacing.
- Data Scraper: Run Scrape/Export Now now use ArrowSync icon during processing; status shows “Processing - Please Wait” in orange while running.
- Data Scraper: Export requires Output path to be set before enabling; Export Now restores scroll position after completion.
- Data Scraper: cache store refactored to metadata + per-session raw files (`ScrapeSessions.json` + `DataScraperRaw/*.json`).
- Data Scraper: history grid raw count now binds `RawEntryCount` (prevents forced raw-row deserialization on window open).
- Data Scraper: raw rows now lazy-load only when the Raw Data tab is opened for a selected run.
- Data Scraper: switching runs now releases previously loaded raw rows from memory.
- Data Scraper: `X` action in History removes old cached runs (metadata + per-session raw file).
- Data Scraper: save path now retries file replace to reduce transient file-lock save failures.
- Data Scraper: Data View tab content now uses cards (no white border).
- Snackbars added across tools (success/error) with larger PresenceAvailable icon, centered, black text, no close X.
- Quick Colour: hierarchy value matching switched to ItemKey (fixes missing types).
- Quick Colour: Hierarchy Builder palette dropdown now matches Quick Colour palettes (Default/Custom/Shades + fixed palettes), with hue picker enabled for Custom/Shades.
- Quick Colour: Apply now normalizes value prefixes and uses typed VariantData (int/bool/double) to match properties.
- Quick Colour: Apply debug logging added (per-value/per-pair hits + expected vs actual counts).
- Viewpoints Generator tool added (dock pane + command) with Smart Sets-style theming.
- Viewpoints Generator UI added: source mode, output folder, view direction, projection, selection/search set picker, preview plan, generate.
- Viewpoints Generator service: load selection/search sets, build plan, fit camera to bounds, create Saved Viewpoints under folder path.
- MicroEng panel now includes Viewpoints Generator card button (Camera24 icon) and state tracking.
- Quick Colour Profiles tab added (library + Apply Temporary/Permanent + Import/Export/Delete).
- MicroEng Colour Profile schema v1 stored in AppData; Profiles preview shows color swatches.
- Quick Colour palettes updated: Deep renamed to Default (top of list), Custom second, Shades added, and fixed palettes added (Beach, Ocean Breeze, Vibrant, Pastel, Autumn, Red Sunset, Forest Hues, Purple Raindrops, Light Steel, Earthy Brown, Earthy Green, Warm Neutrals 1/2, Candy Pop). Fixed palettes scale using light/dark variants when more values are needed.
- Quick Colour palette supports Custom Hue (pick base color; Deep hue assignment retained, lightness adjusted).
- Quick Colour values table swatches are editable with color picker.
- Quick Colour scope selector mirrors Smart Sets (All model, Current selection, Saved selection set, Tree selection, Property filter) with scope picker + summary.
- Load Values / Load Hierarchy and Apply now respect scope; scope kind/path persist in profiles.
- Quick Colour tool implemented (Hierarchy Builder tab + Shade Groups toggle + preview grid).
- Quick Colour applies temp/permanent overrides by Category + Type (dual-property AND search).
- Single-hue mode added with hue pick + category contrast bands; types shade within band.
- Hierarchy Builder: Hue Groups UI removed; Shade Groups toggle controls whether types are shaded around the base colour.
- Quick Colour window and command added; panel now includes Quick Colour card button.
- Smart Set Generator rules grid dropdowns now render with template ComboBoxes (Category/Property/Condition) and are clickable.
- Smart Set Scope Picker window added with WPF-UI styling, dark mode, and tri-state tree checkboxes for model-tree scopes.
- Smart Grouping now has its own output/scope settings, uses WPF-UI dropdowns, and generation applies grouping scope.
- Include Blanks option added to Quick Builder; generation skips empty sets when unchecked (also applied to Smart Grouping).
- Smart Set Generator checkboxes now use the WPF-UI default tickbox style (custom template removed).
- Data Scraper scope list reordered with Entire Model first and default selected.
- MicroEng Tools panel buttons now use card-style `ui:CardAction` with icons; added Smart Set Generator entry.
- Smart Set Generator dock pane + command added (`MicroEng.SmartSetGenerator.DockPane` / `MicroEng.SmartSetGenerator.Command`).
- Smart Set Generator UI includes Quick Builder, Smart Grouping, From Selection, and Packs tabs.
- Property picker dialog added (category/property search + samples); fast preview uses Data Scraper cache.
- Condition operators match Navisworks Find Items: Equals, not equals, Contains, Wildcard, Defined, Undefined.
- Recipe save/load implemented via JSON in ProgramData.
- Panel icons updated: Data Scraper `DatabaseLightning20`, Data Matrix `TableSettings28`, Space Mapper `BoxMultiple20`, Smart Sets `TextBulletListAdd20`.
- Fast preset uses origin-point classification (target bbox center in zone AABB); partial tagging only occurs when Tag Partial Separately or Treat Partial as Contained is enabled.
- Preflight signature and reuse updated to include processing mode and origin-point settings; point-based target indexing when partials are off.
- Advanced Performance adds Fast Traversal (Auto/Zone-major/Target-major). Target-major is only valid when partials are off.
- Step 3 now includes a Benchmark & Testing section:
  - Benchmark button with Compute only / Simulate writeback / Full writeback modes.
  - Writeback strategy selection (Virtual/Optimized/Legacy).
  - Latest benchmark summary shown in a dedicated card.
- Writeback performance options:
  - Skip unchanged targets via per-target signature.
  - Pack outputs into a single property to reduce COM writes.
- Close Navisworks panes during run (restore after): Selection Tree, Find Items, Properties, plus add-in dock panes.
- Close-pane delay is now user-configurable (sec). Progress window opens after panes close + delay, so UI can settle before run starts.
- Geometry planes: only create triangle planes when vertices.Count % 3 == 0 (prevents bogus planes from bbox vertices).
- New containment option: "Target geometry samples (GPU)" computes containment fraction using target mesh samples on GPU when available.
- Zone offsets: added enable toggle, offset-only pass, and "Zone Offset Match" writeback property (sequenced when needed).
- Offsets now apply in accurate containment (planes inflated + mesh probes) and in GPU sampling (CUDA/D3D11).
- GPU batching: D3D11 multi-zone dispatch + adaptive thresholds and pack thresholds; default max batch points = 200k.
- CUDA: brute-force batching (TestPointsBatched) + native update; optional CUDA BVH backend with per-batch scene build.
- GPU backend selection: CUDA BVH -> D3D11 batched -> CUDA brute -> CPU fallback.
- GPU diagnostics: batch metrics and per-zone GPU eligibility table (skip reasons, thresholds) in run reports.
- Variation check: baseline CPU variant + GPU variant added for direct CPU vs GPU comparison.
- Column Builder: modeless keyboard input fixed by calling `ElementHost.EnableModelessKeyboardInterop` when showing the Column Builder window from the Data Matrix dock pane.

## TreeMapper Selection Tree - detailed debugging notes (for handoff)
### Goal
- Expose a custom Selection Tree entry ("TreeMapper") backed by a COM plugin (InwOpUserSelectionTreePlugin) and a .spc spec file, so the dropdown shows a new tree in Navisworks.

### Key assemblies, IDs, and files
- COM plugin assembly: `C:\\ProgramData\\Autodesk\\Navisworks Manage 2025\\Plugins\\MicroEng.Navisworks\\MicroEng.SelectionTreeCom.dll`
- Installer/add-in assembly: `C:\\ProgramData\\Autodesk\\Navisworks Manage 2025\\Plugins\\MicroEng.Navisworks\\MicroEng.Navisworks.dll`
- COM ProgID: `MicroEng.TreeMapperSelectionTreePlugin`
- COM CLSID: `{9D4B2B3D-36B5-4A4E-8B79-9D8E6B7D6C01}`
- COM class: `MicroEng.SelectionTreeCom.TreeMapperSelectionTreePlugin`
- COM registration (effective): `HKLM\\Software\\Classes\\CLSID\\{9D4B2B3D-36B5-4A4E-8B79-9D8E6B7D6C01}\\InprocServer32`
  - Default = `mscoree.dll`
  - CodeBase = `file:///C:/ProgramData/Autodesk/Navisworks Manage 2025/Plugins/MicroEng.Navisworks/MicroEng.SelectionTreeCom.dll`
- Tree spec files (final location): `%APPDATA%\\Autodesk\\Navisworks Manage 2025\\TreeSpec\\`
  - `TreeMapper.spc`
  - `MicroEng.TreeMapperSelectionTreePlugin.spc`
  - `LcUntitledPlugin.spc`
- User Tree Specs registry keys (final):
  - `HKCU\\Software\\Autodesk\\Navisworks Manage\\22.0\\User Tree Specs`
  - `HKLM\\Software\\Autodesk\\Navisworks Manage\\22.0\\User Tree Specs`
  - (also wrote the same under `23.0`, but ProcMon showed 22.0 is what Navisworks actually reads for this flow)
- COM plugin license key (critical):
  - `HKCU\\Software\\Autodesk\\Navisworks Manage\\22.0\\COM Plugins\\MicroEng.TreeMapperSelectionTreePlugin`
  - `HKLM\\Software\\Autodesk\\Navisworks Manage\\22.0\\COM Plugins\\MicroEng.TreeMapperSelectionTreePlugin`
  - Must be **REG_DWORD = 1** (REG_SZ does NOT work)

### Logs and diagnostics used
- Main add-in log: `C:\\Users\\Chris\\AppData\\Local\\MicroEng.Navisworks\\NavisErrors\\MicroEng.log`
- COM plugin log: `MicroEng.SelectionTreeCom.log` (enabled in `MicroEng.SelectionTreeCom/TreeMapperSelectionTreeCom.cs`; check if file is emitted in the same NavisErrors folder)
- ProcMon filters:
  - Process = `Roamer.exe`
  - Operations = `CreateFile`, `RegOpenKey`, `RegQueryValue`, `RegSetValue`
  - Path contains: `TreeSpec`, `User Tree Specs`, `Selection Tree`, `TreeMapper`, `LcUntitledPlugin.spc`
- ProcMon showed Navisworks reading 22.0 registry keys and trying to open .spc files (helped locate wrong registry hive and wrong TreeSpec path).

### Symptoms (what was failing)
- Selection Tree dropdown never showed "TreeMapper".
- Navisworks popup: "third party COM plugin MicroEng.TreeMapperSelectionTreePlugin could not be loaded" (prompt to disable).
- Installer logs repeatedly showed:
  - `CreatePlugin(nwOpUserSelectionTreePlugin) failed: E_ACCESSDENIED`
  - `SaveFileToAppSpecDir` errors or success but still no dropdown
- COM plugin registration existed, but plugin still failed to load or was ignored by Selection Tree.

### What we tried (timeline summary)
1) **COM registration + installer writes to Program Files**  
   - Wrote TreeSpec files to `C:\\Program Files\\Autodesk\\Navisworks Manage 2025\\TreeSpec`.  
   - Added User Tree Specs values under 23.0.  
   - Result: no dropdown.

2) **Run Navisworks as admin / rerun installer**  
   - Reduced E_ACCESSDENIED for file writes but still no dropdown.  
   - Result: no dropdown.

3) **ProcMon deep dive**  
   - Found Navisworks reads **22.0** registry keys (COM Plugins + User Tree Specs).  
   - Found TreeSpec was probed in **Roaming** path, not Program Files.  
   - Found COM plugin license value under 22.0 existed but as **REG_SZ** (string), not DWORD.

4) **Registry + path fixes (breakthrough)**  
   - Set COM plugin license to `REG_DWORD = 1` under HKCU and HKLM for `22.0\\COM Plugins`.  
   - Moved / wrote TreeSpec files to `%APPDATA%\\Autodesk\\Navisworks Manage 2025\\TreeSpec`.  
   - Updated `User Tree Specs` values (22.0) to Roaming paths.  
   - Result: "TreeMapper (stub)" appeared in Selection Tree dropdown.

5) **Code changes (kept)**  
   - `MicroEng.Navisworks/SelectionTreeCom/DevTreeMapperSelectionTreeInstaller.cs`  
     - Added `GetUserTreeSpecDir()` and wrote .spc to Roaming TreeSpec dir.  
   - `MicroEng.Navisworks/SelectionTreeCom/SelectionTreeComRegistrar.cs`  
     - Ensured 22.0 + 23.0 registry coverage (COM Plugins + User Tree Specs).  
   - `MicroEng.SelectionTreeCom/TreeMapperSelectionTreeCom.cs`  
     - Added logs in COM constructor + `iCreateInterface`, `iGetNumRootChildren`, `iGetName` to confirm activation and UI calls.

### Known good configuration (final)
- COM plugin license keys (HKCU + HKLM, 22.0) are DWORD=1.
- TreeSpec files live in `%APPDATA%\\Autodesk\\Navisworks Manage 2025\\TreeSpec`.
- User Tree Specs (22.0) point to those Roaming paths.
- COM registration points to `MicroEng.SelectionTreeCom.dll` under ProgramData.
- No need to rerun the "Install TreeMapper Selection Tree Option" once these are set.
- TreeMapper snapshot lives at `%APPDATA%\\MicroEng\\Navisworks\\TreeMapper\\PublishedTree.json`.
- PublishedTree now includes DocumentFile/DocumentFileKey; Selection Tree only loads when current doc matches.

### Reference: Why this matters for future debugging
- If Selection Tree dropdown is empty again, first verify:
  - License DWORD under 22.0.
  - TreeSpec path in Roaming.
  - User Tree Specs values in 22.0.
  - COM registration still points to the correct CodeBase.
  - PublishedTree.json exists and matches current document file name.

### TreeMapper runtime notes (latest)
- Profiles now persist across Navisworks restarts; reload is immediate via dropdown/Reload.
- Publish is required to update the Selection Tree; after publish, Selection Tree loads from snapshot (no live queries).
- If Selection Tree shows only one root after publish, restart Navisworks or ensure snapshot reload (length+timestamp check).
- “(Missing)” nodes appear for items missing a property at any level; selection count in UI is unrelated.
  - If desired, we can add display counts or hide Missing nodes when count is zero.

## Open issues
- TreeMapper: enforce File as the first/only header row (IsFileHeader exists in VM but not wired in window logic yet).
- TreeMapper: rename preview header to “Tree Preview” and confirm window height shrinks when table shrinks.
- **Column Builder filter-by-selection still freezing** after selection results appear in some runs. Previous logs showed UI stalls during selection index build; new selected-keys scan fix needs validation on new PC.
- Data Scraper: verify Selection/Search set scopes populate and resolve items correctly.
- Data Scraper: validate open performance on large history after v2 cache migration (first load can be longer; subsequent opens should be fast).
- Data Scraper: verify Raw Data tab lazy-load behavior and memory release when switching sessions.
- Data Matrix Column Builder: confirm filtering + rapid toggles no longer crash (WPF-UI animations).
- Data Matrix: verify scope-based column culling + JSONL export formatting across modes.
- Data Matrix: validate Filter Builder behavior (join enablement, per-rule case/trim, regex/numeric/date validation, layout spacing).
- Quick Colour: verify Profiles tab apply no longer asks for category/property when a saved profile is selected.
- Smart Grouping: preview uses Data Scraper cache and does not apply grouping scope (generation does apply scope).
- Quick Colour: scope filtering + profile scope encoding need verification across selection set/tree/property filter.
- Quick Colour: verify palette ordering/labels, shade expansion, and Default stable colors behavior.
- Quick Colour: some hierarchy/quick-colour applies still miss low-count types when Max types is small; decide default for Max types and add UI hint.
- Viewpoints Generator: PropertyGroups mode is a placeholder (plan empty).

## Key files touched
- `MicroEng.Navisworks/TreeMapper/TreeMapperWindow.xaml`
- `MicroEng.Navisworks/TreeMapper/TreeMapperWindow.xaml.cs`
- `MicroEng.Navisworks/TreeMapper/TreeMapperModels.cs`
- `MicroEng.Navisworks/TreeMapper/TreeMapperViewModels.cs`
- `MicroEng.Navisworks/TreeMapper/TreeMapperNodeTypeToLabelConverter.cs`
- `MicroEng.Navisworks/TreeMapper/TreeMapperNodeTypeToSymbolConverter.cs`
- `MicroEng.SelectionTreeCom/TreeMapperSelectionTreeCom.cs`
- `MicroEng.Navisworks/NavisworksSelectionSetUtils.cs`
- `MicroEng.Navisworks/DataScraperService.cs`
- `MicroEng.Navisworks/DataScraperModels.cs`
- `MicroEng.Navisworks/DataMatrixModels.cs`
- `MicroEng.Navisworks/DataMatrixPresetManager.cs`
- `MicroEng.Navisworks/DataMatrixRowBuilder.cs`
- `MicroEng.Navisworks/DataMatrixExporter.cs`
- `MicroEng.Navisworks/DataMatrixControl.xaml`
- `MicroEng.Navisworks/DataMatrixControl.xaml.cs`
- `MicroEng.Navisworks/DataMatrixColumnBuilderWindow.xaml`
- `MicroEng.Navisworks/DataMatrixColumnBuilderWindow.xaml.cs`
- `MicroEng.Navisworks/DataMatrixFilterBuilderWindow.xaml`
- `MicroEng.Navisworks/DataMatrixFilterBuilderWindow.xaml.cs`
- `MicroEng.Navisworks/ViewpointsGenerator/ViewpointsGeneratorModels.cs`
- `MicroEng.Navisworks/ViewpointsGenerator/ViewpointsGeneratorNavisworksService.cs`
- `MicroEng.Navisworks/ViewpointsGenerator/ViewpointsGeneratorControl.xaml`
- `MicroEng.Navisworks/ViewpointsGenerator/ViewpointsGeneratorControl.xaml.cs`
- `MicroEng.Navisworks/ViewpointsGenerator/ViewpointsGeneratorWindow.xaml`
- `MicroEng.Navisworks/ViewpointsGenerator/ViewpointsGeneratorWindow.xaml.cs`
- `MicroEng.Navisworks/ViewpointsGenerator/ViewpointsGeneratorPlugins.cs`
- `MicroEng.Navisworks/QuickColour/NotifyBase.cs`
- `MicroEng.Navisworks/QuickColour/QuickColourModels.cs`
- `MicroEng.Navisworks/QuickColour/QuickColourPalette.cs`
- `MicroEng.Navisworks/Colour/ColourPaletteGenerator.cs`
- `MicroEng.Navisworks/QuickColour/QuickColourHierarchyModels.cs`
- `MicroEng.Navisworks/QuickColour/QuickColourHueGroupModels.cs`
- `MicroEng.Navisworks/QuickColour/HueGroupAutoAssignPreviewModels.cs`
- `MicroEng.Navisworks/QuickColour/DisciplineMapModels.cs`
- `MicroEng.Navisworks/QuickColour/DisciplineMapMatcherEngine.cs`
- `MicroEng.Navisworks/QuickColour/QuickColourHueGroupAutoAssignService.cs`
- `MicroEng.Navisworks/QuickColour/QuickColourNavisworksService.cs`
- `MicroEng.Navisworks/QuickColour/QuickColourValueBuilder.cs`
- `MicroEng.Navisworks/QuickColour/QuickColourLegendExporter.cs`
- `MicroEng.Navisworks/QuickColour/Profiles/MicroEngColourProfileModels.cs`
- `MicroEng.Navisworks/QuickColour/Profiles/MicroEngColourProfileStore.cs`
- `MicroEng.Navisworks/QuickColour/Profiles/QuickColourProfilesPage.xaml`
- `MicroEng.Navisworks/QuickColour/Profiles/QuickColourProfilesPage.xaml.cs`
- `MicroEng.Navisworks/QuickColour/Profiles/ColorHexToBrushConverter.cs`
- `MicroEng.Navisworks/QuickColour/QuickColourControl.xaml`
- `MicroEng.Navisworks/QuickColour/QuickColourControl.xaml.cs`
- `MicroEng.Navisworks/QuickColour/QuickColourWindow.xaml`
- `MicroEng.Navisworks/QuickColour/QuickColourWindow.xaml.cs`
- `MicroEng.Navisworks/QuickColour/QuickColourPlugins.cs`
- `MicroEng.Navisworks/MicroEngPlugins.cs`
- `MicroEng.Navisworks/MicroEngPanelControl.xaml`
- `MicroEng.Navisworks/MicroEngPanelControl.xaml.cs`
- `MicroEng.Navisworks/MicroEng.Navisworks.addin`
- `MicroEng.Navisworks/SmartSets/SmartSetGeneratorControl.xaml`
- `MicroEng.Navisworks/SmartSets/SmartSetGeneratorControl.xaml.cs`
- `MicroEng.Navisworks/SmartSets/SmartSetGeneratorQuickBuilderPage.xaml`
- `MicroEng.Navisworks/SmartSets/SmartSetGeneratorSmartGroupingPage.xaml`
- `MicroEng.Navisworks/SmartSets/SmartSetGeneratorPlugins.cs`
- `MicroEng.Navisworks/SmartSets/SmartSetGeneratorNavisworksService.cs`
- `MicroEng.Navisworks/SmartSets/SmartSetModels.cs`
- `MicroEng.Navisworks/SmartSets/SmartSetFastPreviewService.cs`
- `MicroEng.Navisworks/SmartSets/SmartSetRecipeStore.cs`
- `MicroEng.Navisworks/SmartSets/SmartSetGroupingEngine.cs`
- `MicroEng.Navisworks/SmartSets/SmartSetInferenceEngine.cs`
- `MicroEng.Navisworks/SmartSets/SmartSetPackDefinitions.cs`
- `MicroEng.Navisworks/SmartSets/DataScraperSessionAdapter.cs`
- `MicroEng.Navisworks/SmartSets/PropertyPickerWindow.xaml`
- `MicroEng.Navisworks/SmartSets/PropertyPickerWindow.xaml.cs`
- `MicroEng.Navisworks/SmartSets/SmartSetScopePickerWindow.xaml`
- `MicroEng.Navisworks/SmartSets/SmartSetScopePickerWindow.xaml.cs`
- `MicroEng.Navisworks/DataScraperWindow.xaml`
- `MicroEng.Navisworks/DataScraperWindow.xaml.cs`
- `MicroEng.Navisworks/SpaceMapperModels.cs`
- `MicroEng.Navisworks/SpaceMapperEngines.cs`
- `MicroEng.Navisworks/SpaceMapperGeometry.cs`
- `MicroEng.Navisworks/SpaceMapperPreflightService.cs`
- `MicroEng.Navisworks/SpaceMapperService.cs`
- `MicroEng.Navisworks/SpaceMapperComparisonRunner.cs`
- `MicroEng.Navisworks/SpaceMapperControl.xaml.cs`
- `MicroEng.Navisworks/SpaceMapperStepProcessingPage.xaml`
- `MicroEng.Navisworks/SpaceMapperStepProcessingPage.xaml.cs`
- `MicroEng.Navisworks/NavisworksDockPaneManager.cs`
- `MicroEng.Navisworks/MicroEngStorageSettings.cs`
- `MicroEng.Navisworks/MicroEng.Navisworks.csproj`
- `MicroEng.Navisworks/Gpu/D3D11PointInMeshGpu.cs`
- `MicroEng.Navisworks/SpaceMapper/Gpu/CudaPointInMeshGpu.cs`
- `MicroEng.Navisworks/SpaceMapper/Gpu/CudaBvhPointInMeshGpu.cs`
- `MicroEng.Navisworks/SpaceMapperRunReportWriter.cs`
- `Native/MicroEng.CudaPointInMesh/microeng_cuda_point_in_mesh.cu`
- `Native/MicroEng.CudaBvhPointInMesh/CMakeLists.txt`
- `Native/MicroEng.CudaBvhPointInMesh/microeng_cuda_bvh_point_in_mesh.cu`

## Verification ideas
- Column Builder: set `MICROENG_COLUMNBUILDER_TRACE=1` to log perf; logs are at `%LOCALAPPDATA%\MicroEng.Navisworks\NavisErrors\MicroEng.log`.
- Column Builder: filter-by-selection should do **no work** and show nothing when no items are selected.
- Column Builder: filter-by-selection should scan selected items quickly without freezing; if it stalls, capture VS “Break All” call stack.
- Data Matrix: default view opens with no columns; loading a preset restores columns/scope.
- Data Matrix: Column Builder search filters the tree; tri-state parent reflects child state; right list toggles add/remove without crash.
- Data Matrix: scope culling hides properties not present in selected scope; selection/search set scopes resolve items.
- Data Matrix: JSONL export works in Item docs/Raw rows modes, optional .jsonl.gz.
- Crash reporting: unhandled exception writes a file under CrashReports with log tail + pane summary.
- Data Scraper: status flips to “Processing - Please Wait” in orange while running; returns to Ready after.
- Data Scraper: Export Now disabled until Output path set; Export Now keeps scroll position.
- Data Scraper: ArrowSync icons show during Run/Export; snackbars appear on success/error.
- Data Scraper: History `Raw` column shows counts quickly without loading full raw rows for every run.
- Data Scraper: raw rows load only after opening the `Raw Data` tab; `Properties` tab remains responsive on large history.
- Data Scraper: deleting a run with `X` removes history and deletes matching `DataScraperRaw\\<sessionId>.json`.
- Data Scraper: first load after legacy migration may take longer; subsequent opens should be significantly faster.
- Viewpoints Generator: Refresh loads selection/search sets; Preview shows plan rows + counts.
- Viewpoints Generator: Generate writes Saved Viewpoints under `MicroEng/Viewpoints Generator` using chosen direction/projection.
- Quick Colour: Profiles tab loads profile list, preview shows color swatches, Apply Temp/Permanent uses profile scope.
- Quick Colour: Scope options match Smart Sets and Load Values respects Current selection / Saved selection set / Tree selection / Property filter.
- Quick Colour: Custom Hue palette keeps Deep hue assignment but shifts lightness to the picked base color.
- Quick Colour: Palette list shows Default then Custom, fixed palettes expand when needed, and Default stable colors remain deterministic.
- Quick Colour: Shade Groups toggle keeps base colours and shades types (off = types match base).
- Quick Colour: Values table swatches open the color picker and update row hex.
- Quick Colour: Load hierarchy, edit base hex, Regenerate Type Shades updates type colors.
- Quick Colour: Single hue mode (lock hue) produces category bands + type subshades.
- Quick Colour: Apply Temporary/Permanent colors + optional Search/Snapshot set output.
- Smart Sets: Condition/category/property dropdowns are clickable in Quick Builder rules grid.
- Smart Sets: Include Blanks toggles empty-set generation in Quick Builder and Smart Grouping.
- Smart Sets: Scope picker tree uses WPF-UI checkbox visuals; indeterminate only shows for ancestors.
- Data Scraper: Entire Model is first and default in the scope list.
- Smart Sets: Property picker opens and populates Category/Property + sample values.
- Smart Sets: Fast preview returns counts from Data Scraper cache; live preview still works.
- Smart Sets: Save/Load recipes create JSON in ProgramData.
- Panel: card-style tool buttons show icons and toggle the correct panes.
- Step 3: Benchmark button creates a report and shows a summary in the "Latest benchmark" card.
- Fast preset: origin-point containment works; partial tagging only appears when partial options are enabled.
- Advanced Performance: Target-major is disabled or falls back when partials are enabled; Auto picks Target-major when partials are off and targets >> zones.
- Skip unchanged: repeat runs reduce writes and increment SkippedUnchanged in stats/report.
- Pack outputs: writes a single packed property instead of per-mapping properties.
- Close panes: Selection Tree, Find Items, and Properties close during run and restore after.
- Close pane delay: with Close panes on and delay=6, progress window appears only after the delay; live/run report shows "Close pane delay (sec)".
- Target geometry samples (GPU): option appears in containment calculation dropdown and report shows mode=TargetGeometryGpu.
- Zone offsets: when enabled, offsets change containment results; when disabled, offsets are ignored.
- Offset-only pass: baseline vs offset matches are tagged and "Zone Offset Match" writes Core/OffsetOnly (sequenced if needed).
- Run report includes "GPU Zone Eligibility" table with per-zone skip reason and thresholds.
- GPU backend shows "CUDA BVH" when the BVH DLL is present; otherwise "D3D11 batched" fallback.
- Variation check report includes "Normal / Zone-major (CPU)" baseline and "Normal / Zone-major (GPU)" comparison.

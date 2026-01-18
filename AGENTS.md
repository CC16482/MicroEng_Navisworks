# MicroEng Handoff

## Status
- Last build: not rerun after Smart Set Generator/UI updates (rerun `dotnet build` to confirm).
- Recommended: close Navisworks or use `dotnet build /p:DeployToProgramData=false` if ProgramData deploy is locked.
- Working tree was clean at last check (no uncommitted changes).
- Benchmark reports are saved under `C:\ProgramData\Autodesk\Navisworks Manage 2025\Plugins\MicroEng.Navisworks\Reports\`.
- Smart Set recipes saved under `C:\ProgramData\Autodesk\Navisworks Manage 2025\Plugins\MicroEng.Navisworks\SmartSets\Recipes\`.

## Recent changes
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

## Open issues
- Smart Set Generator: Condition dropdown in the rules grid is not rendering in the UI. Likely needs a template-based ComboBox in the DataGrid cell or Wpf.Ui DataGrid style adjustment.

## Key files touched
- `MicroEng.Navisworks/MicroEngPanelControl.xaml`
- `MicroEng.Navisworks/MicroEngPanelControl.xaml.cs`
- `MicroEng.Navisworks/SmartSets/SmartSetGeneratorControl.xaml`
- `MicroEng.Navisworks/SmartSets/SmartSetGeneratorControl.xaml.cs`
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
- `MicroEng.Navisworks/MicroEng.Navisworks.csproj`
- `MicroEng.Navisworks/Gpu/D3D11PointInMeshGpu.cs`
- `MicroEng.Navisworks/SpaceMapper/Gpu/CudaPointInMeshGpu.cs`
- `MicroEng.Navisworks/SpaceMapper/Gpu/CudaBvhPointInMeshGpu.cs`
- `MicroEng.Navisworks/SpaceMapperRunReportWriter.cs`
- `Native/MicroEng.CudaPointInMesh/microeng_cuda_point_in_mesh.cu`
- `Native/MicroEng.CudaBvhPointInMesh/CMakeLists.txt`
- `Native/MicroEng.CudaBvhPointInMesh/microeng_cuda_bvh_point_in_mesh.cu`

## Verification ideas
- Smart Sets: Condition dropdown appears and is editable in Quick Builder rules grid.
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

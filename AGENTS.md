# Space Mapper Handoff

## Status
- Last build: `dotnet build /p:DeployToProgramData=false` (succeeded)
- Working tree includes build outputs and generated files.
- Benchmark reports are saved under `C:\ProgramData\Autodesk\Navisworks Manage 2025\Plugins\MicroEng.Navisworks\Reports\`.

## Recent changes
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
- Geometry planes: only create triangle planes when vertices.Count % 3 == 0 (prevents bogus planes from bbox vertices).

## Key files touched
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

## Verification ideas
- Step 3: Benchmark button creates a report and shows a summary in the "Latest benchmark" card.
- Fast preset: origin-point containment works; partial tagging only appears when partial options are enabled.
- Advanced Performance: Target-major is disabled or falls back when partials are enabled; Auto picks Target-major when partials are off and targets >> zones.
- Skip unchanged: repeat runs reduce writes and increment SkippedUnchanged in stats/report.
- Pack outputs: writes a single packed property instead of per-mapping properties.
- Close panes: Selection Tree, Find Items, and Properties close during run and restore after.

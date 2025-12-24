# Space Mapper Handoff

## Status
- Last build: `dotnet build /p:DeployToProgramData=false` (succeeded)
- Working tree includes build outputs and generated files.

## Recent changes
- Step 2 Zones/Targets: combined Selection/Search set dropdown; list populated via managed API (`Document.SelectionSets.RootItem`) with COM fallback; selecting a set auto-sets Zone Source or Target Definition and fills Set/Search name.
- Run Space Mapper button moved to the top banner; Refresh Sets button added to the banner.
- Partial tagging now writes a separate "Zone Behaviour" property (configurable in Processing settings); removed Partial Flag mapping UI.
- Property writing now reuses the existing ME_SpaceInfo category to avoid duplicate property tabs.
- Preflight/estimate, performance presets, and index granularity wiring already in place; CPU engine reuses spatial grid where available.
- Step 1 profiles decoupled: Space Mapper templates are independent from Data Scraper profiles.

## Key files touched
- `MicroEng.Navisworks/SpaceMapperControl.xaml`
- `MicroEng.Navisworks/SpaceMapperControl.xaml.cs`
- `MicroEng.Navisworks/SpaceMapperStepZonesTargetsPage.xaml`
- `MicroEng.Navisworks/SpaceMapperStepZonesTargetsPage.xaml.cs`
- `MicroEng.Navisworks/SpaceMapperService.cs`
- `MicroEng.Navisworks/SpaceMapperStepProcessingPage.xaml`
- `MicroEng.Navisworks/SpaceMapperStepProcessingPage.xaml.cs`
- `MicroEng.Navisworks/SpaceMapperStepMappingPage.xaml`
- `MicroEng.Navisworks/SpaceMapperStepMappingPage.xaml.cs`
- `MicroEng.Navisworks/SpaceMapperModels.cs`
- `MicroEng.Navisworks/SpaceMapperTemplates.cs`

## Verification ideas
- Navisworks Step 2: click Refresh Sets and open the dropdown; confirm sets show as "Name (Selection)" or "Name (Search)" and choosing one updates Zone Source / Target Definition.
- Run Space Mapper button in the header works from any step.
- "Write Zone Behaviour property" writes a separate property instead of appending ", Partial" to the zone name.

# ProjectTree

This file is a generated map of the repository layout and how key parts connect.

## Scope and exclusions
- Included: all repo files and folders.
- Excluded (generated/dev-only): .git, .vs, bin, obj.

## Solution and projects
- MicroEng.Navisworks -> MicroEng.Navisworks\MicroEng.Navisworks.csproj
- MicroEng.DesignHost -> MicroEng.DesignHost\MicroEng.DesignHost.csproj

## Runtime entry points (Navisworks plugins)
- AddinManager (defined in ReferenceProjects/NavisAddinManager-dev/NavisAddinManager/Application/App.cs)
- AddinManagerManual (defined in ReferenceProjects/NavisAddinManager-dev/Test/TestAddInManagerManual.cs)
- APICallsCOMPlugin.APICallsCOMPlugin (defined in ReferenceProjects/NavisAddinManager-dev/Test/APICallsCOMPlugin.cs)
- AppendDataTool.AppendDataDockPane.MyCompany (defined in ReferenceProjects/RSS_Plugin/Plugin/AppendDataDockPane.cs)
- AppendDataTool.RibbonHandler.MyCompany (defined in ReferenceProjects/RSS_Plugin/Plugin/RibbonHandler.cs)
- AppInfo (defined in ReferenceProjects/NavisLookup-dev/NavisAppInfo/App.cs)
- EnableToolPluginExample (defined in ReferenceProjects/NavisAddinManager-dev/Test/ToolPluginTest.cs)
- HelloWorld (defined in ReferenceProjects/NavisAddinManager-dev/Test/TestAddInManagerManual.cs)
- InputAndRenderHandling.EnableInputPluginExample (defined in ReferenceProjects/NavisAddinManager-dev/Test/EnableInputPluginExample.cs)
- InputAndRenderHandling.EnableRenderPluginExample (defined in ReferenceProjects/NavisAddinManager-dev/Test/EnableInputPluginExample.cs)
- InputAndRenderHandling.InputPluginExample (defined in ReferenceProjects/NavisAddinManager-dev/Test/EnableInputPluginExample.cs)
- MicroEng.AppendData (defined in MicroEng.Navisworks/MicroEngPlugins.cs)
- MicroEng.DataMatrix.Command (defined in MicroEng.Navisworks/DataMatrixPlugins.cs)
- MicroEng.DataMatrix.DockPane (defined in MicroEng.Navisworks/DataMatrixPlugins.cs)
- MicroEng.DockPane (defined in MicroEng.Navisworks/MicroEngPlugins.cs)
- MicroEng.PanelCommand (defined in MicroEng.Navisworks/MicroEngPlugins.cs)
- MicroEng.QuickColour.Command (defined in MicroEng.Navisworks/QuickColour/QuickColourPlugins.cs)
- MicroEng.Sequence4D.Command (defined in MicroEng.Navisworks/Sequence4DPlugins.cs)
- MicroEng.Sequence4D.DockPane (defined in MicroEng.Navisworks/Sequence4DPlugins.cs)
- MicroEng.SmartSetGenerator.Command (defined in MicroEng.Navisworks/SmartSets/SmartSetGeneratorPlugins.cs)
- MicroEng.SmartSetGenerator.DockPane (defined in MicroEng.Navisworks/SmartSets/SmartSetGeneratorPlugins.cs)
- MicroEng.SpaceMapper.Command (defined in MicroEng.Navisworks/SpaceMapperPlugins.cs)
- MicroEng.SpaceMapper.DockPane (defined in MicroEng.Navisworks/SpaceMapperPlugins.cs)
- MicroEng.ViewpointsGenerator.Command (defined in MicroEng.Navisworks/ViewpointsGenerator/ViewpointsGeneratorPlugins.cs)
- MicroEng.ViewpointsGenerator.DockPane (defined in MicroEng.Navisworks/ViewpointsGenerator/ViewpointsGeneratorPlugins.cs)
- NETInteropTest (defined in ReferenceProjects/NavisAddinManager-dev/Test/ABasicPlugin.cs)
- RenderPluginTest (defined in ReferenceProjects/NavisAddinManager-dev/Test/RenderPluginTest.cs)
- SearchComparisonPlugIn.SearchComparisonPlugIn (defined in ReferenceProjects/NavisAddinManager-dev/Test/SearchComparisonPlugIn.cs)
- ToolPluginTest (defined in ReferenceProjects/NavisAddinManager-dev/Test/ToolPluginTest.cs)

## Key wiring
- MicroEng.Navisworks/MicroEng.Navisworks.addin -> loads MicroEng.Navisworks.dll from the plugin folder.
- MicroEng.Navisworks/MicroEng.Navisworks.csproj -> builds the plugin, references Navisworks APIs, and copies assets.
- MicroEng.Navisworks/MicroEngPlugins.cs -> shared actions + dock pane host; MicroEng panel uses MicroEngPanelControl.
- MicroEng.Navisworks/*Plugins.cs -> individual tool DockPane/Command registrations for Navisworks.
- Native/* -> CUDA native DLLs; copied to output via the csproj after build.

## Folder map
```text
|-- Logos
|   |-- microeng-logo.png
|   |-- microeng-logo2.png
|   |-- microeng-logo3.png
|   `-- microeng_logotray.png
|-- MicroEng.DesignHost
|   |-- App.xaml
|   |-- App.xaml.cs
|   |-- DesignWindow.xaml
|   |-- DesignWindow.xaml.cs
|   `-- MicroEng.DesignHost.csproj
|-- MicroEng.Navisworks
|   |-- Assets
|   |   `-- Theme
|   |       |-- MicroEngUiKit.xaml
|   |       `-- MicroEngWpfUiRoot.xaml
|   |-- Colour
|   |   `-- ColourPaletteGenerator.cs
|   |-- Gpu
|   |   `-- D3D11PointInMeshGpu.cs
|   |-- QuickColour
|   |   |-- Profiles
|   |   |   |-- ColorHexToBrushConverter.cs
|   |   |   |-- MicroEngColourProfileModels.cs
|   |   |   |-- MicroEngColourProfileStore.cs
|   |   |   |-- QuickColourProfilesPage.xaml
|   |   |   `-- QuickColourProfilesPage.xaml.cs
|   |   |-- DisciplineMapMatcherEngine.cs
|   |   |-- DisciplineMapModels.cs
|   |   |-- HueGroupAutoAssignPreviewModels.cs
|   |   |-- NotifyBase.cs
|   |   |-- QuickColourControl.xaml
|   |   |-- QuickColourControl.xaml.cs
|   |   |-- QuickColourHierarchyModels.cs
|   |   |-- QuickColourHueGroupAutoAssignService.cs
|   |   |-- QuickColourHueGroupModels.cs
|   |   |-- QuickColourLegendExporter.cs
|   |   |-- QuickColourModels.cs
|   |   |-- QuickColourNavisworksService.cs
|   |   |-- QuickColourPalette.cs
|   |   |-- QuickColourPlugins.cs
|   |   |-- QuickColourValueBuilder.cs
|   |   |-- QuickColourWindow.xaml
|   |   `-- QuickColourWindow.xaml.cs
|   |-- SmartSets
|   |   |-- DataScraperSessionAdapter.cs
|   |   |-- PropertyPickerWindow.xaml
|   |   |-- PropertyPickerWindow.xaml.cs
|   |   |-- SmartSetFastPreviewService.cs
|   |   |-- SmartSetGeneratorControl.xaml
|   |   |-- SmartSetGeneratorControl.xaml.cs
|   |   |-- SmartSetGeneratorFromSelectionPage.xaml
|   |   |-- SmartSetGeneratorFromSelectionPage.xaml.cs
|   |   |-- SmartSetGeneratorNavisworksService.cs
|   |   |-- SmartSetGeneratorPacksPage.xaml
|   |   |-- SmartSetGeneratorPacksPage.xaml.cs
|   |   |-- SmartSetGeneratorPlugins.cs
|   |   |-- SmartSetGeneratorQuickBuilderPage.xaml
|   |   |-- SmartSetGeneratorQuickBuilderPage.xaml.cs
|   |   |-- SmartSetGeneratorSmartGroupingPage.xaml
|   |   |-- SmartSetGeneratorSmartGroupingPage.xaml.cs
|   |   |-- SmartSetGeneratorWindow.xaml
|   |   |-- SmartSetGeneratorWindow.xaml.cs
|   |   |-- SmartSetGroupingEngine.cs
|   |   |-- SmartSetInferenceEngine.cs
|   |   |-- SmartSetModels.cs
|   |   |-- SmartSetPackDefinitions.cs
|   |   |-- SmartSetRecipeStore.cs
|   |   |-- SmartSetScopePickerWindow.xaml
|   |   `-- SmartSetScopePickerWindow.xaml.cs
|   |-- SpaceMapper
|   |   |-- Estimation
|   |   |   |-- CalibrationStore.cs
|   |   |   |-- PreflightModels.cs
|   |   |   |-- RuntimeEstimator.cs
|   |   |   `-- SpaceMapperPresetLogic.cs
|   |   |-- Geometry
|   |   |   |-- Aabb.cs
|   |   |   |-- SpatialGridSizing.cs
|   |   |   `-- SpatialHashGrid.cs
|   |   |-- Gpu
|   |   |   |-- CudaBvhPointInMeshGpu.cs
|   |   |   |-- CudaPointInMeshGpu.cs
|   |   |   |-- GpuTypes.cs
|   |   |   `-- IPointInMeshGpuBackend.cs
|   |   `-- Util
|   |       `-- AsyncDebouncer.cs
|   |-- ViewpointsGenerator
|   |   |-- ViewpointsGeneratorControl.xaml
|   |   |-- ViewpointsGeneratorControl.xaml.cs
|   |   |-- ViewpointsGeneratorModels.cs
|   |   |-- ViewpointsGeneratorNavisworksService.cs
|   |   |-- ViewpointsGeneratorPlugins.cs
|   |   |-- ViewpointsGeneratorWindow.xaml
|   |   `-- ViewpointsGeneratorWindow.xaml.cs
|   |-- AppendIntegrateDialog.xaml
|   |-- AppendIntegrateDialog.xaml.cs
|   |-- AppendIntegrateExecutor.cs
|   |-- AppendIntegrateModels.cs
|   |-- Class1.cs
|   |-- DataMatrixColumnBuilderWindow.xaml
|   |-- DataMatrixColumnBuilderWindow.xaml.cs
|   |-- DataMatrixControl.xaml
|   |-- DataMatrixControl.xaml.cs
|   |-- DataMatrixExporter.cs
|   |-- DataMatrixFilterBuilderWindow.xaml
|   |-- DataMatrixFilterBuilderWindow.xaml.cs
|   |-- DataMatrixModels.cs
|   |-- DataMatrixPlugins.cs
|   |-- DataMatrixPresetManager.cs
|   |-- DataMatrixRowBuilder.cs
|   |-- DataScraperModels.cs
|   |-- DataScraperRunProgressState.cs
|   |-- DataScraperRunProgressWindow.xaml
|   |-- DataScraperRunProgressWindow.xaml.cs
|   |-- DataScraperService.cs
|   |-- DataScraperWindow.xaml
|   |-- DataScraperWindow.xaml.cs
|   |-- HoverFlyoutController.cs
|   |-- MicroEng.Navisworks.addin
|   |-- MicroEng.Navisworks.csproj
|   |-- MicroEng.Navisworks.csproj.user
|   |-- MicroEngPanelControl.xaml
|   |-- MicroEngPanelControl.xaml.cs
|   |-- MicroEngPlugins.cs
|   |-- MicroEngResourceUris.cs
|   |-- MicroEngSettingsWindow.xaml
|   |-- MicroEngSettingsWindow.xaml.cs
|   |-- MicroEngWindowPositioning.cs
|   |-- MicroEngWpfUiTheme.cs
|   |-- NavisworksDockPaneManager.cs
|   |-- NavisworksSelectionSetUtils.cs
|   |-- PropertyPickerDialog.cs
|   |-- Sequence4DControl.xaml
|   |-- Sequence4DControl.xaml.cs
|   |-- Sequence4DGenerator.cs
|   |-- Sequence4DModelItemUtils.cs
|   |-- Sequence4DModels.cs
|   |-- Sequence4DPlugins.cs
|   |-- SpaceMapperBoundsResolver.cs
|   |-- SpaceMapperComparisonRunner.cs
|   |-- SpaceMapperControl.xaml
|   |-- SpaceMapperControl.xaml.cs
|   |-- SpaceMapperEngines.cs
|   |-- SpaceMapperGeometry.cs
|   |-- SpaceMapperModels.cs
|   |-- SpaceMapperPlugins.cs
|   |-- SpaceMapperPreflightService.cs
|   |-- SpaceMapperRunProgressHost.cs
|   |-- SpaceMapperRunProgressState.cs
|   |-- SpaceMapperRunProgressWindow.xaml
|   |-- SpaceMapperRunProgressWindow.xaml.cs
|   |-- SpaceMapperRunReportWriter.cs
|   |-- SpaceMapperService.cs
|   |-- SpaceMapperStepMappingPage.xaml
|   |-- SpaceMapperStepMappingPage.xaml.cs
|   |-- SpaceMapperStepProcessingPage.xaml
|   |-- SpaceMapperStepProcessingPage.xaml.cs
|   |-- SpaceMapperStepResultsPage.xaml
|   |-- SpaceMapperStepResultsPage.xaml.cs
|   |-- SpaceMapperStepSetupPage.xaml
|   |-- SpaceMapperStepSetupPage.xaml.cs
|   |-- SpaceMapperStepZonesTargetsPage.xaml
|   |-- SpaceMapperStepZonesTargetsPage.xaml.cs
|   `-- SpaceMapperTemplates.cs
|-- Native
|   |-- MicroEng.CudaBvhPointInMesh
|   |   |-- CMakeLists.txt
|   |   `-- microeng_cuda_bvh_point_in_mesh.cu
|   `-- MicroEng.CudaPointInMesh
|       |-- CMakeLists.txt
|       `-- microeng_cuda_point_in_mesh.cu
|-- NavisErrors
|   |-- dmpuserinfo.xml
|   `-- NavisWorksErrorReport.dmp
|-- ReferenceDocuments
|   |-- Updates
|   |   |-- Codex_Instructions.txt
|   |   |-- Codex_Instructions2.txt
|   |   |-- Codex_Instructions3.txt
|   |   |-- Codex_Instructions4.txt
|   |   |-- DataMatrixColumnBuilderWindow.refined.xaml
|   |   |-- DataMatrixColumnBuilderWindow.refined.xaml.cs
|   |   |-- DataMatrixControl.updated.xaml
|   |   |-- DataMatrixControl.updated.xaml.cs
|   |   `-- DataMatrixModels.updated.cs
|   |-- Append_Data_Instructions.txt
|   |-- Codex_Conversation_260104.txt
|   |-- Codex_Conversation_260105.txt
|   |-- Codex_Instructions_01.txt
|   |-- Codex_Instructions_02.txt
|   |-- Codex_Instructions_03.txt
|   |-- Codex_Instructions_04.txt
|   |-- Codex_Instructions_05.txt
|   |-- Codex_Instructions_06.txt
|   |-- Codex_Instructions_07.txt
|   |-- Codex_Instructions_08.txt
|   |-- Codex_Instructions_09.txt
|   |-- Codex_Instructions_10.txt
|   |-- Data_Matrix_Instructions.txt
|   |-- iConstruct_Pro_Append_Data_&_Integrator.docx
|   |-- iConstruct_Pro_Zone_Tools.docx
|   |-- Sequence4D_UpdatedAdditions_01.txt
|   |-- Smart Sets_01.txt
|   |-- Smart Sets_02.txt
|   |-- Space_Mapper-NewFeatures-ProcessingStep_Instructions.txt
|   |-- Space_Mapper-UpdatedAdditions_01.txt
|   |-- Space_Mapper-UpdatedAdditions_02.txt
|   |-- Space_Mapper-UpdatedAdditions_03.txt
|   |-- Space_Mapper-UpdatedAdditions_04.txt
|   |-- Space_Mapper-UpdatedAdditions_05.txt
|   |-- Space_Mapper-UpdatedAdditions_06.txt
|   |-- Space_Mapper-UpdatedAdditions_07.txt
|   |-- Space_Mapper-UpdatedAdditions_08.txt
|   |-- Space_Mapper-UpdatedAdditions_09.txt
|   |-- Space_Mapper-UpdatedAdditions_10.txt
|   |-- Space_Mapper-UpdatedAdditions_11.txt
|   |-- Space_Mapper-UpdatedAdditions_12.txt
|   |-- Space_Mapper-UpdatedAdditions_13.txt
|   |-- Space_Mapper-UpdatedAdditions_14.txt
|   |-- Space_Mapper-UpdatedAdditions_15.txt
|   |-- Space_Mapper-UpdatedAdditions_16.txt
|   |-- Space_Mapper-UpdatedAdditions_17.txt
|   |-- Space_Mapper-UpdatedAdditions_18.txt
|   |-- Space_Mapper-UpdatedAdditions_19.txt
|   |-- Space_Mapper-UpdatedAdditions_20.txt
|   |-- Space_Mapper-UpdatedAdditions_21.txt
|   |-- Space_Mapper-UpdatedAdditions_22.txt
|   |-- Space_Mapper-UpdatedAdditions_23.txt
|   |-- Space_Mapper-UpdatedAdditions_24.txt
|   |-- Space_Mapper-UpdatedAdditions_25.txt
|   |-- Space_Mapper-UpdatedAdditions_26.txt
|   |-- Space_Mapper-UpdatedAdditions_27.txt
|   |-- Space_Mapper_CPU_Normal_Processing_Instructions.txt
|   |-- Space_Mapper_Instructions.txt
|   |-- SpaceMapper_Stats.csv
|   `-- WinDbg_Export.txt
|-- ReferenceMedia
|   `-- SpaceMapper
|       |-- Zone_3D_Offset.gif
|       |-- Zone_3D_Offset_720p.mp4
|       |-- Zone_Bottom.gif
|       |-- Zone_Bottom_720p.mp4
|       |-- Zone_Sides.gif
|       |-- Zone_Sides_720p.mp4
|       |-- Zone_Top.gif
|       `-- Zone_Top_720p.mp4
|-- ReferenceProjects
|   |-- NavisAddinManager-dev
|   |   |-- .github
|   |   |   |-- ISSUE_TEMPLATE
|   |   |   |   |-- bug_report.md
|   |   |   |   `-- feature_request.md
|   |   |   |-- workflows
|   |   |   |   `-- Workflow.yml
|   |   |   `-- PULL_REQUEST_TEMPLATE.md
|   |   |-- .nuke
|   |   |   |-- build.cmd
|   |   |   |-- build.ps1
|   |   |   |-- build.schema.json
|   |   |   `-- parameters.json
|   |   |-- .run
|   |   |   |-- Nuke Clean.run.xml
|   |   |   |-- Nuke Plan.run.xml
|   |   |   |-- Nuke.run.xml
|   |   |   `-- Revit 2022 ENU.run.xml
|   |   |-- build
|   |   |   |-- .editorconfig
|   |   |   |-- Build.Clean.cs
|   |   |   |-- Build.Compile.cs
|   |   |   |-- Build.cs
|   |   |   |-- Build.csproj
|   |   |   |-- Build.csproj.DotSettings
|   |   |   |-- Build.GitHubRelease.cs
|   |   |   |-- Build.Installer.cs
|   |   |   |-- Build.Properties.cs
|   |   |   `-- BuilderExtensions.cs
|   |   |-- Installer
|   |   |   |-- Resources
|   |   |   |   `-- Icons
|   |   |   |       |-- BackgroundImage.png
|   |   |   |       |-- BannerImage.png
|   |   |   |       `-- ShellIcon.ico
|   |   |   |-- Installer.cs
|   |   |   `-- Installer.csproj
|   |   |-- NavisAddinManager
|   |   |   |-- Application
|   |   |   |   `-- App.cs
|   |   |   |-- Command
|   |   |   |   |-- AddinManagerBase.cs
|   |   |   |   |-- AddInManagerCommand.cs
|   |   |   |   `-- DocpanelCommand.cs
|   |   |   |-- DockPanel
|   |   |   |   |-- AddinManagerPane.cs
|   |   |   |   `-- DockPaneBase.cs
|   |   |   |-- en-US
|   |   |   |   `-- AddinManagerRibbon.xaml
|   |   |   |-- Model
|   |   |   |   |-- Addin.cs
|   |   |   |   |-- AddinItem.cs
|   |   |   |   |-- AddinItemComparer.cs
|   |   |   |   |-- Addins.cs
|   |   |   |   |-- AddinType.cs
|   |   |   |   |-- AssemLoader.cs
|   |   |   |   |-- BitmapSourceConverter.cs
|   |   |   |   |-- CodeListener.cs
|   |   |   |   |-- DefaultSetting.cs
|   |   |   |   |-- FileUtils.cs
|   |   |   |   |-- FolderTooBigDialog.cs
|   |   |   |   |-- IAddinNode.cs
|   |   |   |   |-- IEConverter.cs
|   |   |   |   |-- IniFile.cs
|   |   |   |   |-- LogMessageString.cs
|   |   |   |   |-- ManifestFile.cs
|   |   |   |   |-- NavisAddin.cs
|   |   |   |   |-- ProcessManager.cs
|   |   |   |   |-- StaticUtil.cs
|   |   |   |   |-- ViewModelBase.cs
|   |   |   |   `-- VisibilityMode.cs
|   |   |   |-- Properties
|   |   |   |   |-- launchSettings.json
|   |   |   |   |-- Settings.Designer.cs
|   |   |   |   `-- Settings.settings
|   |   |   |-- Resources
|   |   |   |   |-- dev.ico
|   |   |   |   |-- dev.png
|   |   |   |   |-- dev16x16.png
|   |   |   |   |-- dev32x32.png
|   |   |   |   |-- folder.png
|   |   |   |   |-- lab16x16.png
|   |   |   |   `-- lab32x32.png
|   |   |   |-- View
|   |   |   |   |-- Control
|   |   |   |   |   |-- ExtendedTreeView.cs
|   |   |   |   |   |-- LogControl.xaml
|   |   |   |   |   |-- LogControl.xaml.cs
|   |   |   |   |   |-- MouseDoubleClick.cs
|   |   |   |   |   |-- RelayCommand.cs
|   |   |   |   |   `-- VirtualToggleButton.cs
|   |   |   |   |-- AssemblyLoader.xaml
|   |   |   |   |-- AssemblyLoader.xaml.cs
|   |   |   |   |-- FrmAddInManager.xaml
|   |   |   |   `-- FrmAddInManager.xaml.cs
|   |   |   |-- ViewModel
|   |   |   |   |-- AddinManager.cs
|   |   |   |   |-- AddInManagerViewModel.cs
|   |   |   |   |-- AddinModel.cs
|   |   |   |   |-- AddinsApplication.cs
|   |   |   |   |-- AddinsCommand.cs
|   |   |   |   `-- LogControlViewModel.cs
|   |   |   |-- NavisAddinManager.csproj
|   |   |   |-- PackageContents.xml
|   |   |   |-- postbuild.ps1
|   |   |   |-- Resource.Designer.cs
|   |   |   `-- Resource.resx
|   |   |-- pic
|   |   |   |-- 7aF7wDel5L.gif
|   |   |   |-- Addin.png
|   |   |   |-- AddinManager.png
|   |   |   |-- jetbrains.png
|   |   |   |-- MouseWheel.gif
|   |   |   |-- NavisAddinManager.png
|   |   |   |-- Trace-Debug.png
|   |   |   `-- WhyAddinManager.png
|   |   |-- Test
|   |   |   |-- FileSample
|   |   |   |   |-- ClashTest.nwd
|   |   |   |   `-- RevitNative.rvt
|   |   |   |-- Helpers
|   |   |   |   |-- INavisCommand.cs
|   |   |   |   |-- Logger.cs
|   |   |   |   `-- VariantDataUtils.cs
|   |   |   |-- ABasicPlugin.cs
|   |   |   |-- AccessClashReport.cs
|   |   |   |-- AcessSelectionSet.cs
|   |   |   |-- AddCustomProperties.cs
|   |   |   |-- APICallsCOMPlugin.cs
|   |   |   |-- CreateQuickProperties.cs
|   |   |   |-- DebugTrace.cs
|   |   |   |-- EnableInputPluginExample.cs
|   |   |   |-- FocusItem.cs
|   |   |   |-- GetClashTest.cs
|   |   |   |-- GetFileUnit.cs
|   |   |   |-- GetProperties.cs
|   |   |   |-- GetPropertiesFromClashTest.cs
|   |   |   |-- Highlight.cs
|   |   |   |-- RenderPluginTest.cs
|   |   |   |-- SaveViewPort.cs
|   |   |   |-- SearchComparisonPlugIn.cs
|   |   |   |-- SetHidden.cs
|   |   |   |-- Test.csproj
|   |   |   |-- TestAddInManagerManual.cs
|   |   |   |-- ToolPluginTest.cs
|   |   |   |-- Transformbox.cs
|   |   |   |-- traverseComponents.cs
|   |   |   |-- ViewpointTest.cs
|   |   |   `-- ZoomToCurrentSelection.cs
|   |   |-- .gitignore
|   |   |-- AddInManager.sln
|   |   |-- CHANGELOG.md
|   |   |-- CODE_OF_CONDUCT.md
|   |   |-- CONTRIBUTING.md
|   |   |-- License.md
|   |   `-- Readme.MD
|   |-- NavisLookup-dev
|   |   |-- .github
|   |   |   |-- ISSUE_TEMPLATE
|   |   |   |   |-- bug_report.md
|   |   |   |   `-- feature_request.md
|   |   |   |-- workflows
|   |   |   |   `-- Workflow.yml
|   |   |   `-- PULL_REQUEST_TEMPLATE.md
|   |   |-- .nuke
|   |   |   |-- build.cmd
|   |   |   |-- build.ps1
|   |   |   |-- build.schema.json
|   |   |   `-- parameters.json
|   |   |-- build
|   |   |   |-- .editorconfig
|   |   |   |-- Build.Clean.cs
|   |   |   |-- Build.Compile.cs
|   |   |   |-- Build.cs
|   |   |   |-- Build.csproj
|   |   |   |-- Build.csproj.DotSettings
|   |   |   |-- Build.GitHubRelease.cs
|   |   |   |-- Build.Installer.cs
|   |   |   |-- Build.Properties.cs
|   |   |   `-- BuilderExtensions.cs
|   |   |-- Installer
|   |   |   |-- Resources
|   |   |   |   `-- Icons
|   |   |   |       |-- BackgroundImage.png
|   |   |   |       |-- BannerImage.png
|   |   |   |       `-- ShellIcon.ico
|   |   |   |-- Installer.cs
|   |   |   `-- Installer.csproj
|   |   |-- NavisAppInfo
|   |   |   |-- Command
|   |   |   |   |-- BaseCommand.cs
|   |   |   |   |-- EnumSnoopType.cs
|   |   |   |   |-- SnoopActiveSheet.cs
|   |   |   |   |-- SnoopActiveView.cs
|   |   |   |   |-- SnoopApplication.cs
|   |   |   |   |-- SnoopByElementId.cs
|   |   |   |   |-- SnoopClashTest.cs
|   |   |   |   |-- SnoopCurrentSelection.cs
|   |   |   |   |-- SnoopDocument.cs
|   |   |   |   |-- SnoopSearch.cs
|   |   |   |   `-- SnoopTest.cs
|   |   |   |-- en-US
|   |   |   |   `-- AppInfoRibbon.xaml
|   |   |   |-- Events
|   |   |   |   |-- EventDetails.cs
|   |   |   |   |-- EventDetailsArgs.cs
|   |   |   |   `-- EventHandlers.cs
|   |   |   |-- Model
|   |   |   |   |-- FormIcons.cs
|   |   |   |   |-- NodeInfo.cs
|   |   |   |   |-- NodeSearch.cs
|   |   |   |   `-- TypeExtensions.cs
|   |   |   |-- Properties
|   |   |   |   |-- launchSettings.json
|   |   |   |   |-- Resources.Designer.cs
|   |   |   |   `-- Resources.resx
|   |   |   |-- Resources
|   |   |   |   |-- app-16.png
|   |   |   |   |-- app-32.png
|   |   |   |   |-- conflict-16.png
|   |   |   |   |-- conflict-32.png
|   |   |   |   |-- cursor-16.png
|   |   |   |   |-- cursor-32.png
|   |   |   |   |-- document-16.png
|   |   |   |   |-- document-32.png
|   |   |   |   |-- pubclass.gif
|   |   |   |   |-- pubenum.gif
|   |   |   |   |-- pubevent.gif
|   |   |   |   |-- pubmethod.gif
|   |   |   |   |-- pubproperty.gif
|   |   |   |   |-- sheet-16.png
|   |   |   |   |-- sheet-32.png
|   |   |   |   |-- staticclass.gif
|   |   |   |   |-- staticmethod.GIF
|   |   |   |   |-- staticproperty.GIF
|   |   |   |   |-- test-16.png
|   |   |   |   |-- test-32.png
|   |   |   |   |-- view-16.png
|   |   |   |   `-- view-32.png
|   |   |   |-- View
|   |   |   |   |-- AppInfoControl.cs
|   |   |   |   |-- AppInfoControl.Designer.cs
|   |   |   |   |-- AppInfoControl.resx
|   |   |   |   |-- FrmAppInfo.xaml
|   |   |   |   |-- FrmAppInfo.xaml.cs
|   |   |   |   |-- MainWindow.xaml
|   |   |   |   |-- MainWindow.xaml.cs
|   |   |   |   |-- SearchByContains.xaml
|   |   |   |   `-- SearchByContains.xaml.cs
|   |   |   |-- ViewModel
|   |   |   |   `-- AppInfoViewModel.cs
|   |   |   |-- App.cs
|   |   |   |-- NavisLookup.csproj
|   |   |   |-- PackageContents.xml
|   |   |   `-- postbuild.ps1
|   |   |-- pic
|   |   |   |-- AppRibbon.png
|   |   |   |-- jetbrains.png
|   |   |   |-- NavisLookup.png
|   |   |   `-- SnoopCurrentSlection.png
|   |   |-- .gitignore
|   |   |-- CHANGELOG.md
|   |   |-- CODE_OF_CONDUCT.md
|   |   |-- CONTRIBUTING.md
|   |   |-- License.md
|   |   |-- NavisLookup.sln
|   |   `-- Readme.MD
|   |-- RSS_Plugin
|   |   |-- Images
|   |   |   |-- 1_RSS_Logo.png
|   |   |   |-- 2_RSS_White.png
|   |   |   `-- 3_NavTools.png
|   |   |-- Plugin
|   |   |   |-- AppendDataDockPane.cs
|   |   |   `-- RibbonHandler.cs
|   |   |-- UI
|   |   |   |-- AppendDataControl.xaml
|   |   |   `-- AppendDataControl.xaml.cs
|   |   `-- AppendDataTool.csproj
|   `-- wpfui-main
|       |-- .config
|       |   `-- dotnet-tools.json
|       |-- .devcontainer
|       |   |-- devcontainer.json
|       |   `-- post-create.sh
|       |-- .github
|       |   |-- assets
|       |   |   `-- microsoft-badge.png
|       |   |-- chatmodes
|       |   |   `-- documentation_contributor.chatmode.md
|       |   |-- ISSUE_TEMPLATE
|       |   |   |-- bug_report.yaml
|       |   |   |-- config.yml
|       |   |   `-- feature_request.yaml
|       |   |-- policies
|       |   |   |-- cla.yml
|       |   |   `-- platformcontext.yml
|       |   |-- workflows
|       |   |   |-- top-issues-dashboard.yml
|       |   |   |-- wpf-ui-cd-docs.yaml
|       |   |   |-- wpf-ui-cd-extension.yaml
|       |   |   |-- wpf-ui-cd-nuget.yaml
|       |   |   |-- wpf-ui-labeler.yml
|       |   |   |-- wpf-ui-lock.yml
|       |   |   `-- wpf-ui-pr-validator.yaml
|       |   |-- copilot-instructions.md
|       |   |-- dependabot.yml
|       |   |-- FUNDING.yml
|       |   |-- labeler.yml
|       |   |-- labels.yml
|       |   `-- pull_request_template.md
|       |-- branding
|       |   |-- geometric_splash.psd
|       |   |-- microsoft-fluent-resources.psd
|       |   |-- wpfui.ico
|       |   |-- wpfui.png
|       |   |-- wpfui.psd
|       |   `-- wpfui_full.png
|       |-- build
|       |   `-- nuget.png
|       |-- docs
|       |   |-- codesnippet
|       |   |   `-- Rtf
|       |   |       |-- Hyperlink
|       |   |       |   `-- RtfDocumentProcessor.cs
|       |   |       |-- RtfBuildStep.cs
|       |   |       `-- RtfDocumentProcessor.cs
|       |   |-- documentation
|       |   |   |-- .gitignore
|       |   |   |-- about-wpf.md
|       |   |   |-- accent.md
|       |   |   |-- extension.md
|       |   |   |-- fonticon.md
|       |   |   |-- gallery-editor.md
|       |   |   |-- gallery-monaco-editor.md
|       |   |   |-- gallery.md
|       |   |   |-- getting-started.md
|       |   |   |-- icons.md
|       |   |   |-- index.md
|       |   |   |-- menu.md
|       |   |   |-- navigation-view.md
|       |   |   |-- nuget.md
|       |   |   |-- releases.md
|       |   |   |-- symbolicon.md
|       |   |   |-- system-theme-watcher.md
|       |   |   `-- themes.md
|       |   |-- images
|       |   |   |-- favicon.ico
|       |   |   |-- github.svg
|       |   |   |-- icon-192x192.png
|       |   |   |-- icon-256x256.png
|       |   |   |-- icon-384x384.png
|       |   |   |-- icon-512x512.png
|       |   |   |-- ms-download.png
|       |   |   |-- nuget.svg
|       |   |   |-- vs22.svg
|       |   |   |-- wpfui-gallery.png
|       |   |   |-- wpfui-monaco-editor.png
|       |   |   `-- wpfui.png
|       |   |-- migration
|       |   |   |-- v2-migration.md
|       |   |   |-- v3-migration.md
|       |   |   `-- v4-migration.md
|       |   |-- templates
|       |   |   |-- wpfui
|       |   |   |   |-- layout
|       |   |   |   |   `-- _master.tmpl
|       |   |   |   |-- partials
|       |   |   |   |   |-- class.header.tmpl.partial
|       |   |   |   |   |-- class.memberpage.tmpl.partial
|       |   |   |   |   |-- class.tmpl.partial
|       |   |   |   |   |-- collection.tmpl.partial
|       |   |   |   |   |-- customMREFContent.tmpl.partial
|       |   |   |   |   |-- enum.tmpl.partial
|       |   |   |   |   |-- item.tmpl.partial
|       |   |   |   |   `-- namespace.tmpl.partial
|       |   |   |   |-- src
|       |   |   |   |   |-- docfx.scss
|       |   |   |   |   |-- docfx.ts
|       |   |   |   |   |-- dotnet.scss
|       |   |   |   |   |-- helper.test.ts
|       |   |   |   |   |-- helper.ts
|       |   |   |   |   |-- highlight.scss
|       |   |   |   |   |-- highlight.ts
|       |   |   |   |   |-- layout.scss
|       |   |   |   |   |-- markdown.scss
|       |   |   |   |   |-- markdown.ts
|       |   |   |   |   |-- mixins.scss
|       |   |   |   |   |-- nav.scss
|       |   |   |   |   |-- nav.ts
|       |   |   |   |   |-- options.d.ts
|       |   |   |   |   |-- search-worker.ts
|       |   |   |   |   |-- search.scss
|       |   |   |   |   |-- search.ts
|       |   |   |   |   |-- theme.ts
|       |   |   |   |   |-- toc.scss
|       |   |   |   |   |-- toc.ts
|       |   |   |   |   |-- wpfui-index-stats.ts
|       |   |   |   |   `-- wpfui.scss
|       |   |   |   |-- toc.json.js
|       |   |   |   `-- toc.json.tmpl
|       |   |   |-- .eslintrc.js
|       |   |   |-- .gitignore
|       |   |   |-- .stylelintrc.json
|       |   |   |-- build.js
|       |   |   |-- package-lock.json
|       |   |   |-- package.json
|       |   |   |-- README.md
|       |   |   `-- tsconfig.json
|       |   |-- .gitignore
|       |   |-- docfx.json
|       |   |-- index.md
|       |   |-- manifest.webmanifest
|       |   |-- robots.txt
|       |   `-- toc.yml
|       |-- samples
|       |   |-- Wpf.Ui.Demo.Console
|       |   |   |-- Models
|       |   |   |   `-- DataColor.cs
|       |   |   |-- Utilities
|       |   |   |   `-- ThemeUtilities.cs
|       |   |   |-- Views
|       |   |   |   |-- Pages
|       |   |   |   |   |-- DashboardPage.xaml
|       |   |   |   |   |-- DashboardPage.xaml.cs
|       |   |   |   |   |-- DataPage.xaml
|       |   |   |   |   |-- DataPage.xaml.cs
|       |   |   |   |   |-- SettingsPage.xaml
|       |   |   |   |   `-- SettingsPage.xaml.cs
|       |   |   |   |-- MainView.xaml
|       |   |   |   |-- MainView.xaml.cs
|       |   |   |   |-- SimpleView.xaml
|       |   |   |   `-- SimpleView.xaml.cs
|       |   |   |-- GlobalUsings.cs
|       |   |   |-- Program.cs
|       |   |   |-- Wpf.Ui.Demo.Console.csproj
|       |   |   `-- wpfui.ico
|       |   |-- Wpf.Ui.Demo.Dialogs
|       |   |   |-- Assets
|       |   |   |   |-- applicationIcon-1024.png
|       |   |   |   `-- applicationIcon-256.png
|       |   |   |-- app.manifest
|       |   |   |-- App.xaml
|       |   |   |-- App.xaml.cs
|       |   |   |-- applicationIcon.ico
|       |   |   |-- AssemblyInfo.cs
|       |   |   |-- MainWindow.xaml
|       |   |   |-- MainWindow.xaml.cs
|       |   |   `-- Wpf.Ui.Demo.Dialogs.csproj
|       |   |-- Wpf.Ui.Demo.Mvvm
|       |   |   |-- Assets
|       |   |   |   |-- applicationIcon-1024.png
|       |   |   |   `-- applicationIcon-256.png
|       |   |   |-- Helpers
|       |   |   |   `-- EnumToBooleanConverter.cs
|       |   |   |-- Models
|       |   |   |   |-- AppConfig.cs
|       |   |   |   `-- DataColor.cs
|       |   |   |-- Services
|       |   |   |   `-- ApplicationHostService.cs
|       |   |   |-- ViewModels
|       |   |   |   |-- DashboardViewModel.cs
|       |   |   |   |-- DataViewModel.cs
|       |   |   |   |-- MainWindowViewModel.cs
|       |   |   |   |-- SettingsViewModel.cs
|       |   |   |   `-- ViewModel.cs
|       |   |   |-- Views
|       |   |   |   |-- Pages
|       |   |   |   |   |-- DashboardPage.xaml
|       |   |   |   |   |-- DashboardPage.xaml.cs
|       |   |   |   |   |-- DataPage.xaml
|       |   |   |   |   |-- DataPage.xaml.cs
|       |   |   |   |   |-- SettingsPage.xaml
|       |   |   |   |   `-- SettingsPage.xaml.cs
|       |   |   |   |-- MainWindow.xaml
|       |   |   |   `-- MainWindow.xaml.cs
|       |   |   |-- app.manifest
|       |   |   |-- App.xaml
|       |   |   |-- App.xaml.cs
|       |   |   |-- applicationIcon.ico
|       |   |   |-- AssemblyInfo.cs
|       |   |   |-- GlobalUsings.cs
|       |   |   `-- Wpf.Ui.Demo.Mvvm.csproj
|       |   |-- Wpf.Ui.Demo.SetResources.Simple
|       |   |   |-- Assets
|       |   |   |   |-- applicationIcon-1024.png
|       |   |   |   `-- applicationIcon-256.png
|       |   |   |-- Models
|       |   |   |   |-- DataColor.cs
|       |   |   |   `-- DataGroup.cs
|       |   |   |-- Views
|       |   |   |   `-- Pages
|       |   |   |       |-- DashboardPage.xaml
|       |   |   |       |-- DashboardPage.xaml.cs
|       |   |   |       |-- DataPage.xaml
|       |   |   |       |-- DataPage.xaml.cs
|       |   |   |       |-- ExpanderPage.xaml
|       |   |   |       |-- ExpanderPage.xaml.cs
|       |   |   |       |-- SettingsPage.xaml
|       |   |   |       `-- SettingsPage.xaml.cs
|       |   |   |-- app.manifest
|       |   |   |-- App.xaml
|       |   |   |-- App.xaml.cs
|       |   |   |-- applicationIcon.ico
|       |   |   |-- AssemblyInfo.cs
|       |   |   |-- MainWindow.xaml
|       |   |   |-- MainWindow.xaml.cs
|       |   |   `-- Wpf.Ui.Demo.SetResources.Simple.csproj
|       |   `-- Wpf.Ui.Demo.Simple
|       |       |-- Assets
|       |       |   |-- applicationIcon-1024.png
|       |       |   `-- applicationIcon-256.png
|       |       |-- Models
|       |       |   `-- DataColor.cs
|       |       |-- Views
|       |       |   `-- Pages
|       |       |       |-- DashboardPage.xaml
|       |       |       |-- DashboardPage.xaml.cs
|       |       |       |-- DataPage.xaml
|       |       |       |-- DataPage.xaml.cs
|       |       |       |-- SettingsPage.xaml
|       |       |       `-- SettingsPage.xaml.cs
|       |       |-- app.manifest
|       |       |-- App.xaml
|       |       |-- App.xaml.cs
|       |       |-- applicationIcon.ico
|       |       |-- AssemblyInfo.cs
|       |       |-- MainWindow.xaml
|       |       |-- MainWindow.xaml.cs
|       |       `-- Wpf.Ui.Demo.Simple.csproj
|       |-- src
|       |   |-- Wpf.Ui
|       |   |   |-- Animations
|       |   |   |   |-- AnimationProperties.cs
|       |   |   |   |-- Transition.cs
|       |   |   |   `-- TransitionAnimationProvider.cs
|       |   |   |-- Appearance
|       |   |   |   |-- ApplicationAccentColorManager.cs
|       |   |   |   |-- ApplicationTheme.cs
|       |   |   |   |-- ApplicationThemeManager.cs
|       |   |   |   |-- ObservedWindow.cs
|       |   |   |   |-- ResourceDictionaryManager.cs
|       |   |   |   |-- SystemTheme.cs
|       |   |   |   |-- SystemThemeManager.cs
|       |   |   |   |-- SystemThemeWatcher.cs
|       |   |   |   |-- ThemeChangedEvent.cs
|       |   |   |   |-- UISettingsRCW.cs
|       |   |   |   `-- WindowBackgroundManager.cs
|       |   |   |-- AutomationPeers
|       |   |   |   `-- CardControlAutomationPeer.cs
|       |   |   |-- Controls
|       |   |   |   |-- AccessText
|       |   |   |   |   `-- AccessText.xaml
|       |   |   |   |-- Anchor
|       |   |   |   |   |-- Anchor.bmp
|       |   |   |   |   |-- Anchor.cs
|       |   |   |   |   `-- Anchor.xaml
|       |   |   |   |-- Arc
|       |   |   |   |   |-- Arc.bmp
|       |   |   |   |   `-- Arc.cs
|       |   |   |   |-- AutoSuggestBox
|       |   |   |   |   |-- AutoSuggestBox.bmp
|       |   |   |   |   |-- AutoSuggestBox.cs
|       |   |   |   |   |-- AutoSuggestBox.xaml
|       |   |   |   |   |-- AutoSuggestBoxQuerySubmittedEventArgs.cs
|       |   |   |   |   |-- AutoSuggestBoxSuggestionChosenEventArgs.cs
|       |   |   |   |   |-- AutoSuggestBoxTextChangedEventArgs.cs
|       |   |   |   |   `-- AutoSuggestionBoxTextChangeReason.cs
|       |   |   |   |-- Badge
|       |   |   |   |   |-- Badge.cs
|       |   |   |   |   `-- Badge.xaml
|       |   |   |   |-- BreadcrumbBar
|       |   |   |   |   |-- BreadcrumbBar.cs
|       |   |   |   |   |-- BreadcrumbBar.xaml
|       |   |   |   |   |-- BreadcrumbBarItem.cs
|       |   |   |   |   `-- BreadcrumbBarItemClickedEventArgs.cs
|       |   |   |   |-- Button
|       |   |   |   |   |-- Badge.bmp
|       |   |   |   |   |-- Button.cs
|       |   |   |   |   `-- Button.xaml
|       |   |   |   |-- Calendar
|       |   |   |   |   `-- Calendar.xaml
|       |   |   |   |-- CalendarDatePicker
|       |   |   |   |   |-- CalendarDatePicker.cs
|       |   |   |   |   `-- CalendarDatePicker.xaml
|       |   |   |   |-- Card
|       |   |   |   |   |-- Card.bmp
|       |   |   |   |   |-- Card.cs
|       |   |   |   |   `-- Card.xaml
|       |   |   |   |-- CardAction
|       |   |   |   |   |-- CardAction.bmp
|       |   |   |   |   |-- CardAction.cs
|       |   |   |   |   |-- CardAction.xaml
|       |   |   |   |   `-- CardActionAutomationPeer.cs
|       |   |   |   |-- CardColor
|       |   |   |   |   |-- CardColor.cs
|       |   |   |   |   `-- CardColor.xaml
|       |   |   |   |-- CardControl
|       |   |   |   |   |-- CardControl.cs
|       |   |   |   |   `-- CardControl.xaml
|       |   |   |   |-- CardExpander
|       |   |   |   |   |-- CardExpander.bmp
|       |   |   |   |   |-- CardExpander.cs
|       |   |   |   |   `-- CardExpander.xaml
|       |   |   |   |-- CheckBox
|       |   |   |   |   `-- CheckBox.xaml
|       |   |   |   |-- ClientAreaBorder
|       |   |   |   |   `-- ClientAreaBorder.cs
|       |   |   |   |-- ColorPicker
|       |   |   |   |   |-- ColorPicker.cs
|       |   |   |   |   `-- ColorPicker.xaml
|       |   |   |   |-- ComboBox
|       |   |   |   |   `-- ComboBox.xaml
|       |   |   |   |-- ContentDialog
|       |   |   |   |   |-- EventArgs
|       |   |   |   |   |   |-- ContentDialogButtonClickEventArgs.cs
|       |   |   |   |   |   |-- ContentDialogClosedEventArgs.cs
|       |   |   |   |   |   `-- ContentDialogClosingEventArgs.cs
|       |   |   |   |   |-- ContentDialog.cs
|       |   |   |   |   |-- ContentDialog.xaml
|       |   |   |   |   |-- ContentDialogButton.cs
|       |   |   |   |   `-- ContentDialogResult.cs
|       |   |   |   |-- ContextMenu
|       |   |   |   |   |-- ContextMenu.xaml
|       |   |   |   |   |-- ContextMenuLoader.xaml
|       |   |   |   |   `-- ContextMenuLoader.xaml.cs
|       |   |   |   |-- DataGrid
|       |   |   |   |   |-- DataGrid.cs
|       |   |   |   |   `-- DataGrid.xaml
|       |   |   |   |-- DatePicker
|       |   |   |   |   `-- DatePicker.xaml
|       |   |   |   |-- DropDownButton
|       |   |   |   |   |-- DropDownButton.cs
|       |   |   |   |   `-- DropDownButton.xaml
|       |   |   |   |-- DynamicScrollBar
|       |   |   |   |   |-- DynamicScrollBar.bmp
|       |   |   |   |   |-- DynamicScrollBar.cs
|       |   |   |   |   `-- DynamicScrollBar.xaml
|       |   |   |   |-- DynamicScrollViewer
|       |   |   |   |   |-- DynamicScrollViewer.bmp
|       |   |   |   |   |-- DynamicScrollViewer.cs
|       |   |   |   |   `-- DynamicScrollViewer.xaml
|       |   |   |   |-- Expander
|       |   |   |   |   `-- Expander.xaml
|       |   |   |   |-- FluentWindow
|       |   |   |   |   |-- FluentWindow.bmp
|       |   |   |   |   |-- FluentWindow.cs
|       |   |   |   |   `-- FluentWindow.xaml
|       |   |   |   |-- Flyout
|       |   |   |   |   |-- Flyout.cs
|       |   |   |   |   `-- Flyout.xaml
|       |   |   |   |-- Frame
|       |   |   |   |   `-- Frame.xaml
|       |   |   |   |-- GridView
|       |   |   |   |   |-- GridView.cs
|       |   |   |   |   |-- GridViewColumn.cs
|       |   |   |   |   |-- GridViewColumnHeader.xaml
|       |   |   |   |   |-- GridViewHeaderRowIndicator.xaml
|       |   |   |   |   |-- GridViewHeaderRowPresenter.cs
|       |   |   |   |   `-- GridViewRowPresenter.cs
|       |   |   |   |-- HyperlinkButton
|       |   |   |   |   |-- HyperlinkButton.cs
|       |   |   |   |   `-- HyperlinkButton.xaml
|       |   |   |   |-- IconElement
|       |   |   |   |   |-- FontIcon.bmp
|       |   |   |   |   |-- FontIcon.cs
|       |   |   |   |   |-- IconElement.cs
|       |   |   |   |   |-- IconElementConverter.cs
|       |   |   |   |   |-- IconSourceElement.cs
|       |   |   |   |   |-- ImageIcon.cs
|       |   |   |   |   |-- SymbolIcon.bmp
|       |   |   |   |   `-- SymbolIcon.cs
|       |   |   |   |-- IconSource
|       |   |   |   |   |-- FontIconSource.cs
|       |   |   |   |   |-- IconSource.cs
|       |   |   |   |   `-- SymbolIconSource.cs
|       |   |   |   |-- Image
|       |   |   |   |   |-- Image.cs
|       |   |   |   |   `-- Image.xaml
|       |   |   |   |-- InfoBadge
|       |   |   |   |   |-- InfoBadge.cs
|       |   |   |   |   |-- InfoBadge.xaml
|       |   |   |   |   `-- InfoBadgeSeverity.cs
|       |   |   |   |-- InfoBar
|       |   |   |   |   |-- InfoBar.cs
|       |   |   |   |   |-- InfoBar.xaml
|       |   |   |   |   `-- InfoBarSeverity.cs
|       |   |   |   |-- ItemsControl
|       |   |   |   |   `-- ItemsControl.xaml
|       |   |   |   |-- Label
|       |   |   |   |   `-- Label.xaml
|       |   |   |   |-- ListBox
|       |   |   |   |   |-- ListBox.xaml
|       |   |   |   |   `-- ListBoxItem.xaml
|       |   |   |   |-- ListView
|       |   |   |   |   |-- ListView.cs
|       |   |   |   |   |-- ListView.xaml
|       |   |   |   |   |-- ListViewItem.cs
|       |   |   |   |   |-- ListViewItem.xaml
|       |   |   |   |   `-- ListViewViewState.cs
|       |   |   |   |-- LoadingScreen
|       |   |   |   |   |-- LoadingScreen.cs
|       |   |   |   |   `-- LoadingScreen.xaml
|       |   |   |   |-- Menu
|       |   |   |   |   |-- Menu.xaml
|       |   |   |   |   |-- MenuItem.cs
|       |   |   |   |   |-- MenuItem.xaml
|       |   |   |   |   |-- MenuLoader.xaml
|       |   |   |   |   `-- MenuLoader.xaml.cs
|       |   |   |   |-- MessageBox
|       |   |   |   |   |-- MessageBox.bmp
|       |   |   |   |   |-- MessageBox.cs
|       |   |   |   |   |-- MessageBox.xaml
|       |   |   |   |   |-- MessageBoxButton.cs
|       |   |   |   |   `-- MessageBoxResult.cs
|       |   |   |   |-- NavigationView
|       |   |   |   |   |-- INavigationView.cs
|       |   |   |   |   |-- INavigationViewItem.cs
|       |   |   |   |   |-- NavigatedEventArgs.cs
|       |   |   |   |   |-- NavigatingCancelEventArgs.cs
|       |   |   |   |   |-- NavigationCache.cs
|       |   |   |   |   |-- NavigationCacheMode.cs
|       |   |   |   |   |-- NavigationLeftFluent.xaml
|       |   |   |   |   |-- NavigationView.AttachedProperties.cs
|       |   |   |   |   |-- NavigationView.Base.cs
|       |   |   |   |   |-- NavigationView.bmp
|       |   |   |   |   |-- NavigationView.Events.cs
|       |   |   |   |   |-- NavigationView.Navigation.cs
|       |   |   |   |   |-- NavigationView.Properties.cs
|       |   |   |   |   |-- NavigationView.TemplateParts.cs
|       |   |   |   |   |-- NavigationView.xaml
|       |   |   |   |   |-- NavigationViewActivator.cs
|       |   |   |   |   |-- NavigationViewBackButtonVisible.cs
|       |   |   |   |   |-- NavigationViewBasePaneButtonStyle.xaml
|       |   |   |   |   |-- NavigationViewBottom.xaml
|       |   |   |   |   |-- NavigationViewBreadcrumbItem.cs
|       |   |   |   |   |-- NavigationViewBreadcrumbItem.xaml
|       |   |   |   |   |-- NavigationViewCompact.xaml
|       |   |   |   |   |-- NavigationViewConstants.xaml
|       |   |   |   |   |-- NavigationViewContentPresenter.cs
|       |   |   |   |   |-- NavigationViewContentPresenter.xaml
|       |   |   |   |   |-- NavigationViewItem.bmp
|       |   |   |   |   |-- NavigationViewItem.cs
|       |   |   |   |   |-- NavigationViewItemAutomationPeer.cs
|       |   |   |   |   |-- NavigationViewItemDefaultStyle.xaml
|       |   |   |   |   |-- NavigationViewItemHeader.cs
|       |   |   |   |   |-- NavigationViewItemHeader.xaml
|       |   |   |   |   |-- NavigationViewItemSeparator.cs
|       |   |   |   |   |-- NavigationViewItemSeparator.xaml
|       |   |   |   |   |-- NavigationViewLeftMinimalCompact.xaml
|       |   |   |   |   |-- NavigationViewPaneDisplayMode.cs
|       |   |   |   |   `-- NavigationViewTop.xaml
|       |   |   |   |-- NumberBox
|       |   |   |   |   |-- INumberFormatter.cs
|       |   |   |   |   |-- INumberParser.cs
|       |   |   |   |   |-- NumberBox.bmp
|       |   |   |   |   |-- NumberBox.cs
|       |   |   |   |   |-- NumberBox.xaml
|       |   |   |   |   |-- NumberBoxSpinButtonPlacementMode.cs
|       |   |   |   |   |-- NumberBoxValidationMode.cs
|       |   |   |   |   |-- NumberBoxValueChangedEventArgs.cs
|       |   |   |   |   `-- ValidateNumberFormatter.cs
|       |   |   |   |-- Page
|       |   |   |   |   `-- Page.xaml
|       |   |   |   |-- PasswordBox
|       |   |   |   |   |-- PasswordBox.cs
|       |   |   |   |   |-- PasswordBox.xaml
|       |   |   |   |   `-- PasswordHelper.cs
|       |   |   |   |-- ProgressBar
|       |   |   |   |   `-- ProgressBar.xaml
|       |   |   |   |-- ProgressRing
|       |   |   |   |   |-- ProgressRing.bmp
|       |   |   |   |   |-- ProgressRing.cs
|       |   |   |   |   `-- ProgressRing.xaml
|       |   |   |   |-- RadioButton
|       |   |   |   |   `-- RadioButton.xaml
|       |   |   |   |-- RatingControl
|       |   |   |   |   |-- RatingControl.bmp
|       |   |   |   |   |-- RatingControl.cs
|       |   |   |   |   `-- RatingControl.xaml
|       |   |   |   |-- RichTextBox
|       |   |   |   |   |-- RichTextBox.cs
|       |   |   |   |   `-- RichTextBox.xaml
|       |   |   |   |-- ScrollBar
|       |   |   |   |   `-- ScrollBar.xaml
|       |   |   |   |-- ScrollViewer
|       |   |   |   |   `-- ScrollViewer.xaml
|       |   |   |   |-- Separator
|       |   |   |   |   `-- Separator.xaml
|       |   |   |   |-- Slider
|       |   |   |   |   `-- Slider.xaml
|       |   |   |   |-- Snackbar
|       |   |   |   |   |-- Snackbar.cs
|       |   |   |   |   |-- Snackbar.xaml
|       |   |   |   |   `-- SnackbarPresenter.cs
|       |   |   |   |-- SplitButton
|       |   |   |   |   |-- SplitButton.cs
|       |   |   |   |   `-- SplitButton.xaml
|       |   |   |   |-- StatusBar
|       |   |   |   |   `-- StatusBar.xaml
|       |   |   |   |-- TabControl
|       |   |   |   |   `-- TabControl.xaml
|       |   |   |   |-- TabView
|       |   |   |   |   |-- TabView.cs
|       |   |   |   |   `-- TabViewItem.cs
|       |   |   |   |-- TextBlock
|       |   |   |   |   |-- TextBlock.cs
|       |   |   |   |   `-- TextBlock.xaml
|       |   |   |   |-- TextBox
|       |   |   |   |   |-- TextBox.cs
|       |   |   |   |   `-- TextBox.xaml
|       |   |   |   |-- ThumbRate
|       |   |   |   |   |-- ThumbRate.bmp
|       |   |   |   |   |-- ThumbRate.cs
|       |   |   |   |   |-- ThumbRate.xaml
|       |   |   |   |   `-- ThumbRateState.cs
|       |   |   |   |-- TimePicker
|       |   |   |   |   |-- ClockIdentifier.cs
|       |   |   |   |   |-- TimePicker.cs
|       |   |   |   |   `-- TimePicker.xaml
|       |   |   |   |-- TitleBar
|       |   |   |   |   |-- HwndProcEventArgs.cs
|       |   |   |   |   |-- TitleBar.cs
|       |   |   |   |   |-- TitleBar.WindowResize.cs
|       |   |   |   |   |-- TitleBar.xaml
|       |   |   |   |   |-- TitleBarButton.cs
|       |   |   |   |   `-- TitleBarButtonType.cs
|       |   |   |   |-- ToggleButton
|       |   |   |   |   `-- ToggleButton.xaml
|       |   |   |   |-- ToggleSwitch
|       |   |   |   |   |-- ToggleSwitch.bmp
|       |   |   |   |   |-- ToggleSwitch.cs
|       |   |   |   |   `-- ToggleSwitch.xaml
|       |   |   |   |-- ToolBar
|       |   |   |   |   `-- ToolBar.xaml
|       |   |   |   |-- ToolTip
|       |   |   |   |   `-- ToolTip.xaml
|       |   |   |   |-- TreeGrid
|       |   |   |   |   |-- TreeGrid.cs
|       |   |   |   |   |-- TreeGrid.xaml
|       |   |   |   |   |-- TreeGridHeader.cs
|       |   |   |   |   `-- TreeGridItem.cs
|       |   |   |   |-- TreeView
|       |   |   |   |   |-- TreeView.xaml
|       |   |   |   |   |-- TreeViewItem.cs
|       |   |   |   |   `-- TreeViewItem.xaml
|       |   |   |   |-- VirtualizingGridView
|       |   |   |   |   |-- VirtualizingGridView.cs
|       |   |   |   |   `-- VirtualizingGridView.xaml
|       |   |   |   |-- VirtualizingItemsControl
|       |   |   |   |   |-- VirtualizingItemsControl.bmp
|       |   |   |   |   |-- VirtualizingItemsControl.cs
|       |   |   |   |   `-- VirtualizingItemsControl.xaml
|       |   |   |   |-- VirtualizingUniformGrid
|       |   |   |   |   `-- VirtualizingUniformGrid.cs
|       |   |   |   |-- VirtualizingWrapPanel
|       |   |   |   |   |-- VirtualizingPanelBase.cs
|       |   |   |   |   |-- VirtualizingWrapPanel.bmp
|       |   |   |   |   |-- VirtualizingWrapPanel.cs
|       |   |   |   |   `-- VirtualizingWrapPanel.xaml
|       |   |   |   |-- Window
|       |   |   |   |   |-- Window.xaml
|       |   |   |   |   |-- WindowBackdrop.cs
|       |   |   |   |   |-- WindowBackdropType.cs
|       |   |   |   |   `-- WindowCornerPreference.cs
|       |   |   |   |-- ControlAppearance.cs
|       |   |   |   |-- ControlsServices.cs
|       |   |   |   |-- DateTimeHelper.cs
|       |   |   |   |-- EffectThicknessDecorator.cs
|       |   |   |   |-- ElementPlacement.cs
|       |   |   |   |-- EventIdentifier.cs
|       |   |   |   |-- FontTypography.cs
|       |   |   |   |-- IAppearanceControl.cs
|       |   |   |   |-- IDpiAwareControl.cs
|       |   |   |   |-- IIconControl.cs
|       |   |   |   |-- ItemRange.cs
|       |   |   |   |-- IThemeControl.cs
|       |   |   |   |-- PassiveScrollViewer.cs
|       |   |   |   |-- ScrollDirection.cs
|       |   |   |   |-- SpacingMode.cs
|       |   |   |   |-- SymbolFilled.cs
|       |   |   |   |-- SymbolGlyph.cs
|       |   |   |   |-- SymbolRegular.cs
|       |   |   |   |-- TextColor.cs
|       |   |   |   `-- TypedEventHandler.cs
|       |   |   |-- Converters
|       |   |   |   |-- AnimationFactorToValueConverter.cs
|       |   |   |   |-- BackButtonVisibilityToVisibilityConverter.cs
|       |   |   |   |-- BoolToVisibilityConverter.cs
|       |   |   |   |-- BrushToColorConverter.cs
|       |   |   |   |-- ClipConverter.cs
|       |   |   |   |-- ContentDialogButtonEnumToBoolConverter.cs
|       |   |   |   |-- CornerRadiusSplitConverter.cs
|       |   |   |   |-- DatePickerButtonPaddingConverter.cs
|       |   |   |   |-- EnumToBoolConverter.cs
|       |   |   |   |-- FallbackBrushConverter.cs
|       |   |   |   |-- IconSourceElementConverter.cs
|       |   |   |   |-- LeftSplitCornerRadiusConverter.cs
|       |   |   |   |-- LeftSplitThicknessConverter.cs
|       |   |   |   |-- NullToVisibilityConverter.cs
|       |   |   |   |-- ProgressThicknessConverter.cs
|       |   |   |   |-- RightSplitCornerRadiusConverter.cs
|       |   |   |   |-- RightSplitThicknessConverter.cs
|       |   |   |   `-- TextToAsteriskConverter.cs
|       |   |   |-- Designer
|       |   |   |   `-- DesignerHelper.cs
|       |   |   |-- Extensions
|       |   |   |   |-- ColorExtensions.cs
|       |   |   |   |-- ContentDialogServiceExtensions.cs
|       |   |   |   |-- ContextMenuExtensions.cs
|       |   |   |   |-- DateTimeExtensions.cs
|       |   |   |   |-- FrameExtensions.cs
|       |   |   |   |-- NavigationServiceExtensions.cs
|       |   |   |   |-- PInvokeExtensions.cs
|       |   |   |   |-- SnackbarServiceExtensions.cs
|       |   |   |   |-- StringExtensions.cs
|       |   |   |   |-- SymbolExtensions.cs
|       |   |   |   |-- TextBlockFontTypographyExtensions.cs
|       |   |   |   |-- TextColorExtensions.cs
|       |   |   |   |-- UiElementExtensions.cs
|       |   |   |   `-- UriExtensions.cs
|       |   |   |-- Hardware
|       |   |   |   |-- DisplayDpi.cs
|       |   |   |   |-- DpiHelper.cs
|       |   |   |   |-- HardwareAcceleration.cs
|       |   |   |   `-- RenderingTier.cs
|       |   |   |-- Input
|       |   |   |   |-- IRelayCommand.cs
|       |   |   |   |-- IRelayCommand{T}.cs
|       |   |   |   `-- RelayCommand{T}.cs
|       |   |   |-- Interop
|       |   |   |   |-- PInvoke.cs
|       |   |   |   |-- UnsafeNativeMethods.cs
|       |   |   |   `-- UnsafeReflection.cs
|       |   |   |-- Markup
|       |   |   |   |-- ControlsDictionary.cs
|       |   |   |   |-- Design.cs
|       |   |   |   |-- FontIconExtension.cs
|       |   |   |   |-- ImageIconExtension.cs
|       |   |   |   |-- SymbolIconExtension.cs
|       |   |   |   |-- ThemeResource.cs
|       |   |   |   |-- ThemeResourceExtension.cs
|       |   |   |   `-- ThemesDictionary.cs
|       |   |   |-- Properties
|       |   |   |   `-- AssemblyInfo.cs
|       |   |   |-- Resources
|       |   |   |   |-- Fonts
|       |   |   |   |   |-- FluentSystemIcons-Filled.ttf
|       |   |   |   |   `-- FluentSystemIcons-Regular.ttf
|       |   |   |   |-- Theme
|       |   |   |   |   |-- Dark.xaml
|       |   |   |   |   |-- HC1.xaml
|       |   |   |   |   |-- HC2.xaml
|       |   |   |   |   |-- HCBlack.xaml
|       |   |   |   |   |-- HCWhite.xaml
|       |   |   |   |   `-- Light.xaml
|       |   |   |   |-- Accent.xaml
|       |   |   |   |-- DefaultContextMenu.xaml
|       |   |   |   |-- DefaultFocusVisualStyle.xaml
|       |   |   |   |-- DefaultTextBoxScrollViewerStyle.xaml
|       |   |   |   |-- Fonts.xaml
|       |   |   |   |-- Palette.xaml
|       |   |   |   |-- StaticColors.xaml
|       |   |   |   |-- Typography.xaml
|       |   |   |   |-- Variables.xaml
|       |   |   |   `-- Wpf.Ui.xaml
|       |   |   |-- Taskbar
|       |   |   |   |-- TaskbarProgress.cs
|       |   |   |   `-- TaskbarProgressState.cs
|       |   |   |-- Win32
|       |   |   |   `-- Utilities.cs
|       |   |   |-- ContentDialogService.cs
|       |   |   |-- GlobalUsings.cs
|       |   |   |-- IContentDialogService.cs
|       |   |   |-- INavigationService.cs
|       |   |   |-- INavigationWindow.cs
|       |   |   |-- ISnackbarService.cs
|       |   |   |-- ITaskBarService.cs
|       |   |   |-- IThemeService.cs
|       |   |   |-- NativeMethods.txt
|       |   |   |-- NavigationService.cs
|       |   |   |-- SimpleContentDialogCreateOptions.cs
|       |   |   |-- SnackbarService.cs
|       |   |   |-- TaskBarService.cs
|       |   |   |-- ThemeService.cs
|       |   |   |-- UiApplication.cs
|       |   |   |-- UiAssembly.cs
|       |   |   |-- VisualStudioToolsManifest.xml
|       |   |   `-- Wpf.Ui.csproj
|       |   |-- Wpf.Ui.Abstractions
|       |   |   |-- Controls
|       |   |   |   |-- INavigableView.cs
|       |   |   |   |-- INavigationAware.cs
|       |   |   |   `-- NavigationAware.cs
|       |   |   |-- GlobalUsings.cs
|       |   |   |-- INavigationViewPageProvider.cs
|       |   |   |-- NavigationException.cs
|       |   |   |-- NavigationViewPageProviderExtensions.cs
|       |   |   `-- Wpf.Ui.Abstractions.csproj
|       |   |-- Wpf.Ui.DependencyInjection
|       |   |   |-- DependencyInjectionNavigationViewPageProvider.cs
|       |   |   |-- GlobalUsings.cs
|       |   |   |-- ServiceCollectionExtensions.cs
|       |   |   `-- Wpf.Ui.DependencyInjection.csproj
|       |   |-- Wpf.Ui.Extension
|       |   |   |-- license.txt
|       |   |   |-- preview.png
|       |   |   |-- source.extension.vsixmanifest
|       |   |   |-- Wpf.Ui.Extension.csproj
|       |   |   `-- wpfui.png
|       |   |-- Wpf.Ui.Extension.Template.Blank
|       |   |   |-- Assets
|       |   |   |   |-- wpfui-icon-1024.png
|       |   |   |   `-- wpfui-icon-256.png
|       |   |   |-- __PreviewImage.png
|       |   |   |-- __TemplateIcon.png
|       |   |   |-- app.manifest
|       |   |   |-- App.xaml
|       |   |   |-- App.xaml.cs
|       |   |   |-- AssemblyInfo.cs
|       |   |   |-- Wpf.Ui.Blank.csproj
|       |   |   |-- Wpf.Ui.Blank.vstemplate
|       |   |   |-- Wpf.Ui.Extension.Template.Blank.csproj
|       |   |   `-- wpfui-icon.ico
|       |   |-- Wpf.Ui.Extension.Template.Compact
|       |   |   |-- Assets
|       |   |   |   |-- wpfui-icon-1024.png
|       |   |   |   `-- wpfui-icon-256.png
|       |   |   |-- Helpers
|       |   |   |   `-- EnumToBooleanConverter.cs
|       |   |   |-- Models
|       |   |   |   |-- AppConfig.cs
|       |   |   |   `-- DataColor.cs
|       |   |   |-- Resources
|       |   |   |   `-- Translations.cs
|       |   |   |-- Services
|       |   |   |   `-- ApplicationHostService.cs
|       |   |   |-- ViewModels
|       |   |   |   |-- Pages
|       |   |   |   |   |-- DashboardViewModel.cs
|       |   |   |   |   |-- DataViewModel.cs
|       |   |   |   |   `-- SettingsViewModel.cs
|       |   |   |   `-- Windows
|       |   |   |       `-- MainWindowViewModel.cs
|       |   |   |-- Views
|       |   |   |   |-- Pages
|       |   |   |   |   |-- DashboardPage.xaml
|       |   |   |   |   |-- DashboardPage.xaml.cs
|       |   |   |   |   |-- DataPage.xaml
|       |   |   |   |   |-- DataPage.xaml.cs
|       |   |   |   |   |-- SettingsPage.xaml
|       |   |   |   |   `-- SettingsPage.xaml.cs
|       |   |   |   `-- Windows
|       |   |   |       |-- MainWindow.xaml
|       |   |   |       `-- MainWindow.xaml.cs
|       |   |   |-- __PreviewImage.png
|       |   |   |-- __TemplateIcon.png
|       |   |   |-- app.manifest
|       |   |   |-- App.xaml
|       |   |   |-- App.xaml.cs
|       |   |   |-- AssemblyInfo.cs
|       |   |   |-- Usings.cs
|       |   |   |-- Wpf.Ui.Compact.csproj
|       |   |   |-- Wpf.Ui.Compact.vstemplate
|       |   |   |-- Wpf.Ui.Extension.Template.Compact.csproj
|       |   |   `-- wpfui-icon.ico
|       |   |-- Wpf.Ui.Extension.Template.Fluent
|       |   |   |-- Assets
|       |   |   |   |-- wpfui-icon-1024.png
|       |   |   |   `-- wpfui-icon-256.png
|       |   |   |-- Helpers
|       |   |   |   `-- EnumToBooleanConverter.cs
|       |   |   |-- Models
|       |   |   |   |-- AppConfig.cs
|       |   |   |   `-- DataColor.cs
|       |   |   |-- Resources
|       |   |   |   `-- Translations.cs
|       |   |   |-- Services
|       |   |   |   `-- ApplicationHostService.cs
|       |   |   |-- ViewModels
|       |   |   |   |-- Pages
|       |   |   |   |   |-- DashboardViewModel.cs
|       |   |   |   |   |-- DataViewModel.cs
|       |   |   |   |   `-- SettingsViewModel.cs
|       |   |   |   `-- Windows
|       |   |   |       `-- MainWindowViewModel.cs
|       |   |   |-- Views
|       |   |   |   |-- Pages
|       |   |   |   |   |-- DashboardPage.xaml
|       |   |   |   |   |-- DashboardPage.xaml.cs
|       |   |   |   |   |-- DataPage.xaml
|       |   |   |   |   |-- DataPage.xaml.cs
|       |   |   |   |   |-- SettingsPage.xaml
|       |   |   |   |   `-- SettingsPage.xaml.cs
|       |   |   |   `-- Windows
|       |   |   |       |-- MainWindow.xaml
|       |   |   |       `-- MainWindow.xaml.cs
|       |   |   |-- __PreviewImage.png
|       |   |   |-- __TemplateIcon.png
|       |   |   |-- app.manifest
|       |   |   |-- App.xaml
|       |   |   |-- App.xaml.cs
|       |   |   |-- AssemblyInfo.cs
|       |   |   |-- Usings.cs
|       |   |   |-- Wpf.Ui.Extension.Template.Fluent.csproj
|       |   |   |-- Wpf.Ui.Fluent.csproj
|       |   |   |-- Wpf.Ui.Fluent.vstemplate
|       |   |   `-- wpfui-icon.ico
|       |   |-- Wpf.Ui.FlaUI
|       |   |   |-- AutoSuggestBox.cs
|       |   |   |-- GlobalUsings.cs
|       |   |   `-- Wpf.Ui.FlaUI.csproj
|       |   |-- Wpf.Ui.FontMapper
|       |   |   |-- FontSource.cs
|       |   |   |-- GitTag.cs
|       |   |   |-- GlobalUsings.cs
|       |   |   |-- License - Fluent System Icons.txt
|       |   |   |-- Program.cs
|       |   |   `-- Wpf.Ui.FontMapper.csproj
|       |   |-- Wpf.Ui.Gallery
|       |   |   |-- Assets
|       |   |   |   |-- Monaco
|       |   |   |   |   |-- min
|       |   |   |   |   |   `-- vs
|       |   |   |   |   |       |-- base
|       |   |   |   |   |       |   |-- browser
|       |   |   |   |   |       |   |   `-- ui
|       |   |   |   |   |       |   |       `-- codicons
|       |   |   |   |   |       |   |           `-- codicon
|       |   |   |   |   |       |   |               `-- codicon.ttf
|       |   |   |   |   |       |   |-- common
|       |   |   |   |   |       |   |   `-- worker
|       |   |   |   |   |       |   |       |-- simpleWorker.nls.de.js
|       |   |   |   |   |       |   |       |-- simpleWorker.nls.es.js
|       |   |   |   |   |       |   |       |-- simpleWorker.nls.fr.js
|       |   |   |   |   |       |   |       |-- simpleWorker.nls.it.js
|       |   |   |   |   |       |   |       |-- simpleWorker.nls.ja.js
|       |   |   |   |   |       |   |       |-- simpleWorker.nls.js
|       |   |   |   |   |       |   |       |-- simpleWorker.nls.ko.js
|       |   |   |   |   |       |   |       |-- simpleWorker.nls.ru.js
|       |   |   |   |   |       |   |       |-- simpleWorker.nls.zh-cn.js
|       |   |   |   |   |       |   |       `-- simpleWorker.nls.zh-tw.js
|       |   |   |   |   |       |   `-- worker
|       |   |   |   |   |       |       `-- workerMain.js
|       |   |   |   |   |       |-- basic-languages
|       |   |   |   |   |       |   |-- abap
|       |   |   |   |   |       |   |   `-- abap.js
|       |   |   |   |   |       |   |-- apex
|       |   |   |   |   |       |   |   `-- apex.js
|       |   |   |   |   |       |   |-- azcli
|       |   |   |   |   |       |   |   `-- azcli.js
|       |   |   |   |   |       |   |-- bat
|       |   |   |   |   |       |   |   `-- bat.js
|       |   |   |   |   |       |   |-- bicep
|       |   |   |   |   |       |   |   `-- bicep.js
|       |   |   |   |   |       |   |-- cameligo
|       |   |   |   |   |       |   |   `-- cameligo.js
|       |   |   |   |   |       |   |-- clojure
|       |   |   |   |   |       |   |   `-- clojure.js
|       |   |   |   |   |       |   |-- coffee
|       |   |   |   |   |       |   |   `-- coffee.js
|       |   |   |   |   |       |   |-- cpp
|       |   |   |   |   |       |   |   `-- cpp.js
|       |   |   |   |   |       |   |-- csharp
|       |   |   |   |   |       |   |   `-- csharp.js
|       |   |   |   |   |       |   |-- csp
|       |   |   |   |   |       |   |   `-- csp.js
|       |   |   |   |   |       |   |-- css
|       |   |   |   |   |       |   |   `-- css.js
|       |   |   |   |   |       |   |-- cypher
|       |   |   |   |   |       |   |   `-- cypher.js
|       |   |   |   |   |       |   |-- dart
|       |   |   |   |   |       |   |   `-- dart.js
|       |   |   |   |   |       |   |-- dockerfile
|       |   |   |   |   |       |   |   `-- dockerfile.js
|       |   |   |   |   |       |   |-- ecl
|       |   |   |   |   |       |   |   `-- ecl.js
|       |   |   |   |   |       |   |-- elixir
|       |   |   |   |   |       |   |   `-- elixir.js
|       |   |   |   |   |       |   |-- flow9
|       |   |   |   |   |       |   |   `-- flow9.js
|       |   |   |   |   |       |   |-- freemarker2
|       |   |   |   |   |       |   |   `-- freemarker2.js
|       |   |   |   |   |       |   |-- fsharp
|       |   |   |   |   |       |   |   `-- fsharp.js
|       |   |   |   |   |       |   |-- go
|       |   |   |   |   |       |   |   `-- go.js
|       |   |   |   |   |       |   |-- graphql
|       |   |   |   |   |       |   |   `-- graphql.js
|       |   |   |   |   |       |   |-- handlebars
|       |   |   |   |   |       |   |   `-- handlebars.js
|       |   |   |   |   |       |   |-- hcl
|       |   |   |   |   |       |   |   `-- hcl.js
|       |   |   |   |   |       |   |-- html
|       |   |   |   |   |       |   |   `-- html.js
|       |   |   |   |   |       |   |-- ini
|       |   |   |   |   |       |   |   `-- ini.js
|       |   |   |   |   |       |   |-- java
|       |   |   |   |   |       |   |   `-- java.js
|       |   |   |   |   |       |   |-- javascript
|       |   |   |   |   |       |   |   `-- javascript.js
|       |   |   |   |   |       |   |-- julia
|       |   |   |   |   |       |   |   `-- julia.js
|       |   |   |   |   |       |   |-- kotlin
|       |   |   |   |   |       |   |   `-- kotlin.js
|       |   |   |   |   |       |   |-- less
|       |   |   |   |   |       |   |   `-- less.js
|       |   |   |   |   |       |   |-- lexon
|       |   |   |   |   |       |   |   `-- lexon.js
|       |   |   |   |   |       |   |-- liquid
|       |   |   |   |   |       |   |   `-- liquid.js
|       |   |   |   |   |       |   |-- lua
|       |   |   |   |   |       |   |   `-- lua.js
|       |   |   |   |   |       |   |-- m3
|       |   |   |   |   |       |   |   `-- m3.js
|       |   |   |   |   |       |   |-- markdown
|       |   |   |   |   |       |   |   `-- markdown.js
|       |   |   |   |   |       |   |-- mdx
|       |   |   |   |   |       |   |   `-- mdx.js
|       |   |   |   |   |       |   |-- mips
|       |   |   |   |   |       |   |   `-- mips.js
|       |   |   |   |   |       |   |-- msdax
|       |   |   |   |   |       |   |   `-- msdax.js
|       |   |   |   |   |       |   |-- mysql
|       |   |   |   |   |       |   |   `-- mysql.js
|       |   |   |   |   |       |   |-- objective-c
|       |   |   |   |   |       |   |   `-- objective-c.js
|       |   |   |   |   |       |   |-- pascal
|       |   |   |   |   |       |   |   `-- pascal.js
|       |   |   |   |   |       |   |-- pascaligo
|       |   |   |   |   |       |   |   `-- pascaligo.js
|       |   |   |   |   |       |   |-- perl
|       |   |   |   |   |       |   |   `-- perl.js
|       |   |   |   |   |       |   |-- pgsql
|       |   |   |   |   |       |   |   `-- pgsql.js
|       |   |   |   |   |       |   |-- php
|       |   |   |   |   |       |   |   `-- php.js
|       |   |   |   |   |       |   |-- pla
|       |   |   |   |   |       |   |   `-- pla.js
|       |   |   |   |   |       |   |-- postiats
|       |   |   |   |   |       |   |   `-- postiats.js
|       |   |   |   |   |       |   |-- powerquery
|       |   |   |   |   |       |   |   `-- powerquery.js
|       |   |   |   |   |       |   |-- powershell
|       |   |   |   |   |       |   |   `-- powershell.js
|       |   |   |   |   |       |   |-- protobuf
|       |   |   |   |   |       |   |   `-- protobuf.js
|       |   |   |   |   |       |   |-- pug
|       |   |   |   |   |       |   |   `-- pug.js
|       |   |   |   |   |       |   |-- python
|       |   |   |   |   |       |   |   `-- python.js
|       |   |   |   |   |       |   |-- qsharp
|       |   |   |   |   |       |   |   `-- qsharp.js
|       |   |   |   |   |       |   |-- r
|       |   |   |   |   |       |   |   `-- r.js
|       |   |   |   |   |       |   |-- razor
|       |   |   |   |   |       |   |   `-- razor.js
|       |   |   |   |   |       |   |-- redis
|       |   |   |   |   |       |   |   `-- redis.js
|       |   |   |   |   |       |   |-- redshift
|       |   |   |   |   |       |   |   `-- redshift.js
|       |   |   |   |   |       |   |-- restructuredtext
|       |   |   |   |   |       |   |   `-- restructuredtext.js
|       |   |   |   |   |       |   |-- ruby
|       |   |   |   |   |       |   |   `-- ruby.js
|       |   |   |   |   |       |   |-- rust
|       |   |   |   |   |       |   |   `-- rust.js
|       |   |   |   |   |       |   |-- sb
|       |   |   |   |   |       |   |   `-- sb.js
|       |   |   |   |   |       |   |-- scala
|       |   |   |   |   |       |   |   `-- scala.js
|       |   |   |   |   |       |   |-- scheme
|       |   |   |   |   |       |   |   `-- scheme.js
|       |   |   |   |   |       |   |-- scss
|       |   |   |   |   |       |   |   `-- scss.js
|       |   |   |   |   |       |   |-- shell
|       |   |   |   |   |       |   |   `-- shell.js
|       |   |   |   |   |       |   |-- solidity
|       |   |   |   |   |       |   |   `-- solidity.js
|       |   |   |   |   |       |   |-- sophia
|       |   |   |   |   |       |   |   `-- sophia.js
|       |   |   |   |   |       |   |-- sparql
|       |   |   |   |   |       |   |   `-- sparql.js
|       |   |   |   |   |       |   |-- sql
|       |   |   |   |   |       |   |   `-- sql.js
|       |   |   |   |   |       |   |-- st
|       |   |   |   |   |       |   |   `-- st.js
|       |   |   |   |   |       |   |-- swift
|       |   |   |   |   |       |   |   `-- swift.js
|       |   |   |   |   |       |   |-- systemverilog
|       |   |   |   |   |       |   |   `-- systemverilog.js
|       |   |   |   |   |       |   |-- tcl
|       |   |   |   |   |       |   |   `-- tcl.js
|       |   |   |   |   |       |   |-- twig
|       |   |   |   |   |       |   |   `-- twig.js
|       |   |   |   |   |       |   |-- typescript
|       |   |   |   |   |       |   |   `-- typescript.js
|       |   |   |   |   |       |   |-- vb
|       |   |   |   |   |       |   |   `-- vb.js
|       |   |   |   |   |       |   |-- wgsl
|       |   |   |   |   |       |   |   `-- wgsl.js
|       |   |   |   |   |       |   |-- xml
|       |   |   |   |   |       |   |   `-- xml.js
|       |   |   |   |   |       |   `-- yaml
|       |   |   |   |   |       |       `-- yaml.js
|       |   |   |   |   |       |-- editor
|       |   |   |   |   |       |   |-- editor.main.css
|       |   |   |   |   |       |   |-- editor.main.js
|       |   |   |   |   |       |   |-- editor.main.nls.de.js
|       |   |   |   |   |       |   |-- editor.main.nls.es.js
|       |   |   |   |   |       |   |-- editor.main.nls.fr.js
|       |   |   |   |   |       |   |-- editor.main.nls.it.js
|       |   |   |   |   |       |   |-- editor.main.nls.ja.js
|       |   |   |   |   |       |   |-- editor.main.nls.js
|       |   |   |   |   |       |   |-- editor.main.nls.ko.js
|       |   |   |   |   |       |   |-- editor.main.nls.ru.js
|       |   |   |   |   |       |   |-- editor.main.nls.zh-cn.js
|       |   |   |   |   |       |   `-- editor.main.nls.zh-tw.js
|       |   |   |   |   |       |-- language
|       |   |   |   |   |       |   |-- css
|       |   |   |   |   |       |   |   |-- cssMode.js
|       |   |   |   |   |       |   |   `-- cssWorker.js
|       |   |   |   |   |       |   |-- html
|       |   |   |   |   |       |   |   |-- htmlMode.js
|       |   |   |   |   |       |   |   `-- htmlWorker.js
|       |   |   |   |   |       |   |-- json
|       |   |   |   |   |       |   |   |-- jsonMode.js
|       |   |   |   |   |       |   |   `-- jsonWorker.js
|       |   |   |   |   |       |   `-- typescript
|       |   |   |   |   |       |       |-- tsMode.js
|       |   |   |   |   |       |       `-- tsWorker.js
|       |   |   |   |   |       `-- loader.js
|       |   |   |   |   `-- index.html
|       |   |   |   |-- WinUiGallery
|       |   |   |   |   |-- Button.png
|       |   |   |   |   |-- Flyout.png
|       |   |   |   |   |-- LICENSE
|       |   |   |   |   `-- MenuBar.png
|       |   |   |   |-- geo_icons.png
|       |   |   |   |-- octonaut.jpg
|       |   |   |   |-- pexels-johannes-plenio-1103970.jpg
|       |   |   |   |-- wpfui.png
|       |   |   |   `-- wpfui_full.png
|       |   |   |-- CodeSamples
|       |   |   |   `-- Typography
|       |   |   |       `-- TypographySample_xaml.txt
|       |   |   |-- Controllers
|       |   |   |   `-- MonacoController.cs
|       |   |   |-- Controls
|       |   |   |   |-- ControlExample.xaml
|       |   |   |   |-- ControlExample.xaml.cs
|       |   |   |   |-- GalleryNavigationPresenter.xaml
|       |   |   |   |-- GalleryNavigationPresenter.xaml.cs
|       |   |   |   |-- PageControlDocumentation.xaml
|       |   |   |   |-- PageControlDocumentation.xaml.cs
|       |   |   |   |-- TermsOfUseContentDialog.xaml
|       |   |   |   |-- TermsOfUseContentDialog.xaml.cs
|       |   |   |   |-- TypographyControl.xaml
|       |   |   |   `-- TypographyControl.xaml.cs
|       |   |   |-- ControlsLookup
|       |   |   |   |-- ControlPages.cs
|       |   |   |   |-- GalleryPage.cs
|       |   |   |   `-- GalleryPageAttribute.cs
|       |   |   |-- DependencyModel
|       |   |   |   `-- ServiceCollectionExtensions.cs
|       |   |   |-- Effects
|       |   |   |   |-- Snowflake.cs
|       |   |   |   `-- SnowflakeEffect.cs
|       |   |   |-- Helpers
|       |   |   |   |-- EnumToBooleanConverter.cs
|       |   |   |   |-- NameToPageTypeConverter.cs
|       |   |   |   |-- NullToVisibilityConverter.cs
|       |   |   |   |-- PaneDisplayModeToIndexConverter.cs
|       |   |   |   `-- ThemeToIndexConverter.cs
|       |   |   |-- Models
|       |   |   |   |-- Monaco
|       |   |   |   |   |-- MonacoLanguage.cs
|       |   |   |   |   `-- MonacoTheme.cs
|       |   |   |   |-- DisplayableIcon.cs
|       |   |   |   |-- Folder.cs
|       |   |   |   |-- NavigationCard.cs
|       |   |   |   |-- Person.cs
|       |   |   |   |-- Product.cs
|       |   |   |   |-- Unit.cs
|       |   |   |   `-- WindowCard.cs
|       |   |   |-- Resources
|       |   |   |   |-- Translations.cs
|       |   |   |   |-- Translations.pl-PL.Designer.cs
|       |   |   |   `-- Translations.pl-PL.resx
|       |   |   |-- Services
|       |   |   |   |-- Contracts
|       |   |   |   |   `-- IWindow.cs
|       |   |   |   |-- ApplicationHostService.cs
|       |   |   |   `-- WindowsProviderService.cs
|       |   |   |-- ViewModels
|       |   |   |   |-- Pages
|       |   |   |   |   |-- BasicInput
|       |   |   |   |   |   |-- AnchorViewModel.cs
|       |   |   |   |   |   |-- BasicInputViewModel.cs
|       |   |   |   |   |   |-- ButtonViewModel.cs
|       |   |   |   |   |   |-- CheckBoxViewModel.cs
|       |   |   |   |   |   |-- ComboBoxViewModel.cs
|       |   |   |   |   |   |-- DropDownButtonViewModel.cs
|       |   |   |   |   |   |-- HyperlinkButtonViewModel.cs
|       |   |   |   |   |   |-- RadioButtonViewModel.cs
|       |   |   |   |   |   |-- RatingViewModel.cs
|       |   |   |   |   |   |-- SliderViewModel.cs
|       |   |   |   |   |   |-- SplitButtonViewModel.cs
|       |   |   |   |   |   |-- ThumbRateViewModel.cs
|       |   |   |   |   |   |-- ToggleButtonViewModel.cs
|       |   |   |   |   |   `-- ToggleSwitchViewModel.cs
|       |   |   |   |   |-- Collections
|       |   |   |   |   |   |-- CollectionsViewModel.cs
|       |   |   |   |   |   |-- DataGridViewModel.cs
|       |   |   |   |   |   |-- ListBoxViewModel.cs
|       |   |   |   |   |   |-- ListViewViewModel.cs
|       |   |   |   |   |   |-- TreeListViewModel.cs
|       |   |   |   |   |   `-- TreeViewViewModel.cs
|       |   |   |   |   |-- DateAndTime
|       |   |   |   |   |   |-- CalendarDatePickerViewModel.cs
|       |   |   |   |   |   |-- CalendarViewModel.cs
|       |   |   |   |   |   |-- DateAndTimeViewModel.cs
|       |   |   |   |   |   |-- DatePickerViewModel.cs
|       |   |   |   |   |   `-- TimePickerViewModel.cs
|       |   |   |   |   |-- DesignGuidance
|       |   |   |   |   |   |-- ColorsViewModel.cs
|       |   |   |   |   |   |-- IconsViewModel.cs
|       |   |   |   |   |   `-- TypographyViewModel.cs
|       |   |   |   |   |-- DialogsAndFlyouts
|       |   |   |   |   |   |-- ContentDialogViewModel.cs
|       |   |   |   |   |   |-- DialogsAndFlyoutsViewModel.cs
|       |   |   |   |   |   |-- FlyoutViewModel.cs
|       |   |   |   |   |   |-- MessageBoxViewModel.cs
|       |   |   |   |   |   `-- SnackbarViewModel.cs
|       |   |   |   |   |-- Layout
|       |   |   |   |   |   |-- CardActionViewModel.cs
|       |   |   |   |   |   |-- CardControlViewModel.cs
|       |   |   |   |   |   |-- ExpanderViewModel.cs
|       |   |   |   |   |   `-- LayoutViewModel.cs
|       |   |   |   |   |-- Media
|       |   |   |   |   |   |-- CanvasViewModel.cs
|       |   |   |   |   |   |-- ImageViewModel.cs
|       |   |   |   |   |   |-- MediaViewModel.cs
|       |   |   |   |   |   |-- WebBrowserViewModel.cs
|       |   |   |   |   |   `-- WebViewViewModel.cs
|       |   |   |   |   |-- Navigation
|       |   |   |   |   |   |-- BreadcrumbBarViewModel.cs
|       |   |   |   |   |   |-- MenuViewModel.cs
|       |   |   |   |   |   |-- MultilevelNavigationSample.cs
|       |   |   |   |   |   |-- NavigationViewModel.cs
|       |   |   |   |   |   |-- NavigationViewViewModel.cs
|       |   |   |   |   |   |-- TabControlViewModel.cs
|       |   |   |   |   |   `-- TabViewViewModel.cs
|       |   |   |   |   |-- OpSystem
|       |   |   |   |   |   |-- ClipboardViewModel.cs
|       |   |   |   |   |   |-- FilePickerViewModel.cs
|       |   |   |   |   |   `-- OpSystemViewModel.cs
|       |   |   |   |   |-- StatusAndInfo
|       |   |   |   |   |   |-- InfoBadgeViewModel.cs
|       |   |   |   |   |   |-- InfoBarViewModel.cs
|       |   |   |   |   |   |-- ProgressBarViewModel.cs
|       |   |   |   |   |   |-- ProgressRingViewModel.cs
|       |   |   |   |   |   |-- StatusAndInfoViewModel.cs
|       |   |   |   |   |   `-- ToolTipViewModel.cs
|       |   |   |   |   |-- Text
|       |   |   |   |   |   |-- AutoSuggestBoxViewModel.cs
|       |   |   |   |   |   |-- LabelViewModel.cs
|       |   |   |   |   |   |-- NumberBoxViewModel.cs
|       |   |   |   |   |   |-- PasswordBoxViewModel.cs
|       |   |   |   |   |   |-- RichTextBoxViewModel.cs
|       |   |   |   |   |   |-- TextBlockViewModel.cs
|       |   |   |   |   |   |-- TextBoxViewModel.cs
|       |   |   |   |   |   `-- TextViewModel.cs
|       |   |   |   |   |-- Windows
|       |   |   |   |   |   `-- WindowsViewModel.cs
|       |   |   |   |   |-- AllControlsViewModel.cs
|       |   |   |   |   |-- DashboardViewModel.cs
|       |   |   |   |   `-- SettingsViewModel.cs
|       |   |   |   |-- Windows
|       |   |   |   |   |-- EditorWindowViewModel.cs
|       |   |   |   |   |-- MainWindowViewModel.cs
|       |   |   |   |   |-- MonacoWindowViewModel.cs
|       |   |   |   |   `-- SandboxWindowViewModel.cs
|       |   |   |   `-- ViewModel.cs
|       |   |   |-- Views
|       |   |   |   |-- Pages
|       |   |   |   |   |-- BasicInput
|       |   |   |   |   |   |-- AnchorPage.xaml
|       |   |   |   |   |   |-- AnchorPage.xaml.cs
|       |   |   |   |   |   |-- BasicInputPage.xaml
|       |   |   |   |   |   |-- BasicInputPage.xaml.cs
|       |   |   |   |   |   |-- ButtonPage.xaml
|       |   |   |   |   |   |-- ButtonPage.xaml.cs
|       |   |   |   |   |   |-- CheckBoxPage.xaml
|       |   |   |   |   |   |-- CheckBoxPage.xaml.cs
|       |   |   |   |   |   |-- ComboBoxPage.xaml
|       |   |   |   |   |   |-- ComboBoxPage.xaml.cs
|       |   |   |   |   |   |-- DropDownButtonPage.xaml
|       |   |   |   |   |   |-- DropDownButtonPage.xaml.cs
|       |   |   |   |   |   |-- HyperlinkButtonPage.xaml
|       |   |   |   |   |   |-- HyperlinkButtonPage.xaml.cs
|       |   |   |   |   |   |-- RadioButtonPage.xaml
|       |   |   |   |   |   |-- RadioButtonPage.xaml.cs
|       |   |   |   |   |   |-- RatingPage.xaml
|       |   |   |   |   |   |-- RatingPage.xaml.cs
|       |   |   |   |   |   |-- SliderPage.xaml
|       |   |   |   |   |   |-- SliderPage.xaml.cs
|       |   |   |   |   |   |-- SplitButtonPage.xaml
|       |   |   |   |   |   |-- SplitButtonPage.xaml.cs
|       |   |   |   |   |   |-- ThumbRatePage.xaml
|       |   |   |   |   |   |-- ThumbRatePage.xaml.cs
|       |   |   |   |   |   |-- ToggleButtonPage.xaml
|       |   |   |   |   |   |-- ToggleButtonPage.xaml.cs
|       |   |   |   |   |   |-- ToggleSwitchPage.xaml
|       |   |   |   |   |   `-- ToggleSwitchPage.xaml.cs
|       |   |   |   |   |-- Collections
|       |   |   |   |   |   |-- CollectionsPage.xaml
|       |   |   |   |   |   |-- CollectionsPage.xaml.cs
|       |   |   |   |   |   |-- DataGridPage.xaml
|       |   |   |   |   |   |-- DataGridPage.xaml.cs
|       |   |   |   |   |   |-- ListBoxPage.xaml
|       |   |   |   |   |   |-- ListBoxPage.xaml.cs
|       |   |   |   |   |   |-- ListViewPage.xaml
|       |   |   |   |   |   |-- ListViewPage.xaml.cs
|       |   |   |   |   |   |-- TreeListPage.xaml
|       |   |   |   |   |   |-- TreeListPage.xaml.cs
|       |   |   |   |   |   |-- TreeViewPage.xaml
|       |   |   |   |   |   `-- TreeViewPage.xaml.cs
|       |   |   |   |   |-- DateAndTime
|       |   |   |   |   |   |-- CalendarDatePickerPage.xaml
|       |   |   |   |   |   |-- CalendarDatePickerPage.xaml.cs
|       |   |   |   |   |   |-- CalendarPage.xaml
|       |   |   |   |   |   |-- CalendarPage.xaml.cs
|       |   |   |   |   |   |-- DateAndTimePage.xaml
|       |   |   |   |   |   |-- DateAndTimePage.xaml.cs
|       |   |   |   |   |   |-- DatePickerPage.xaml
|       |   |   |   |   |   |-- DatePickerPage.xaml.cs
|       |   |   |   |   |   |-- TimePickerPage.xaml
|       |   |   |   |   |   `-- TimePickerPage.xaml.cs
|       |   |   |   |   |-- DesignGuidance
|       |   |   |   |   |   |-- ColorsPage.xaml
|       |   |   |   |   |   |-- ColorsPage.xaml.cs
|       |   |   |   |   |   |-- IconsPage.xaml
|       |   |   |   |   |   |-- IconsPage.xaml.cs
|       |   |   |   |   |   |-- TypographyPage.xaml
|       |   |   |   |   |   `-- TypographyPage.xaml.cs
|       |   |   |   |   |-- DialogsAndFlyouts
|       |   |   |   |   |   |-- ContentDialogPage.xaml
|       |   |   |   |   |   |-- ContentDialogPage.xaml.cs
|       |   |   |   |   |   |-- DialogsAndFlyoutsPage.xaml
|       |   |   |   |   |   |-- DialogsAndFlyoutsPage.xaml.cs
|       |   |   |   |   |   |-- FlyoutPage.xaml
|       |   |   |   |   |   |-- FlyoutPage.xaml.cs
|       |   |   |   |   |   |-- MessageBoxPage.xaml
|       |   |   |   |   |   |-- MessageBoxPage.xaml.cs
|       |   |   |   |   |   |-- SnackbarPage.xaml
|       |   |   |   |   |   `-- SnackbarPage.xaml.cs
|       |   |   |   |   |-- Layout
|       |   |   |   |   |   |-- CardActionPage.xaml
|       |   |   |   |   |   |-- CardActionPage.xaml.cs
|       |   |   |   |   |   |-- CardControlPage.xaml
|       |   |   |   |   |   |-- CardControlPage.xaml.cs
|       |   |   |   |   |   |-- ExpanderPage.xaml
|       |   |   |   |   |   |-- ExpanderPage.xaml.cs
|       |   |   |   |   |   |-- LayoutPage.xaml
|       |   |   |   |   |   `-- LayoutPage.xaml.cs
|       |   |   |   |   |-- Media
|       |   |   |   |   |   |-- CanvasPage.xaml
|       |   |   |   |   |   |-- CanvasPage.xaml.cs
|       |   |   |   |   |   |-- ImagePage.xaml
|       |   |   |   |   |   |-- ImagePage.xaml.cs
|       |   |   |   |   |   |-- MediaPage.xaml
|       |   |   |   |   |   |-- MediaPage.xaml.cs
|       |   |   |   |   |   |-- WebBrowserPage.xaml
|       |   |   |   |   |   |-- WebBrowserPage.xaml.cs
|       |   |   |   |   |   |-- WebViewPage.xaml
|       |   |   |   |   |   `-- WebViewPage.xaml.cs
|       |   |   |   |   |-- Navigation
|       |   |   |   |   |   |-- BreadcrumbBarPage.xaml
|       |   |   |   |   |   |-- BreadcrumbBarPage.xaml.cs
|       |   |   |   |   |   |-- MenuPage.xaml
|       |   |   |   |   |   |-- MenuPage.xaml.cs
|       |   |   |   |   |   |-- MultilevelNavigationPage.xaml
|       |   |   |   |   |   |-- MultilevelNavigationPage.xaml.cs
|       |   |   |   |   |   |-- NavigationPage.xaml
|       |   |   |   |   |   |-- NavigationPage.xaml.cs
|       |   |   |   |   |   |-- NavigationViewPage.xaml
|       |   |   |   |   |   |-- NavigationViewPage.xaml.cs
|       |   |   |   |   |   |-- TabControlPage.xaml
|       |   |   |   |   |   |-- TabControlPage.xaml.cs
|       |   |   |   |   |   |-- TabViewPage.xaml
|       |   |   |   |   |   `-- TabViewPage.xaml.cs
|       |   |   |   |   |-- OpSystem
|       |   |   |   |   |   |-- ClipboardPage.xaml
|       |   |   |   |   |   |-- ClipboardPage.xaml.cs
|       |   |   |   |   |   |-- FilePickerPage.xaml
|       |   |   |   |   |   |-- FilePickerPage.xaml.cs
|       |   |   |   |   |   |-- OpSystemPage.xaml
|       |   |   |   |   |   `-- OpSystemPage.xaml.cs
|       |   |   |   |   |-- Samples
|       |   |   |   |   |   |-- MultilevelNavigationSamplePage1.xaml
|       |   |   |   |   |   |-- MultilevelNavigationSamplePage1.xaml.cs
|       |   |   |   |   |   |-- MultilevelNavigationSamplePage2.xaml
|       |   |   |   |   |   |-- MultilevelNavigationSamplePage2.xaml.cs
|       |   |   |   |   |   |-- MultilevelNavigationSamplePage3.xaml
|       |   |   |   |   |   |-- MultilevelNavigationSamplePage3.xaml.cs
|       |   |   |   |   |   |-- SamplePage1.xaml
|       |   |   |   |   |   |-- SamplePage1.xaml.cs
|       |   |   |   |   |   |-- SamplePage2.xaml
|       |   |   |   |   |   |-- SamplePage2.xaml.cs
|       |   |   |   |   |   |-- SamplePage3.xaml
|       |   |   |   |   |   `-- SamplePage3.xaml.cs
|       |   |   |   |   |-- StatusAndInfo
|       |   |   |   |   |   |-- InfoBadgePage.xaml
|       |   |   |   |   |   |-- InfoBadgePage.xaml.cs
|       |   |   |   |   |   |-- InfoBarPage.xaml
|       |   |   |   |   |   |-- InfoBarPage.xaml.cs
|       |   |   |   |   |   |-- ProgressBarPage.xaml
|       |   |   |   |   |   |-- ProgressBarPage.xaml.cs
|       |   |   |   |   |   |-- ProgressRingPage.xaml
|       |   |   |   |   |   |-- ProgressRingPage.xaml.cs
|       |   |   |   |   |   |-- StatusAndInfoPage.xaml
|       |   |   |   |   |   |-- StatusAndInfoPage.xaml.cs
|       |   |   |   |   |   |-- ToolTipPage.xaml
|       |   |   |   |   |   `-- ToolTipPage.xaml.cs
|       |   |   |   |   |-- Text
|       |   |   |   |   |   |-- AutoSuggestBoxPage.xaml
|       |   |   |   |   |   |-- AutoSuggestBoxPage.xaml.cs
|       |   |   |   |   |   |-- LabelPage.xaml
|       |   |   |   |   |   |-- LabelPage.xaml.cs
|       |   |   |   |   |   |-- NumberBoxPage.xaml
|       |   |   |   |   |   |-- NumberBoxPage.xaml.cs
|       |   |   |   |   |   |-- PasswordBoxPage.xaml
|       |   |   |   |   |   |-- PasswordBoxPage.xaml.cs
|       |   |   |   |   |   |-- RichTextBoxPage.xaml
|       |   |   |   |   |   |-- RichTextBoxPage.xaml.cs
|       |   |   |   |   |   |-- TextBlockPage.xaml
|       |   |   |   |   |   |-- TextBlockPage.xaml.cs
|       |   |   |   |   |   |-- TextBoxPage.xaml
|       |   |   |   |   |   |-- TextBoxPage.xaml.cs
|       |   |   |   |   |   |-- TextPage.xaml
|       |   |   |   |   |   `-- TextPage.xaml.cs
|       |   |   |   |   |-- Windows
|       |   |   |   |   |   |-- WindowsPage.xaml
|       |   |   |   |   |   `-- WindowsPage.xaml.cs
|       |   |   |   |   |-- AllControlsPage.xaml
|       |   |   |   |   |-- AllControlsPage.xaml.cs
|       |   |   |   |   |-- DashboardPage.xaml
|       |   |   |   |   |-- DashboardPage.xaml.cs
|       |   |   |   |   |-- SettingsPage.xaml
|       |   |   |   |   `-- SettingsPage.xaml.cs
|       |   |   |   `-- Windows
|       |   |   |       |-- EditorWindow.xaml
|       |   |   |       |-- EditorWindow.xaml.cs
|       |   |   |       |-- MainWindow.xaml
|       |   |   |       |-- MainWindow.xaml.cs
|       |   |   |       |-- MonacoWindow.xaml
|       |   |   |       |-- MonacoWindow.xaml.cs
|       |   |   |       |-- SandboxWindow.xaml
|       |   |   |       `-- SandboxWindow.xaml.cs
|       |   |   |-- app.manifest
|       |   |   |-- App.xaml
|       |   |   |-- App.xaml.cs
|       |   |   |-- AssemblyInfo.cs
|       |   |   |-- GalleryAssembly.cs
|       |   |   |-- GlobalUsings.cs
|       |   |   |-- License - Images.txt
|       |   |   |-- License - Monaco.txt
|       |   |   |-- Wpf.Ui.Gallery.csproj
|       |   |   `-- wpfui.ico
|       |   |-- Wpf.Ui.Gallery.Package
|       |   |   |-- Images
|       |   |   |   |-- LargeTile.scale-100.png
|       |   |   |   |-- LargeTile.scale-125.png
|       |   |   |   |-- LargeTile.scale-150.png
|       |   |   |   |-- LargeTile.scale-200.png
|       |   |   |   |-- LargeTile.scale-400.png
|       |   |   |   |-- LockScreenLogo.scale-200.png
|       |   |   |   |-- SmallTile.scale-100.png
|       |   |   |   |-- SmallTile.scale-125.png
|       |   |   |   |-- SmallTile.scale-150.png
|       |   |   |   |-- SmallTile.scale-200.png
|       |   |   |   |-- SmallTile.scale-400.png
|       |   |   |   |-- SplashScreen.scale-100.png
|       |   |   |   |-- SplashScreen.scale-125.png
|       |   |   |   |-- SplashScreen.scale-150.png
|       |   |   |   |-- SplashScreen.scale-200.png
|       |   |   |   |-- SplashScreen.scale-400.png
|       |   |   |   |-- Square150x150Logo.scale-100.png
|       |   |   |   |-- Square150x150Logo.scale-125.png
|       |   |   |   |-- Square150x150Logo.scale-150.png
|       |   |   |   |-- Square150x150Logo.scale-200.png
|       |   |   |   |-- Square150x150Logo.scale-400.png
|       |   |   |   |-- Square44x44Logo.altform-lightunplated_targetsize-16.png
|       |   |   |   |-- Square44x44Logo.altform-lightunplated_targetsize-24.png
|       |   |   |   |-- Square44x44Logo.altform-lightunplated_targetsize-256.png
|       |   |   |   |-- Square44x44Logo.altform-lightunplated_targetsize-32.png
|       |   |   |   |-- Square44x44Logo.altform-lightunplated_targetsize-48.png
|       |   |   |   |-- Square44x44Logo.altform-unplated_targetsize-16.png
|       |   |   |   |-- Square44x44Logo.altform-unplated_targetsize-256.png
|       |   |   |   |-- Square44x44Logo.altform-unplated_targetsize-32.png
|       |   |   |   |-- Square44x44Logo.altform-unplated_targetsize-48.png
|       |   |   |   |-- Square44x44Logo.scale-100.png
|       |   |   |   |-- Square44x44Logo.scale-125.png
|       |   |   |   |-- Square44x44Logo.scale-150.png
|       |   |   |   |-- Square44x44Logo.scale-200.png
|       |   |   |   |-- Square44x44Logo.scale-400.png
|       |   |   |   |-- Square44x44Logo.targetsize-16.png
|       |   |   |   |-- Square44x44Logo.targetsize-24.png
|       |   |   |   |-- Square44x44Logo.targetsize-24_altform-unplated.png
|       |   |   |   |-- Square44x44Logo.targetsize-256.png
|       |   |   |   |-- Square44x44Logo.targetsize-32.png
|       |   |   |   |-- Square44x44Logo.targetsize-48.png
|       |   |   |   |-- StoreLogo.backup.png
|       |   |   |   |-- StoreLogo.scale-100.png
|       |   |   |   |-- StoreLogo.scale-125.png
|       |   |   |   |-- StoreLogo.scale-150.png
|       |   |   |   |-- StoreLogo.scale-200.png
|       |   |   |   |-- StoreLogo.scale-400.png
|       |   |   |   |-- Wide310x150Logo.scale-100.png
|       |   |   |   |-- Wide310x150Logo.scale-125.png
|       |   |   |   |-- Wide310x150Logo.scale-150.png
|       |   |   |   |-- Wide310x150Logo.scale-200.png
|       |   |   |   `-- Wide310x150Logo.scale-400.png
|       |   |   |-- Package.appxmanifest
|       |   |   `-- Wpf.Ui.Gallery.Package.wapproj
|       |   |-- Wpf.Ui.SyntaxHighlight
|       |   |   |-- Controls
|       |   |   |   |-- CodeBlock.bmp
|       |   |   |   |-- CodeBlock.cs
|       |   |   |   `-- CodeBlock.xaml
|       |   |   |-- Fonts
|       |   |   |   `-- FiraCode-Regular.ttf
|       |   |   |-- Markup
|       |   |   |   `-- SyntaxHighlightDictionary.cs
|       |   |   |-- Properties
|       |   |   |   `-- AssemblyInfo.cs
|       |   |   |-- Highlighter.cs
|       |   |   |-- License - Fira Code.txt
|       |   |   |-- SyntaxHighlight.xaml
|       |   |   |-- SyntaxLanguage.cs
|       |   |   `-- Wpf.Ui.SyntaxHighlight.csproj
|       |   |-- Wpf.Ui.ToastNotifications
|       |   |   |-- Properties
|       |   |   |   `-- AssemblyInfo.cs
|       |   |   |-- GlobalUsings.cs
|       |   |   |-- Toast.cs
|       |   |   |-- VisualStudioToolsManifest.xml
|       |   |   `-- Wpf.Ui.ToastNotifications.csproj
|       |   `-- Wpf.Ui.Tray
|       |       |-- Controls
|       |       |   |-- NotifyIcon.bmp
|       |       |   `-- NotifyIcon.cs
|       |       |-- Internal
|       |       |   `-- InternalNotifyIconManager.cs
|       |       |-- Interop
|       |       |   |-- Libraries.cs
|       |       |   |-- Shell32.cs
|       |       |   `-- User32.cs
|       |       |-- Properties
|       |       |   `-- AssemblyInfo.cs
|       |       |-- GlobalUsings.cs
|       |       |-- Hicon.cs
|       |       |-- INotifyIcon.cs
|       |       |-- INotifyIconService.cs
|       |       |-- NotifyIconEventHandler.cs
|       |       |-- NotifyIconService.cs
|       |       |-- RoutedNotifyIconEvent.cs
|       |       |-- TrayData.cs
|       |       |-- TrayHandler.cs
|       |       |-- TrayManager.cs
|       |       |-- VisualStudioToolsManifest.xml
|       |       `-- Wpf.Ui.Tray.csproj
|       |-- tests
|       |   |-- Wpf.Ui.Gallery.IntegrationTests
|       |   |   |-- Fixtures
|       |   |   |   |-- TestedApplication.cs
|       |   |   |   `-- UiTest.cs
|       |   |   |-- GlobalUsings.cs
|       |   |   |-- NavigationTests.cs
|       |   |   |-- TitleBarTests.cs
|       |   |   |-- WindowTests.cs
|       |   |   |-- Wpf.Ui.Gallery.IntegrationTests.csproj
|       |   |   `-- xunit.runner.json
|       |   `-- Wpf.Ui.UnitTests
|       |       |-- Animations
|       |       |   `-- TransitionAnimationProviderTests.cs
|       |       |-- Extensions
|       |       |   `-- SymbolExtensionsTests.cs
|       |       |-- Usings.cs
|       |       `-- Wpf.Ui.UnitTests.csproj
|       |-- .csharpierignore
|       |-- .csharpierrc
|       |-- .editorconfig
|       |-- .gitattributes
|       |-- .gitignore
|       |-- .vsconfig
|       |-- build.cmd
|       |-- build.ps1
|       |-- CNAME
|       |-- CODE_OF_CONDUCT.md
|       |-- CODEOWNERS
|       |-- CONTRIBUTING.md
|       |-- Directory.Build.props
|       |-- Directory.Build.targets
|       |-- Directory.Packages.props
|       |-- LICENSE
|       |-- LICENSE.md
|       |-- nuget.config
|       |-- README.md
|       |-- SECURITY.md
|       |-- Settings.XamlStyler
|       |-- ThirdPartyNotices.txt
|       |-- Wpf.Ui.Gallery.slnf
|       |-- Wpf.Ui.Library.slnf
|       |-- Wpf.Ui.sln
|       `-- wpfui-main.sln
|-- .gitignore
|-- AGENTS.md
|-- MicroEng_Navisworks.sln
|-- MICROENG_THEME_GUIDE.md
`-- README.md
```


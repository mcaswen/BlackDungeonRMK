## Tooling: UAsset Json Exporter

Plugin: `Plugins/UAssetJsonExporter` (Editor-only)

### Available Commandlets

| Commandlet | Run Name | Description |
|---|---|---|
| BlueprintEdGraphExportCommandlet | `BlueprintEdGraphExport` | Blueprint graphs, nodes, pins, connections |
| AnimMontageExportCommandlet | `AnimMontageExport` | Montage sections, slots, ANS/AN placement and parameters |
| WidgetLayoutExportCommandlet | `WidgetLayoutExport` | Widget tree, layout, animations, EdGraph |
| DataAssetExportCommandlet | `DataAssetExport` | DataAsset subclass properties |
| DataTableExportCommandlet | `DataTableExport` | DataTable row struct and all row data |
| NiagaraSystemExportCommandlet | `NiagaraSystemExport` | Niagara emitters, scripts, renderers |
| MaterialExportCommandlet | `MaterialExport` | Material expressions and connections; MI parameter overrides |
| BehaviorTreeExportCommandlet | `BehaviorTreeExport` | BT tree structure, node parameters, Blackboard keys |
| AnimBlueprintExportCommandlet | `AnimBlueprintExport` | AnimBP EdGraph, StateMachines (states, transitions, blend settings) |
| LevelExportCommandlet | `LevelExport` | Level (.umap) actors / components, delta-from-archetype properties, collision / static mesh / ISM summary, streaming levels |

### Usage

bash Plugins/UAssetJsonExporter/scripts/run_commandlet.sh "<UE_PATH>" "Project.uproject" <RunName> "/Game/Path/Asset"

Wrapper routes automatically by `Saved/UAssetExportQueue/.alive` heartbeat: editor open → in-editor subsystem (Slate toast feedback), editor closed → UnrealEditor-Cmd commandlet. Output format is identical either way.

### Output

`Intermediate/UAssetExport/<AssetPath>.json`

### Reading Strategy

Do NOT read the entire file at once. Instead:
1. Use Grep to locate relevant node titles, function names, or pin connections.
2. Use Read with offset/limit to inspect only the relevant sections.

### When to Use

- A task references a Blueprint/Widget and its internal logic is relevant
- A bug may be in Blueprint wiring rather than C++
- Need to verify variable defaults, component setup, or event flow
- Need to check AnimMontage notify timing, DataAsset configuration, or material setup
- Need to inspect Niagara emitter parameters or DataTable values
- Need to understand BehaviorTree logic flow or AnimBP state machine transitions
- Need to audit a Level: actor placements, static mesh / collision setup, streaming level config, per-instance overrides
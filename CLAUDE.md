# SWEEP SCRIPT — Project Reference

3ds Max **MaxScript** (`.ms`) automation for laying out furniture "packages" and rendering setups.
Reference implementation: [REFERENCE SCRIPT/pkg_automation.ms](REFERENCE%20SCRIPT/pkg_automation.ms).

This file is my (Claude's) working knowledge of the scripts so I don't re-analyze them
every session. When a script changes materially, update this file.

**Scripts in this project:**
- [REFERENCE SCRIPT/pkg_automation.ms](REFERENCE%20SCRIPT/pkg_automation.ms) — original reference (documented below).
- [Sweep Script/sweep_script.ms](Sweep%20Script/sweep_script.ms) — the NEW Sweep Script (see "SWEEP SCRIPT" section at bottom).

---

## What the tool does (one line)
Reads an Excel list of "packages" (each = a set of furniture kits + quantities), clones the
matching scene objects, auto-arranges them into a photographable bedroom/headboard layout,
puts each package on its own hidden layer, and optionally clones a template VRay camera framed
on each package.

## Environment / language notes
- **Language:** MaxScript (3ds Max). Not a general-purpose language — runs inside Max.
- **UI:** built with `rollout` + `dotNetControl` (WinForms `DataGridView`, buttons, fonts).
- **Excel I/O:** via OLE automation (`createOleObject "Excel.Application"`), reads `UsedRange`.
- **Units:** scene units treated as cm. GAP entered in inches, converted `inch * 2.54`.
- **Persistence:** settings saved to `settings\pkg_automation_settings.ini` next to the script
  via `setINISetting`/`getINISetting`.

## Data model (structs)
- `PKGItem (kitName, qty)` — one kit line + instance count.
- `PKGPackage (packageName, items)` — a package name + array of `PKGItem`.

## Excel input format
One package per row, e.g.:
```
PKG026197 (B779B16,1;B779B1,1;B779-46,1;B779-92,1)
```
- Name before `(`; kit entries inside `()` split by `;`.
- Each entry `KitID,Qty` (qty after comma; defaults to 1).
- Series prefix (e.g. `B779`) and a leading `-` are stripped via `stripSeriesPrefix`.
- Parsing: `parsePackageLine` → `parseKitEntry`.

## UI workflow (rollout `pkgRollout`)
1. Enter **SERIES NAME**.
2. Per kit row in the `DataGridView` (`dgvKits`): **KIT ID**, **TYPE**
   (AUTO/BED/HEADBOARD/NIGHTSTAND/CHEST/DRESSER/MIRROR/OTHER), **SELECT OBJECT** (scene node).
3. **+ ADD MORE KIT** / **DELETE KIT** rows; **REFRESH SCENE OBJECTS** rescans scene names.
4. Browse to the **Excel** file.
5. Set **FRONT AXIS** (axis the models' fronts face; most libraries +Y), **BED HEAD**,
   **GAP**, **CHEST x**, and camera options (**CAM FIT %**, **CAM MODE**, **RATIO**).
6. Optional **PICK VRAY CAMERA** template.
7. **RUN** → preview dialog (`pkgPreviewDlg`) lists packages → RUN confirms → `executePending`.

## Layout logic (the heart of the script)
`arrangeBedroom` → `arrangeBedMode`. Camera assumed at **-Y looking +Y**; pieces rotated so
fronts face the camera, then an artistic angle is applied.
- **Bedroom mode** (a BED is mapped): BED centered (auto-detects headboard direction via
  `detectBedHeadAngleDeg`); NIGHTSTANDs to the left; CHESTs then DRESSERs to the right;
  MIRRORs stacked on matching dresser (`stackMirrorOnDresser`); OTHERs in a back row.
- **Headboard mode** (no BED, HEADBOARD mapped): single back row centered on origin.
- **AUTO type**: classify by bounding-box shape (`classifyByBBox`); largest item becomes bed.
- Pieces placed corner-to-corner with gap using rotated-rect corner math
  (`cornersOfRotatedRect`, `cornerMinX`/`cornerMaxX`).
- `sideAngle`: single piece on a side = 35°; multiple = 32.5° closest to anchor, else 35°;
  right side negated. (Spec: paired pieces differ by 2.5° to show equal side in frame.)

## Key helper functions (where to look)
| Concern | Function(s) |
|---|---|
| Scene object dropdown | `collectSceneObjectNames`, `refreshSceneObjects` |
| Kit → node/type lookup | `buildLookup`, `lookupNode`, `lookupType` |
| Excel parse | `readExcelLines`, `parsePackageLine`, `parseKitEntry`, `stripSeriesPrefix` |
| Cloning (keeps hierarchy — fixes "missing models") | `cloneAsInstance`, `collectHierarchy`, `findCloneRoot` |
| Bounding box (group-aware) | `groupUnionBBox`, `bboxVolume`, `getLocalSize` |
| Placement | `placeAndRotate` (resets rotation, centers, rotates Z, drops to floor on bbox min Z) |
| Rotation offsets | `frontAxisOffsetDeg`, `bedHeadOffsetDeg`, `detectBedHeadAngleDeg` |
| Camera | `cloneCameraToLayer`, `cameraHFOVDeg` (VRay/Autodesk/standard) |
| Layers | `safeCreateHiddenLayer` (one hidden layer per package; ALL nodes added so hide works) |
| Build/run | `buildPackage`, `executePending`, `getPendingPackages` |
| Settings | `getSettingsPath`, `saveSettings`, `loadSettings` |

## Camera framing (`cloneCameraToLayer`)
Clones template camera+target (`#copy` so all props carry). Locks Cam X / Target X / Target Y
to package center; Cam Z = `100 + kitCount*20`; distance from **CAM MODE**:
- `AUTO+CAP` (default) = min(auto-fit, scaled template) · `AUTO` = pure auto-fit ·
  `TEMPLATE` = template distance · `RATIO` = `pkgW * spnCamRatio`.
- **CAM FIT %** scales margin (100 = edge-to-edge, lower = more padding).

## Gotchas / non-obvious things
- **Enter-in-cell bug fix:** `editorKeyDown` swallows Enter inside grid cells so it doesn't
  trigger RUN and lose the typed value. RUN also calls `dgvKits.EndEdit()` before reading.
- **Cloning whole hierarchy** (`collectHierarchy` + re-parent) is deliberate — cloning only a
  group head drops members ("missing models" bug).
- **Floor alignment:** `placeAndRotate` aligns XY to bbox center but Z to bbox **min** so
  pieces sit on the floor instead of sinking halfway.
- All cloned nodes (not just roots) are added to the layer, else hiding the layer leaves
  members visible (each Max node carries its own layer membership).
- Excel reads are wrapped in try/catch with a "kill Excel.exe in Task Manager" hint on failure.

---

# SWEEP SCRIPT — [Sweep Script/sweep_script.ms](Sweep%20Script/sweep_script.ms)

A separate, simpler automation built from the pkg patterns. Reuses the same proven helpers
(`DataGridView` UI, Excel OLE read, `cloneAsInstance`/`collectHierarchy`/`findCloneRoot`,
`groupUnionBBox`, hidden-layer-per-item, INI settings, Enter-in-cell fix, preview dialog).

## What it does (one line)
Reads an Excel column of full names, and for each name clones the mapped model(s), lays them
out along +X touching (no gap) in name order, groups them, and rotates the whole group on Z by
an angle implied by the name's view-type. Each name → its own hidden layer + group.

## Name format
`SERIES - MODEL(s) - VIEWTYPE - [SW] - AGR`, e.g. `87214-10-17-08-ANGLE-CLSD-SW-AGR`.
- `87214` = series (stripped via the SERIES NAME field).
- `10-17-08` = one or more numeric model tokens, mapped to scene objects in the grid.
- `ANGLE-CLSD` = view-type (+ optional CLSD/OPEN/CONTROL variant).
- `SW`, `AGR` = ignored suffix tokens. Some rows are just `-AGR` (no `SW`) — handled.
- Parsing: `parseSweepName` (struct `SweepName(fullName, nums, variant, rot, viewStr)`).

## Rotation table (whole-group Z, by base keyword — test BACK-ANGLE before ANGLE)
| View type | Z rot |
|---|---|
| ANGLE (incl. ANGLE-CLSD/OPEN) | 35° |
| BACK-ANGLE | -145° |
| SIDE (incl. SIDE-CLSD/OPEN) | 90° |
| HEAD / HEAD-ON | 0° |
| DETAIL (incl. DETAIL-CONTROL) | 0° |

## Decisions baked in (confirmed with user 2026-06-17)
- **Input** = Excel file of names (not manual entry).
- **Mapping grid** keys on `(MODEL #, VARIANT)` → scene object. Columns: MODEL # / VARIANT
  (DEFAULT/CLSD/OPEN/CONTROL) / SELECT OBJECT. `buildLookup` → `lookupObj` with **DEFAULT
  fallback** when a requested variant isn't mapped for that number.
- **CLSD/OPEN/CONTROL are separate scene objects** (open vs closed geometry), mapped per variant.
- **Arrangement** = world +X, name order, bounding boxes touching, zero gap (`placeAdjacentX`).
  Models keep their own orientation; only the final group is rotated (`about ctr rotate ... Z`).
- No FRONT-AXIS handling and **no camera cloning** (intentionally omitted; pkg has them, sweep doesn't).

## Key functions
`parseSweepName` (name→struct), `buildLookup`/`lookupObj` (mapping + fallback), `numEqual`
(08 vs 8), `isAllDigits` (model-token detection), `placeAdjacentX` (touching X layout),
`buildSweep` (clone→arrange→group→rotate→layer), `executePending`, settings under
`settings\sweep_script_settings.ini`.

## Open assumptions to verify with user if issues arise
- Variant applies per-model-number with DEFAULT fallback (so multi-model names mix variant +
  default objects). If a name's variant should force ALL its models to that variant only, revisit.
- Rotation is about the assembled group's bbox center. If the user wants rotation about world
  origin or a specific pivot, change the `about ctr` center in `buildSweep`.
- Touching means world-axis bbox edges meet along X with each model at its current orientation.

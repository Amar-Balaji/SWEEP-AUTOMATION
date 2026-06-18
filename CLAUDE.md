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
out along +X touching (no gap) in name order with all BACK edges on one line, groups them,
rotates the whole group on Z by an angle implied by the name's view-type, and clones a framed
template camera into the layer. Each name → its own hidden layer + group + camera.

## Name format
`SERIES - MODEL(s) - VIEWTYPE - [SW] - AGR`, e.g. `87214-10-17-08-ANGLE-CLSD-SW-AGR`.
- `87214` = series (stripped via the SERIES NAME field).
- `10-17-08` = one or more numeric model tokens, mapped to scene objects in the grid.
- `ANGLE-CLSD` = view-type. The base keyword drives rotation (see table). `CONTROL` is inert.
- **`OPEN` / `CLSD` words DO drive model selection (added 2026-06-18).** If the name's view words
  contain `OPEN`, each model token `N` is looked up as `N_OPEN` first (e.g. `25` → `25_OPEN`),
  falling back to plain `N` if no `N_OPEN` row is mapped; `CLSD`/`CLOSED` → `N_CLSD` likewise.
  No `OPEN`/`CLSD` word → plain `N`. So `87214-25-SIDE-OPEN-SW-AGR` clones the `25_OPEN` grid row
  (whose VARIANT is typically `SINGLE RECLINER OPEN`), while `87214-25-SIDE-SW-AGR` clones `25`.
  The word does NOT change rotation. Implemented as `modSuffix` in `buildSweep`; relies on the
  tightened `numEqual` (numeric compare only when both tokens are pure digits, so `25` never
  matches a `25_OPEN` row and vice-versa).
- `SW`, `AGR` = ignored suffix tokens. Some rows are just `-AGR` (no `SW`) — handled.
- Parsing: `parseSweepName` (struct `SweepName(fullName, nums, rot, viewStr)`).

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
- **Mapping grid** = MODEL # → ONE scene object. Lookup is by MODEL # only (`lookupPair`).
  Columns: MODEL # / VARIANT / SELECT OBJECT. (An auto-derive-MODEL#-from-object-name variant
  was prototyped and reverted 2026-06-18 — the manual MODEL # column is intentional because the
  open/closed suffix cases, e.g. `8721425_OPEN`, made auto-derivation ambiguous.) **VARIANT is a
  descriptive seating-type label only** (DEFAULT / SINGLE SEATER / DOUBLE SEATER / LOVE SEAT /
  ARM CHAIR / SINGLE RECLINER OPEN / SINGLE RECLINER CLSD / LEFT SIDE SOFA / RIGHT SIDE SOFA /
  OTTOMAN) — it does NOT affect which object is cloned (but LEFT/RIGHT SIDE SOFA drive ordering +
  rotation overrides, and it drives camera height; see below).
- **Arrangement** = world +X, name order, bounding boxes touching, zero gap (`placeAdjacentX`).
  Each piece's BACK edge (+Y / bbox max.y) is aligned to a common line (Y=0); pieces extend
  toward −Y (front). Models keep their own orientation; only the final group is rotated
  (`about ctr rotate ... Z`). After rotation the whole group is moved so its bbox centre sits
  at the world origin (X=0, Y=0).
- **SIDE SOFA rules** (variant-driven; `lookupPair` returns node + variant):
  - `LEFT SIDE SOFA` → forced to the **extreme left** of the +X line (placed first); overall
    group rotation overridden to **−55°**.
  - `RIGHT SIDE SOFA` → forced to the **extreme right** (placed last); overall rotation **+55°**.
  - Order is: left sofa(s) → middle pieces (name order) → right sofa(s).
  - Rotation: `LEFT SIDE SOFA` → −55, `RIGHT SIDE SOFA` → +55. RIGHT is checked **last**, so if a
    right side sofa is anywhere in the list it wins (+55) even alongside a left one. E.g.
    `87214-66-17-ANGLE` where 17 = RIGHT SIDE SOFA → +55°.
- **Camera** = inline `PICK VRAY CAMERA` button only (CAM FIT/MODE/RATIO controls removed). A
  copy of the template camera+target is cloned into EVERY created layer (`cloneCameraToLayer`).
  Placement: **target at world X=0, Y=0**; **camera distance B = A / 0.33** where A = object
  width (X extent), camera pulled straight back on −Y to `[0, -B, camZ]`. **Camera Z and target
  Z are fixed per variant** (`variantCameraHeights`, in cm): OTTOMAN `45/23`; SINGLE SEATER /
  ARM CHAIR / LOVE SEAT / SINGLE RECLINER OPEN / SINGLE RECLINER CLSD `78/41`; everything else
  incl. DEFAULT / DOUBLE SEATER / LEFT-RIGHT SIDE SOFA `107/43` (recliner height = single-seat,
  set 2026-06-18 pending confirmation). The arrangement's height is taken from the **FIRST model's variant**
  (`rootVars[1]`).

- **Settings INI (added 2026-06-18)** = under the Excel row, **SAVE INI FILE** + **LOAD INI
  FILE** buttons. Default auto-save is still `settings\sweep_script_settings.ini`. SAVE picks a
  custom path that becomes `activeIniPath` — the active settings file all further auto-saves
  (close / RUN / add-delete row) write to. LOAD reads a chosen `.ini` back into the UI and makes
  it active. Persistence: a pointer `[main] ActiveIniPath` written into the **default** ini lets
  `bootstrapActiveIni` resume from the user's chosen file after a Max restart. Path resolution:
  `getDefaultSettingsPath` (script-dir default) vs `getSettingsPath` (active if set, else default).

- **SERIES NAME is a dotNet TextBox, not a MaxScript `edittext` (changed 2026-06-18).** A plain
  `edittext` lost its typed buffer on Enter because the dotNet-hosted dialog swallowed the key as
  a default-button "accept" before the field committed (clearing `87214` etc.). Fix: `edtSeries`
  is `dotNetControl "...TextBox"` with a `KeyDown` handler (`seriesKeyDown`) that on Enter sets
  `SuppressKeyPress/Handled`, calls `saveSettings`, and moves focus to REFRESH — same swallow path
  the grid uses (`editorKeyDown`). `frm.AcceptButton = undefined` on open also defangs the accept.

- **Combo cells ignore the mouse wheel (added 2026-06-18).** `EditingControlShowing` attaches
  `comboMouseWheel` (sets `HandledMouseEventArgs.Handled = true`) to the VARIANT / SELECT OBJECT
  edit control so an accidental hover-scroll never changes the selected option. The grid itself
  still scrolls normally when no cell is being edited (the handler lives only on the edit control).

## Key functions
`parseSweepName` (name→struct), `buildLookup`/`lookupPair` (MODEL #→object+variant), `numEqual`
(08 vs 8), `isAllDigits` (model-token detection), `placeAdjacentX` (touching X layout + back-edge
alignment), `buildSweep` (clone→arrange→group→rotate→layer→camera), `cloneCameraToLayer`,
`cameraHFOVDeg`, `executePending`. Settings path: `getDefaultSettingsPath` / `getSettingsPath` /
`bootstrapActiveIni`; default file `settings\sweep_script_settings.ini`, optional user-chosen
file via SAVE/LOAD INI.

## Open assumptions to verify with user if issues arise
- Rotation is about the assembled group's bbox center. If rotation should be about world origin
  or a specific pivot, change the `about ctr` center in `buildSweep`.
- Back-edge = world-axis bbox max.y of each model at its current orientation. If a model's
  geometric "back" isn't its +Y bbox side, its orientation must be fixed in the scene first.

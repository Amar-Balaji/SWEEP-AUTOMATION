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

## Rotation table (whole-group Z, by base keyword — order matters: test the longer/ALT keyword first)
| View type | Z rot |
|---|---|
| ANGLE (incl. ANGLE-CLSD/OPEN) | 35° |
| ANGLE-ALT (added 2026-06-24; like ANGLE, mirrored) | -35° |
| BACK-ANGLE | -145° |
| BACK (added 2026-06-24; like HEAD but 180° around) | -180° |
| SIDE (incl. SIDE-CLSD/OPEN) | 90° |
| SIDE-ALT (added 2026-06-24; like SIDE, mirrored) | -90° |
| HEAD / HEAD-ON | 0° |
| DETAIL (incl. DETAIL-CONTROL) | 0° |
| TOP (overhead, added 2026-06-24) | 0° |

Test order in `parseSweepName` / CSV scoring: `BACK-ANGLE` → `BACK` → `ANGLE-ALT` → `ANGLE` →
`SIDE-ALT` → `SIDE` → `HEAD` → `DETAIL` → `TOP`. The `-ALT` rows reuse the base view's camera
X/Y/Z/target in `camera_coordinates.csv`; only ROT flips sign (ANGLE-ALT -35, SIDE-ALT -90).
`lookupCameraCoords` longest-keyword scoring keeps `ANGLE-ALT`/`SIDE-ALT` winning over plain
`ANGLE`/`SIDE` for an *-ALT view, while a plain ANGLE/SIDE name matches only its own row.

`BACK` reuses HEAD's camera X/Y/Z (and OTTOMAN's tuned values) in `camera_coordinates.csv`;
only ROT differs (-180). Tested after `BACK-ANGLE` (which also contains "BACK") in both
`parseSweepName` and `lookupCameraCoords` (longest-keyword scoring keeps BACK-ANGLE winning for
back-angle views). Sofa BACK rows keep the ±55 override, same as their BACK-ANGLE rows.

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
  (`rootVars[1]`). **Camera X / Y / Z can be overridden per (variant, view-type) by the external
  `camera_coordinates.csv` (added 2026-06-23).**

- **Editable camera-coordinate CSV (added 2026-06-23, RATIO + full matrix 2026-06-23)** =
  `Sweep Script\camera_coordinates.csv`, next to the script. Columns
  `VARIANT,VIEWTYPE,X,Y,Z,RATIO`. Lets the user move cameras without editing code. Ships as a full
  matrix of **all 10 variants × all 5 view types** (50 rows), grouped by `#` section-header lines
  (CSV cannot store colours — `#` lines are the visual separation and are skipped by the reader).
  Matching: VARIANT exact (or `*`/`ANY` = any), VIEWTYPE = substring of the parsed view string (or
  `*`/`ANY`).
  - **Restructured into a `*` "variable" block (2026-06-24).** Instead of repeating every view for
    all 10 variants, the file now has ONE **`# ALL VARIANTS (*)`** block defining each view type
    once (`*,ANGLE`, `*,HEAD`, `*,BACK`, `*,ANGLE-ALT`, …) that applies to every variant, plus a
    per-variant override block ONLY where a variant truly differs (currently just **OTTOMAN**, which
    has hand-tuned X/Z/target for ANGLE/ANGLE-ALT/SIDE/SIDE-ALT/HEAD/BACK/DETAIL). A view a variant
    does not list is inherited from `*`. Edit a `*` row once → changes all variants. This is the
    "define the variable in one place" the user asked for.
  - **Scoring is now lexicographic (changed 2026-06-24): VIEWTYPE specificity is PRIMARY
    (`wk.count * 1000`), exact VARIANT is the tiebreaker (`+1`).** The old "exact variant `+1000`
    dominates" scoring broke `*` inheritance: a short exact-variant viewtype (e.g. `OTTOMAN,ANGLE`,
    score 1005) would hijack a longer view it's merely a substring of (an OTTOMAN **BACK-ANGLE**
    shot, where `*,BACK-ANGLE` should win). Now `*,BACK-ANGLE` (10000) beats `OTTOMAN,ANGLE` (5001),
    while `OTTOMAN,ANGLE` (5001) still beats `*,ANGLE` (5000) for an ANGLE shot. So longest viewtype
    always wins; exact variant only breaks ties between equally-specific viewtypes. (Lets OTTOMAN
    omit its BACK-ANGLE/TOP rows and inherit them from `*`.)
  - A blank X/Y/Z cell keeps the computed default
  (X=0, Y=−width/RATIO pullback, Z=`variantCameraHeights`); blank RATIO uses the built-in `0.33`.
  RATIO overrides the camera pull-back (`camDist = pkgW / RATIO`). Read via
  `readCoordRows`/`splitCsvLine` (manual comma split — `filterString` collapses empty fields and
  would lose blank columns). Seed values: `OTTOMAN,DETAIL → X=11,Z=81`; every `*,HEAD`/`*,SIDE →
  Z=57`; RATIO `0.33` everywhere. Wired through `buildSweep` → `cloneCameraToLayer ... coords`.

- **TX/TY/TZ target columns added to the CSV (added 2026-06-24)** — 3 columns appended after ROT
  (`...,ROT,TX,TY,TZ`, 10 cols total) move the camera **TARGET** node, mirroring the X/Y/Z camera
  columns. Row shape is now 18 elements: `#(vk, wk, hasX,x, hasY,y, hasZ,z, hasRatio,ratio,
  hasRot,rot, hasTX,tx, hasTY,ty, hasTZ,tz)` (target at indices 13–18). `readCoordRows` pads to ≥10
  fields and parses c[8]/c[9]/c[10]. In `cloneCameraToLayer` the target default `[0,0,tgtZ]` is
  overridden cell-by-cell (blank cell = keep default), same pattern as the camera position. Seed
  TX/TY/TZ are all blank (keep computed defaults).
  - **Camera/target identified by CLASS, not clone order (fixed 2026-06-24).** Earlier code took
    `camNode = newNodes[1]` / `tgtNode = newNodes[2]`, assuming `maxOps.cloneNodes` returns clones
    in the same order as `sources`. It does NOT guarantee that, so when Max returned them swapped,
    the camera X/Y/Z landed on the target node and the TX/TY/TZ on the camera — looked like "the
    CSV target isn't fetching / OTTOMAN DETAIL came out wrong". Now `cloneCameraToLayer` picks
    `camNode` via `superClassOf n == camera` and `tgtNode` as the other clone (fallback
    `camNode.target`); naming, `applyCameraLens`, and the position/target overrides all key off
    those. Target move is guarded by `isValidNode tgtNode`, so a free (targetless) camera still
    gets its X/Y/Z + lens applied.

- **Preview dialog now shows the REAL rotation (fixed 2026-06-24).** `sweepPreviewDlg` used to
  display `e.rot` — the name-based default from `parseSweepName` — which ignores the model's VARIANT
  and the CSV `ROT` column. So `OTTOMAN ... DETAIL` (real rot −15, but name-default 0) showed
  "no rot" in the preview even though `buildSweep` applied −15 correctly. The build was never wrong;
  the preview was. Fix: `sweepRollout.rotForEntry entry` mirrors `buildSweep`'s rotation decision
  WITHOUT cloning (resolve each model token's grid VARIANT via `lookupPair pendingPairs` honoring the
  OPEN/CLSD suffix → first = `chosenVar`, detect side-sofa overrides → `groupRotFor`), and the
  preview calls it instead of reading `e.rot`. Lives next to `getPendingEntries`; relies on
  `pendingPairs` being set before the preview opens (it is — RUN sets it just before `createDialog`).
  Edge case: if a model maps but fails to clone, `buildSweep`'s `chosenVar` (first *cloned*) could
  differ from `rotForEntry`'s (first *mapped*) — acceptable for a preview estimate.

- **Rotate failure is now reported, not swallowed (fixed 2026-06-24).** The `about ctr (rotate grp
  ...)` in `buildSweep` had a bare `catch()` that silently dropped any error, which would look like
  "rotation not working". It now prints `SWEEP: rotate FAILED for <name> : <exception>` via
  `getCurrentException()`.

- **TOP (overhead) view added 2026-06-24.** New view-type for overhead shots (e.g.
  `87214-08-TOP-SW-AGR` for the ottoman = model 08). The camera looks straight DOWN:
  default camera pos `[0, 0, width/RATIO]` (RATIO seeded `0.3` for top, so height ≈ 3.33×width) and
  target `[0, 0, 0]`. Implemented as an `isTop` arg to `cloneCameraToLayer` (computed at the call
  site via `matchPattern entry.viewStr "*TOP*"`): when `isTop`, the camera default Y becomes 0 and
  Z becomes `camDist` (instead of −camDist on Y at variant height), and the target default Z becomes
  0. The CSV X/Y/Z + TX/TY/TZ overrides still apply on top (blank = keep these top defaults). Group
  rotation for TOP is 0 (sofas keep ±55 via their CSV ROT). CSV ships a `TOP` row for **all 10
  variants** (`VARIANT,TOP,0,0,,0.3,0,0,0,0`; sofas use ROT ∓55). The ottoman `TOP` name was
  appended to `EXCEL\EXCEL FILE.xlsx` as an **inline string** row (row 65) — the stdlib
  `xlsx_reader.py` reads `inlineStr` cells, verified. `parseSweepName` has an explicit (no-op, 0°)
  `*TOP*` branch for clarity.

- **Forced Sensor & Lens on the cloned camera (added 2026-06-24)** — `applyCameraLens camNode`
  overrides the cloned VRay Physical Camera's lens to the agreed render setup: `specify_fov = false`
  (Field of view OFF), `film_width = 36.0` (Film gate mm), `focal_length = 100.0`,
  `zoom_factor = 1.0`. No new camera is created — these are set on `newNodes[1]` right after the
  `#copy` clone+rename in `cloneCameraToLayer`. Each assignment is wrapped in its own try/catch so a
  different camera class (where a property name differs) is a silent no-op, not a crash.

- **ROT column added to the CSV (added 2026-06-23)** — 7th column controls the whole-group Z
  rotation, so the rotation table is now editable without code. Row shape is now
  `#(vk, wk, hasX,x, hasY,y, hasZ,z, hasRatio,ratio, hasRot,rot)` (indices 11/12). `groupRotFor
  variant viewStr fallback` returns the CSV ROT when filled, else the fallback. In `buildSweep`:
  `overallRot = groupRotFor chosenVar viewStr entry.rot`, **then** `if hasLeftSofa → groupRotFor
  "LEFT SIDE SOFA" viewStr -55`, `if hasRightSofa → groupRotFor "RIGHT SIDE SOFA" viewStr 55`. The
  sofa overrides are applied AFTER the base lookup, so the documented "right sofa anywhere wins +55"
  case is preserved (chosenVar = first model's variant, but the sofa override keys on its own row).
  Seed ROT: ANGLE 35 / BACK-ANGLE −145 / SIDE 90 / HEAD 0 / DETAIL 0; LEFT SIDE SOFA rows −55,
  RIGHT SIDE SOFA rows +55 (all views). `chosenVar` is now computed once before the rotation block
  and reused for the camera. (User also hand-edited OTTOMAN/ANGLE RATIO to 0.3.)

- **Excel read now Python-first (added 2026-06-23).** `readExcelLines` tries `readExcelLinesPython`
  (Max's embedded interpreter runs `xlsx_reader.py`, a **pure-stdlib** `zipfile`+`re` xlsx parser —
  no Excel.exe, no pip install) and falls back to `readExcelLinesOLE` (the old OLE path) if Python
  is unavailable or returns nothing. Bridge: MaxScript sets `sweep_xlsx_path`/`sweep_out_path` and
  `exec(open(reader).read())` in one `python.Execute` call (only relies on `python.Execute`); the
  reader writes one row per line (UTF-8) to a temp `.txt` that MaxScript reads back. Chosen because
  the user's company **blocks installing any AI/ML or pip packages** — stdlib only.

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
file via SAVE/LOAD INI. Camera coords: `getCoordsPath`/`readCoordRows`/`splitCsvLine`/
`lookupCameraCoords` (reads `camera_coordinates.csv`). Excel: `readExcelLines` (wrapper) →
`readExcelLinesPython` (via `xlsx_reader.py`) / `readExcelLinesOLE` (fallback).

**New files (2026-06-23):** `Sweep Script\camera_coordinates.csv` (editable camera X/Y/Z),
`Sweep Script\xlsx_reader.py` (pure-stdlib xlsx reader).

## Gotchas (Sweep Script)
- **`global sweepRollout` forward declaration (added 2026-06-23).** `sweepPreviewDlg` is defined
  *before* `sweepRollout` but its handlers call `sweepRollout.getPendingEntries()` /
  `.executePending()`. Without the forward `global` declaration, those names bind to an `undefined`
  local on a fresh Max session's first evaluation → "Unknown property ... in undefined" at the
  preview's `open` handler. (It silently "worked" only on a *second* run, when the global already
  existed.) Keep the `global sweepRollout` line above `rollout sweepPreviewDlg`.
- **MaxScript has no `continue`.** Skip logic must use nested `if/else if`, not `continue` (a bare
  `continue` reads as an undefined global = silent no-op). See `readCoordRows`.

## Open assumptions to verify with user if issues arise
- Rotation is about the assembled group's bbox center. If rotation should be about world origin
  or a specific pivot, change the `about ctr` center in `buildSweep`.
- Back-edge = world-axis bbox max.y of each model at its current orientation. If a model's
  geometric "back" isn't its +Y bbox side, its orientation must be fixed in the scene first.

/*
SWEEP Script - 3ds Max
=======================

Builds named "sweep" arrangements from an Excel list of full names like:

    87214-08-ANGLE-SW-AGR
    87214-10-17-08-ANGLE-CLSD-SW-AGR
    87214-38-35-20-14-ANGLE-SW-AGR
    87214-25-SIDE-OPEN-SW-AGR

Name format:  SERIES - MODEL(s) - VIEWTYPE - [SW] - AGR
    SERIES    : e.g. 87214 (stripped using the SERIES NAME field)
    MODEL(s)  : one or more numeric tokens (08 / 10-17-08 / 38-35-20-14).
                Each number is mapped to a scene object in the grid.
    VIEWTYPE  : ANGLE / BACK-ANGLE / SIDE / HEAD(-ON) / DETAIL (plus any
                extra words such as CLSD / OPEN / CONTROL, which only
                affect rotation via the base keyword - see table below).
    SW, AGR   : ignored suffix tokens.

Per name the script:
    1. Clones the mapped object for every model number (in name order).
    2. Lays them out along +X, in name order, bounding boxes TOUCHING
       (no gap), with every piece's BACK edge (+Y) aligned to one straight
       line so all backs sit on the same row.
    3. Groups the whole set and rotates the group on Z by the angle that
       the VIEWTYPE implies.
    4. Drops the group on its own hidden layer, named after the full name,
       and clones the picked template camera into that layer, framed on
       the arrangement.

Rotation by VIEWTYPE (whole-group Z rotation):
    ANGLE  (incl. ANGLE-CLSD / ANGLE-OPEN) ....  35
    BACK-ANGLE ("ANGLE BACK") ................. -145
    SIDE   (incl. SIDE-CLSD / SIDE-OPEN) ......  90
    HEAD / HEAD-ON ............................   0  (no rotation)
    DETAIL (incl. DETAIL-CONTROL) ............   0  (no rotation)

Model -> object map:
    Each MODEL # maps to exactly ONE scene object (lookup is by MODEL #).
    The VARIANT column is a descriptive seating-type tag only
    (Single Seater / Double Seater / Arm Chair / Side Sofa / Ottoman).

Camera:
    Click "PICK VRAY CAMERA" and pick a template camera. A framed copy of
    that camera (+ its target) is cloned into EVERY package layer.

Workflow:
    1. Enter SERIES NAME (e.g. 87214).
    2. Map each MODEL # to a scene object (tag its seating type in VARIANT).
    3. Browse to the Excel file (one full name per row).
    4. Optional: PICK VRAY CAMERA (a framed copy is cloned into every layer).
    5. Click RUN -> preview lists every arrangement -> RUN to confirm.
*/

struct SweepName (fullName, nums, rot, viewStr)

rollout sweepPreviewDlg "Confirm Sweeps" width:460 height:530
(
    label   lblCount  ""       pos:[10,10]  width:440 height:18
    listbox lstNames  ""       pos:[10,33]  width:440 height:22
    button  btnRun    "RUN"    pos:[10,488] width:215 height:32
    button  btnCancel "Cancel" pos:[235,488] width:215 height:32

    on sweepPreviewDlg open do
    (
        local entries = sweepRollout.getPendingEntries()
        local items = #()
        for e in entries do
        (
            local nc = e.nums.count
            local suffix = if nc == 1 then " model" else " models"
            local rotTxt = if e.rot == 0.0 then "no rot" else (e.rot as string + " deg")
            append items (e.fullName + "   (" + nc as string + suffix + ", " + rotTxt + ")")
        )
        lstNames.items = items
        lblCount.text = entries.count as string + " arrangement(s) will be created. Click RUN to confirm."
    )

    on btnRun pressed do
    (
        destroyDialog sweepPreviewDlg
        sweepRollout.executePending()
    )

    on btnCancel pressed do destroyDialog sweepPreviewDlg
)

rollout sweepRollout "Sweep Script" width:560 height:620
(
    local noneTag        = "(none)"
    local pendingSeries  = ""
    local pendingPairs   = #()
    local pendingEntries = #()
    local pickedCamera   = undefined

    label lblTitle "SWEEP SCRIPT" pos:[210,12] width:200 height:18

    label    lblSeries "SERIES NAME" pos:[20,44]  width:110
    edittext edtSeries ""            pos:[140,41] width:400 height:22

    label  lblMapHeader "MODEL -> OBJECT MAP"   pos:[20,74]  width:160
    button btnRefresh   "REFRESH SCENE OBJECTS" pos:[360,70] width:180 height:24

    dotNetControl dgvKits "System.Windows.Forms.DataGridView" pos:[20,100] width:520 height:300

    dotNetControl btnDelKit "Button" pos:[160,410] width:130 height:30
    dotNetControl btnAddKit "Button" pos:[300,410] width:160 height:30

    label    lblExcel  "EXCEL FILE" pos:[20,455]  width:100
    edittext edtExcel  ""           pos:[125,452] width:320 height:22 readonly:true
    button   btnBrowse "..."        pos:[450,452] width:90  height:22

    -- ---- Camera section (inline) -----------------------------------
    label  lblCamHeader "CAMERA"             pos:[20,486]  width:100
    button btnPickCam   "PICK VRAY CAMERA"   pos:[20,506]  width:160 height:30
    label  lblPickedCam "(no camera picked)" pos:[190,513] width:350 height:18

    button btnRun "RUN" pos:[20,550] width:520 height:34

    fn trimAll s = trimLeft (trimRight s)

    fn isAllDigits s =
    (
        if s == "" then return false
        local ok = true
        for i = 1 to s.count do
        (
            local c = s[i]
            if c < "0" or c > "9" then ( ok = false ; exit )
        )
        ok
    )

    fn numEqual a b =
    (
        if a == b then return true
        local ia = undefined
        local ib = undefined
        try (ia = a as integer) catch()
        try (ib = b as integer) catch()
        (ia != undefined and ib != undefined and ia == ib)
    )

    fn collectSceneObjectNames =
    (
        local raw = #()
        for n in objects do append raw n.name
        sort raw
        local result = #(noneTag)
        for s in raw do append result s
        result
    )

    fn refreshSceneObjects =
    (
        local names = collectSceneObjectNames()
        local savedValues = #()
        local rowCount = dgvKits.RowCount
        for i = 0 to (rowCount - 1) do
        (
            local cell = dgvKits.Rows.Item[i].Cells.Item[2]
            local v = cell.Value
            local s = ""
            try (if v != undefined then s = v as string) catch()
            append savedValues s
        )
        for i = 0 to (rowCount - 1) do
            dgvKits.Rows.Item[i].Cells.Item[2].Value = undefined
        local colObj = dgvKits.Columns.Item[2]
        colObj.Items.Clear()
        for s in names do colObj.Items.Add s
        for i = 0 to (rowCount - 1) do
        (
            local prev = savedValues[i + 1]
            if prev != "" and findItem names prev > 0 then
                dgvKits.Rows.Item[i].Cells.Item[2].Value = prev
        )
    )

    -- Grid -> array of #(numStr, node, variantStr). Lookup is by MODEL #;
    -- variantStr is a descriptive seating-type tag only.
    fn buildLookup =
    (
        local pairs = #()
        local rowCount = dgvKits.RowCount
        for i = 0 to (rowCount - 1) do
        (
            local row = dgvKits.Rows.Item[i]
            local numStr = ""
            local varStr = ""
            local objStr = ""
            try (if row.Cells.Item[0].Value != undefined then numStr = trimAll (row.Cells.Item[0].Value as string)) catch()
            try (if row.Cells.Item[1].Value != undefined then varStr = (row.Cells.Item[1].Value as string))         catch()
            try (if row.Cells.Item[2].Value != undefined then objStr = (row.Cells.Item[2].Value as string))         catch()
            if numStr != "" and objStr != "" and objStr != noneTag then
            (
                local n = getNodeByName objStr
                if n != undefined and isValidNode n then
                    append pairs #(numStr, n, varStr)
            )
        )
        pairs
    )

    -- First #(object, variant) mapped to this MODEL #.
    fn lookupPair pairs num =
    (
        for p in pairs where (numEqual p[1] num) do return #(p[2], p[3])
        undefined
    )

    -- Parse one Excel name line into a SweepName, or undefined if it is
    -- not a valid row for this series.
    fn parseSweepName line series =
    (
        local txt = trimAll line
        if txt == "" then return undefined
        local parts = filterString txt "-"
        if parts.count == 0 then return undefined

        -- Strip trailing AGR then optional SW.
        if parts.count > 0 and (toUpper parts[parts.count]) == "AGR" then deleteItem parts parts.count
        if parts.count > 0 and (toUpper parts[parts.count]) == "SW"  then deleteItem parts parts.count

        -- Strip the series token (first). If a series is set and the first
        -- token does not match it, this row belongs to another series.
        if parts.count > 0 then
        (
            if series == "" or parts[1] == series then deleteItem parts 1
            else return undefined
        )

        -- Leading all-digit tokens are model numbers; the rest is viewtype.
        local nums = #()
        while parts.count > 0 and (isAllDigits parts[1]) do
        (
            append nums parts[1]
            deleteItem parts 1
        )
        if nums.count == 0 then return undefined

        local viewStr = ""
        for v in parts do
        (
            if viewStr != "" then viewStr += "-"
            viewStr += toUpper v
        )

        -- Whole-group Z rotation from the base keyword. BACK-ANGLE must be
        -- tested before ANGLE because it also contains "ANGLE".
        local rot = 0.0
        if matchPattern viewStr pattern:"*BACK*ANGLE*" then rot = -145.0
        else if matchPattern viewStr pattern:"*ANGLE*"  then rot = 35.0
        else if matchPattern viewStr pattern:"*SIDE*"   then rot = 90.0
        else if matchPattern viewStr pattern:"*HEAD*"   then rot = 0.0
        else if matchPattern viewStr pattern:"*DETAIL*" then rot = 0.0
        else rot = 0.0

        SweepName fullName:txt nums:nums rot:rot viewStr:viewStr
    )

    fn readExcelLines filePath =
    (
        local rows = #()
        local excelApp = undefined
        local workbook = undefined
        try
        (
            excelApp = createOleObject "Excel.Application"
            excelApp.Visible = false
            excelApp.DisplayAlerts = false
            workbook = excelApp.Workbooks.Open filePath
            local sheet = workbook.Worksheets.Item 1
            local used = sheet.UsedRange
            local rc = used.Rows.Count
            local cc = used.Columns.Count
            for r = 1 to rc do
            (
                local line = ""
                for c = 1 to cc do
                (
                    local cell = sheet.Cells.Item r c
                    local v = cell.Value2
                    if v != undefined then
                    (
                        local s = v as string
                        if s != "" then
                        (
                            if line != "" then line += " "
                            line += s
                        )
                    )
                )
                if line != "" then append rows line
            )
        )
        catch
        (
            local errInfo = getCurrentException()
            messageBox ("Failed to read Excel file.\n\nDetails:\n" + (errInfo as string) + "\n\nIf Excel.exe is hung, end it from Task Manager and try again.") title:"Excel Error"
        )
        if workbook != undefined then try (workbook.Close false) catch()
        if excelApp != undefined then try (excelApp.Quit()) catch()
        if excelApp != undefined then try (releaseOLEObject excelApp) catch()
        rows
    )

    fn safeCreateHiddenLayer layerName =
    (
        local layer = undefined
        try (layer = LayerManager.getLayerFromName layerName) catch()
        if layer == undefined then layer = LayerManager.newLayerFromName layerName
        try (layer.isHidden = true) catch()
        layer
    )

    fn collectHierarchy node =
    (
        local result = #(node)
        local idx = 1
        while idx <= result.count do
        (
            local cur = result[idx]
            for c in cur.children do append result c
            idx += 1
        )
        result
    )

    -- Clone the entire hierarchy of src and re-link parents so clones keep
    -- the same parent/child structure (guards against dropped group members).
    fn cloneAsInstance src =
    (
        local sources = collectHierarchy src
        local newNodes
        try (maxOps.cloneNodes sources cloneType:#instance newNodes:&newNodes) catch(newNodes = undefined)
        if newNodes == undefined or newNodes.count == 0 then return #()
        if newNodes.count == sources.count then
        (
            for i = 1 to sources.count do
            (
                local oldNode = sources[i]
                if oldNode.parent != undefined then
                (
                    local pIdx = findItem sources oldNode.parent
                    if pIdx > 0 then
                        try (newNodes[i].parent = newNodes[pIdx]) catch()
                )
            )
        )
        newNodes
    )

    -- The cloned root: node whose parent is not in the cloned set.
    fn findCloneRoot nodes =
    (
        for n in nodes do
        (
            if n.parent == undefined then return n
            local pInSet = false
            for other in nodes do
                if other == n.parent then ( pInSet = true ; exit )
            if not pInSet then return n
        )
        nodes[1]
    )

    -- World bbox of node, recursing through group members.
    fn groupUnionBBox topNode =
    (
        if not isGroupHead topNode then return #(topNode.min, topNode.max)
        local mn = [1e30, 1e30, 1e30]
        local mx = [-1e30, -1e30, -1e30]
        local found = false
        local visit = #(topNode)
        while visit.count > 0 do
        (
            local cur = visit[1]
            deleteItem visit 1
            for c in cur.children do
            (
                if isGroupHead c then append visit c
                else if isValidNode c then
                (
                    local dmn = c.min
                    local dmx = c.max
                    if dmn.x < mn.x then mn.x = dmn.x
                    if dmn.y < mn.y then mn.y = dmn.y
                    if dmn.z < mn.z then mn.z = dmn.z
                    if dmx.x > mx.x then mx.x = dmx.x
                    if dmx.y > mx.y then mx.y = dmx.y
                    if dmx.z > mx.z then mx.z = dmx.z
                    found = true
                )
            )
        )
        if not found then return #(topNode.min, topNode.max)
        #(mn, mx)
    )

    -- Move node (keeping its own orientation) so its bbox sits against the
    -- running X cursor, with its BACK edge (+Y, bbox max.y) on the common
    -- back line (Y=0) and resting on Z=0. Returns the new cursor (= this
    -- node's right edge) so the next piece touches it.
    fn placeAdjacentX node cursorX =
    (
        local bb = groupUnionBBox node
        local w  = bb[2].x - bb[1].x
        node.pos = node.pos + [cursorX - bb[1].x, 0.0 - bb[2].y, 0.0 - bb[1].z]
        cursorX + w
    )

    -- Clone the picked template camera + target into `layer`, rename after
    -- the package, and position to frame the arrangement (camera X locked
    -- to the package centre so it stays centred in frame).
    fn cloneCameraToLayer layer pkgName pkgNodes kitCount =
    (
        if pickedCamera == undefined or not isValidNode pickedCamera then return undefined
        local sources = #(pickedCamera)
        local tgt = undefined
        try (tgt = pickedCamera.target) catch()
        if tgt != undefined and isValidNode tgt then append sources tgt
        local newNodes
        try (maxOps.cloneNodes sources cloneType:#copy newNodes:&newNodes) catch(newNodes = undefined)
        if newNodes == undefined or newNodes.count == 0 then return undefined
        for n in newNodes where isValidNode n do
            try (layer.addNode n) catch()
        try
        (
            if newNodes.count >= 1 then newNodes[1].name = uniqueName (pkgName + "_Cam")
            if newNodes.count >= 2 then newNodes[2].name = uniqueName (pkgName + "_Cam_Target")
        )
        catch()

        if newNodes.count >= 2 then
        (
            local mn = [1e30, 1e30, 1e30]
            local mx = [-1e30, -1e30, -1e30]
            local found = false
            for n in pkgNodes where isValidNode n do
            (
                local nmn = n.min
                local nmx = n.max
                if nmn.x < mn.x then mn.x = nmn.x
                if nmn.y < mn.y then mn.y = nmn.y
                if nmn.z < mn.z then mn.z = nmn.z
                if nmx.x > mx.x then mx.x = nmx.x
                if nmx.y > mx.y then mx.y = nmx.y
                if nmx.z > mx.z then mx.z = nmx.z
                found = true
            )

            if found then
            (
                local camNode = newNodes[1]
                local tgtNode = newNodes[2]

                local pkgW = mx.x - mn.x          -- A: object width, end to end (X).
                local tgtZ = (mn.z + mx.z) / 2.0  -- target vertical mid of the object.
                local camZ = 100.0 + (kitCount as float) * 20.0

                -- Width : (target -> camera) distance ratio  A / B = 0.3,
                -- so B = A / 0.3. Camera is pulled back along -Y by B.
                local widthRatio = 0.3
                local camDist = if widthRatio > 0.0001 then (pkgW / widthRatio) else pkgW

                -- Target on the world X,Y origin; camera straight back on -Y.
                try (tgtNode.pos = [0, 0, tgtZ]) catch()
                try (camNode.pos = [0, -camDist, camZ]) catch()
            )
        )

        newNodes
    )

    fn buildSweep entry pairs missingOut =
    (
        local createdAll = #()
        local roots = #()
        local rootVars = #()
        local hasLeftSofa = false
        local hasRightSofa = false
        for num in entry.nums do
        (
            local pr = lookupPair pairs num
            if pr == undefined or not isValidNode pr[1] then
                appendIfUnique missingOut num
            else
            (
                local newNodes = #()
                try (newNodes = cloneAsInstance pr[1]) catch(newNodes = #())
                if newNodes.count > 0 then
                (
                    local root = findCloneRoot newNodes
                    if root != undefined and isValidNode root then
                    (
                        append roots root
                        append rootVars pr[2]
                        local vtag = trimAll (toUpper pr[2])
                        if vtag == "LEFT SIDE SOFA"  then hasLeftSofa = true
                        if vtag == "RIGHT SIDE SOFA" then hasRightSofa = true
                        for n in newNodes where isValidNode n do
                            appendIfUnique createdAll n
                    )
                )
            )
        )
        if roots.count == 0 then return 0

        -- Order left -> right: any LEFT SIDE SOFA first (extreme left), the
        -- rest in name order, then any RIGHT SIDE SOFA last (extreme right).
        -- Lay them out along +X, touching, backs on one line.
        local ordered = #()
        for i = 1 to roots.count do
            if (trimAll (toUpper rootVars[i])) == "LEFT SIDE SOFA" then append ordered roots[i]
        for i = 1 to roots.count do
        (
            local v = trimAll (toUpper rootVars[i])
            if v != "LEFT SIDE SOFA" and v != "RIGHT SIDE SOFA" then append ordered roots[i]
        )
        for i = 1 to roots.count do
            if (trimAll (toUpper rootVars[i])) == "RIGHT SIDE SOFA" then append ordered roots[i]

        local cursor = 0.0
        for r in ordered do cursor = placeAdjacentX r cursor

        -- Overall Z rotation: LEFT SIDE SOFA -> -55, RIGHT SIDE SOFA -> +55,
        -- otherwise the view-type angle. RIGHT is checked last, so if a right
        -- side sofa is anywhere in the list it wins (+55).
        local overallRot = entry.rot
        if hasLeftSofa  then overallRot = -55.0
        if hasRightSofa then overallRot =  55.0

        -- Group the set, rotate it about its bbox centre on Z, then move the
        -- whole group so its bbox centre sits at the world origin (X=0, Y=0).
        local grp = undefined
        try (grp = group ordered name:entry.fullName) catch(grp = undefined)
        if grp != undefined then
        (
            if overallRot != 0.0 then
            (
                local gbb = groupUnionBBox grp
                local ctr = (gbb[1] + gbb[2]) / 2.0
                try (about ctr (rotate grp (eulerangles 0 0 overallRot))) catch()
            )
            local gbb2 = groupUnionBBox grp
            local gctr = (gbb2[1] + gbb2[2]) / 2.0
            try (grp.pos = grp.pos - [gctr.x, gctr.y, 0.0]) catch()
        )

        -- Hidden layer named after the full name; add every node so hiding
        -- the layer hides the whole arrangement.
        local layer = safeCreateHiddenLayer entry.fullName
        for n in createdAll where isValidNode n do
            try (layer.addNode n) catch()
        if grp != undefined then try (layer.addNode grp) catch()

        -- Clone a framed copy of the template camera into this layer.
        cloneCameraToLayer layer entry.fullName createdAll roots.count

        createdAll.count
    )

    fn getSettingsPath =
    (
        local scriptDir = ""
        try ( scriptDir = getFilenamePath (getThisScriptFilename()) ) catch()
        if scriptDir == "" or scriptDir == undefined do scriptDir = "C:\\SWEEP SCRIPT\\Sweep Script\\"
        local settingsDir = scriptDir + "settings\\"
        try ( makeDir settingsDir all:true ) catch()
        settingsDir + "sweep_script_settings.ini"
    )

    fn saveSettings =
    (
        local p = getSettingsPath()
        try
        (
            setINISetting p "main" "Series"    edtSeries.text
            setINISetting p "main" "ExcelPath" edtExcel.text
            local camName = ""
            if pickedCamera != undefined and isValidNode pickedCamera then camName = pickedCamera.name
            setINISetting p "main" "CameraName" camName

            local rowCount = dgvKits.RowCount
            try
            (
                local oldKeys = getINISetting p "map"
                for k in oldKeys do delIniSetting p "map" k
            )
            catch()
            setINISetting p "map" "Count" (rowCount as string)
            for i = 0 to (rowCount - 1) do
            (
                local row = dgvKits.Rows.Item[i]
                local num = ""
                local var = ""
                local obj = ""
                try (if row.Cells.Item[0].Value != undefined then num = row.Cells.Item[0].Value as string) catch()
                try (if row.Cells.Item[1].Value != undefined then var = row.Cells.Item[1].Value as string) catch()
                try (if row.Cells.Item[2].Value != undefined then obj = row.Cells.Item[2].Value as string) catch()
                local idxStr = (i + 1) as string
                setINISetting p "map" ("Row" + idxStr + "_Num") num
                setINISetting p "map" ("Row" + idxStr + "_Var") var
                setINISetting p "map" ("Row" + idxStr + "_Obj") obj
            )
        )
        catch()
        OK
    )

    fn loadSettings =
    (
        local p = getSettingsPath()
        if not doesFileExist p then return undefined
        try
        (
            local s = getINISetting p "main" "Series"
            if s != "" then edtSeries.text = s
            local e = getINISetting p "main" "ExcelPath"
            if e != "" then edtExcel.text = e

            local cn = getINISetting p "main" "CameraName"
            if cn != "" then
            (
                local n = getNodeByName cn
                if n != undefined and isValidNode n then
                (
                    pickedCamera = n
                    lblPickedCam.text = "Camera: " + n.name
                )
            )

            local cs = getINISetting p "map" "Count"
            local count = 0
            if cs != "" then try (count = cs as integer) catch(count = 0)
            if count > 0 then
            (
                dgvKits.Rows.Clear()
                for i = 1 to count do dgvKits.Rows.Add()
                for i = 1 to count do
                (
                    local row = dgvKits.Rows.Item[i-1]
                    local idxStr = i as string
                    local num = getINISetting p "map" ("Row" + idxStr + "_Num")
                    local var = getINISetting p "map" ("Row" + idxStr + "_Var")
                    local obj = getINISetting p "map" ("Row" + idxStr + "_Obj")
                    try (if num != "" then row.Cells.Item[0].Value = num) catch()
                    try (if var != "" then row.Cells.Item[1].Value = var) catch()
                    try (if obj != "" then row.Cells.Item[2].Value = obj) catch()
                )
            )
        )
        catch()
        OK
    )

    -- Commit cell on Enter and swallow the key so it never triggers RUN.
    fn editorKeyDown s e =
    (
        local kv = 0
        try (kv = e.KeyValue) catch()
        if kv == 13 then
        (
            try (dgvKits.EndEdit()) catch()
            try (dgvKits.ClearSelection()) catch()
            try (e.SuppressKeyPress = true) catch()
            try (e.Handled = true) catch()
        )
    )

    fn getPendingEntries = pendingEntries

    fn executePending =
    (
        local series  = pendingSeries
        local pairs   = pendingPairs
        local entries = pendingEntries

        local createdCount = 0
        local skippedCount = 0
        local missing = #()

        undo "Sweep Run" on
        (
            for e in entries do
            (
                local n = buildSweep e pairs missing
                if n > 0 then createdCount += 1
                else skippedCount += 1
            )
        )

        local msg = "Created " + createdCount as string + " sweep layer(s)."
        if skippedCount > 0 then
            msg += "\nSkipped " + skippedCount as string + " (no mapped objects matched)."
        if missing.count > 0 then
        (
            msg += "\n\nUnmapped Model #:"
            for m in missing do msg += "\n  - " + m
        )
        if pickedCamera == undefined then
            msg += "\n\nNote: No template camera was picked, so no cameras were added to layers."
        messageBox msg title:"Run Complete"
    )

    on sweepRollout open do
    (
        dgvKits.AllowUserToAddRows = false
        dgvKits.AllowUserToDeleteRows = true
        dgvKits.AllowUserToResizeRows = false
        dgvKits.RowHeadersVisible = true
        dgvKits.RowHeadersWidth = 28
        dgvKits.MultiSelect = false
        dgvKits.AutoSizeColumnsMode = (dotNetClass "System.Windows.Forms.DataGridViewAutoSizeColumnsMode").Fill
        dgvKits.SelectionMode = (dotNetClass "System.Windows.Forms.DataGridViewSelectionMode").FullRowSelect
        dgvKits.ScrollBars = (dotNetClass "System.Windows.Forms.ScrollBars").Vertical

        local colNum = dotNetObject "System.Windows.Forms.DataGridViewTextBoxColumn"
        colNum.HeaderText = "MODEL #"
        colNum.FillWeight = 22
        dgvKits.Columns.Add colNum

        local colVar = dotNetObject "System.Windows.Forms.DataGridViewComboBoxColumn"
        colVar.HeaderText = "VARIANT"
        colVar.FillWeight = 32
        local varOptions = #("DEFAULT", "SINGLE SEATER", "DOUBLE SEATER", "LOVE SEAT", "ARM CHAIR", "LEFT SIDE SOFA", "RIGHT SIDE SOFA", "OTTOMAN")
        for s in varOptions do colVar.Items.Add s
        dgvKits.Columns.Add colVar

        local colObj = dotNetObject "System.Windows.Forms.DataGridViewComboBoxColumn"
        colObj.HeaderText = "SELECT OBJECT"
        colObj.FillWeight = 46
        dgvKits.Columns.Add colObj

        local names = collectSceneObjectNames()
        for s in names do colObj.Items.Add s

        dgvKits.Rows.Add()
        dgvKits.Rows.Add()

        local colorCls = dotNetClass "System.Drawing.Color"
        local flatStyle = (dotNetClass "System.Windows.Forms.FlatStyle").Flat
        local boldFont = dotNetObject "System.Drawing.Font" "Tahoma" 9.0 \
                         ((dotNetClass "System.Drawing.FontStyle").Bold)

        btnDelKit.Text      = "DELETE ROW"
        btnDelKit.BackColor = colorCls.FromArgb 178 34 34
        btnDelKit.ForeColor = colorCls.White
        btnDelKit.FlatStyle = flatStyle
        btnDelKit.Font      = boldFont

        btnAddKit.Text      = "+ ADD ROW"
        btnAddKit.BackColor = colorCls.FromArgb 60 130 60
        btnAddKit.ForeColor = colorCls.White
        btnAddKit.FlatStyle = flatStyle
        btnAddKit.Font      = boldFont

        loadSettings()
    )

    on sweepRollout close do saveSettings()

    on dgvKits DataError s e do (e.ThrowException = false)

    on dgvKits EditingControlShowing s e do
    (
        local ctrl = e.Control
        try (dotNet.removeAllEventHandlers ctrl "KeyDown") catch()
        dotNet.addEventHandler ctrl "KeyDown" editorKeyDown
    )

    on btnRefresh pressed do refreshSceneObjects()

    on btnAddKit click s e do dgvKits.Rows.Add()

    on btnDelKit click s e do
    (
        try
        (
            local rowIdx = -1
            if dgvKits.SelectedRows.Count > 0 then
                rowIdx = dgvKits.SelectedRows.Item[0].Index
            else if dgvKits.CurrentCell != undefined then
                rowIdx = dgvKits.CurrentCell.RowIndex
            if rowIdx >= 0 and rowIdx < dgvKits.RowCount then
            (
                dgvKits.Rows.RemoveAt rowIdx
                saveSettings()
            )
        )
        catch()
    )

    on btnBrowse pressed do
    (
        local startDir = ""
        if edtExcel.text != "" and doesFileExist edtExcel.text then
            startDir = edtExcel.text
        else
        (
            local scriptDir = ""
            try ( scriptDir = getFilenamePath (getThisScriptFilename()) ) catch()
            if scriptDir != "" and scriptDir != undefined then
            (
                local excelDir = scriptDir + "excel\\"
                try ( makeDir excelDir all:true ) catch()
                startDir = excelDir
            )
        )
        local fp = getOpenFileName caption:"Select Excel file" \
            filename:startDir \
            types:"Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*|"
        if fp != undefined then
        (
            edtExcel.text = fp
            saveSettings()
        )
    )

    on btnPickCam pressed do
    (
        local pick = pickObject message:"Pick a VRayPhysicalCamera (its target will be copied too)"
        if pick != undefined and isValidNode pick then
        (
            pickedCamera = pick
            lblPickedCam.text = "Camera: " + pick.name
            saveSettings()
        )
    )

    on btnRun pressed do
    (
        try (dgvKits.EndEdit()) catch()
        saveSettings()

        local series = trimAll edtSeries.text
        if series == "" then
        (
            messageBox "Enter the SERIES NAME first." title:"Missing Series"
            return undefined
        )
        local pairs = buildLookup()
        if pairs.count == 0 then
        (
            messageBox "Map at least one Model # to a scene object." title:"No Map"
            return undefined
        )
        local excelPath = edtExcel.text
        if excelPath == "" or not doesFileExist excelPath then
        (
            messageBox "Browse to a valid Excel file." title:"Missing Excel"
            return undefined
        )
        local lines = readExcelLines excelPath
        if lines.count == 0 then
        (
            messageBox "No data was read from the Excel file." title:"Empty"
            return undefined
        )

        local entries = #()
        for line in lines do
        (
            local e = parseSweepName line series
            if e != undefined then append entries e
        )
        if entries.count == 0 then
        (
            messageBox ("No valid name rows found for series \"" + series + "\".") title:"No Names"
            return undefined
        )

        pendingSeries  = series
        pendingPairs   = pairs
        pendingEntries = entries

        try (destroyDialog sweepPreviewDlg) catch()
        createDialog sweepPreviewDlg style:#(#dialog, #minimize, #sysmenu)
    )
)

createDialog sweepRollout

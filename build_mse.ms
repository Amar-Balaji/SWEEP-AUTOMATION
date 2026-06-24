/*
BUILD_MSE  -  encrypt the CURRENT sweep_script.ms into a shareable output folder.

Run this INSIDE 3ds Max (encryptScript only works in Max):
    Scripting -> Run Script... -> pick build_mse.ms
    (or in the Listener:  fileIn "C:\\SWEEP SCRIPT AGR\\Sweep Script\\build_mse.ms" )

It will:
    1. Encrypt the current sweep_script.ms  ->  sweep_script.mse
    2. Put a fresh sweep_script.mse + a copy of camera_coordinates.csv in  output\
    -> share the whole  output\  folder.
*/
(
    local d   = "C:\\SWEEP SCRIPT AGR\\Sweep Script\\"
    local out = d + "output\\"
    local srcMs = d + "sweep_script.ms"
    local mse   = d + "sweep_script.mse"
    local csv   = d + "camera_coordinates.csv"

    if not (doesFileExist srcMs) do ( messageBox ("Not found:\n" + srcMs) title:"build_mse"; return false )

    makeDir out all:true

    -- 1. Encrypt the CURRENT source -> sweep_script.mse (next to the source).
    try (deleteFile mse) catch()
    encryptScript srcMs

    if not (doesFileExist mse) do ( messageBox "encryptScript produced no .mse." title:"build_mse"; return false )

    -- 2. Place the fresh .mse + CSV copy into output\.
    try (deleteFile (out + "sweep_script.mse")) catch()
    copyFile mse (out + "sweep_script.mse")

    if doesFileExist csv do
    (
        try (deleteFile (out + "camera_coordinates.csv")) catch()
        copyFile csv (out + "camera_coordinates.csv")
    )

    messageBox ("Done. Fresh .mse built from current code.\nShare this folder:\n" + out) title:"build_mse - done"
    out
)

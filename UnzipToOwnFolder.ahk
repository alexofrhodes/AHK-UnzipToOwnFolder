IniFile := "config.ini"
Section := "Settings"
Key := "Hotkey"
Menu, Tray, Icon , UnzipToOwnFolder.ico
; Create INI with default hotkey if missing
IfNotExist, %IniFile%
    IniWrite, F1, %IniFile%, %Section%, %Key%

; Read hotkey from INI
IniRead, myHotkey, %IniFile%, %Section%, %Key%, F1

; Detect 7-Zip installation (Registry & Common Paths)
SevenZipPath := ""
RegRead, SevenZipPath, HKEY_LOCAL_MACHINE, SOFTWARE\7-Zip, Path
if (SevenZipPath != "")
    SevenZipPath := SevenZipPath "\7z.exe"

if (!FileExist(SevenZipPath))
{
    IfExist, C:\Program Files\7-Zip\7z.exe
        SevenZipPath := "C:\Program Files\7-Zip\7z.exe"
    else IfExist, C:\Program Files (x86)\7-Zip\7z.exe
        SevenZipPath := "C:\Program Files (x86)\7-Zip\7z.exe"
}

; Fallback to portable 7z.exe in script directory
if (!FileExist(SevenZipPath)) {
    Portable7z := A_ScriptDir "\7-ZipPortable\7-ZipPortable.exe"
    if (FileExist(Portable7z))
        SevenZipPath := Portable7z
}

Use7Zip := FileExist(SevenZipPath)

; Set hotkey only when Explorer is active
Hotkey, IfWinActive, ahk_exe explorer.exe
Hotkey, %myHotkey%, Main
Return

Main:
    ; Get selected files in Explorer
    clip := ClipToVar()
    if (clip = "") {
        MsgBox, No files selected!
        Return
    }

    Loop, parse, clip, `n, `r
    {
        targetFile := A_LoopField
        if !FileExist(targetFile)
            Continue

        ; Ensure it's a valid compressed file
        SplitPath, targetFile, , folder, ext, zipname_noext
        ext := "." ext
        validExtensions := ".zip .7z .rar .tar .gz .iso"

        if !InStr(validExtensions, ext)
            Continue

        ; Create temp extraction folder
        tempfolder := folder "\temp_" A_Now
        FileCreateDir, %tempfolder%
        Sleep, 100  ; Small delay to ensure the folder is recognized

        ; Use Shell COM for .zip, otherwise use 7-Zip
        if (ext = ".zip" && !Use7Zip) {
            success := UnzipWithShell(targetFile, tempfolder)
        } else if (Use7Zip) {
            success := UnzipWith7Zip(targetFile, tempfolder, SevenZipPath)
        } else {
            MsgBox, No valid extractor found! Please install 7-Zip or place `7z.exe` in the script folder.
            Continue
        }

        if (!success) {
            MsgBox, Failed to extract: %targetFile%
            Continue
        }

        ; Count extracted items
        allcount := 0, firstlevelcount := 0
        Loop, Files, %tempfolder%\*.*, FDR
        {
            allcount := A_Index
            onlyitem := A_LoopFilePath
        }
        Loop, Files, %tempfolder%\*.*, FD
        {
            firstlevelcount := A_Index
            firstlevelitem := A_LoopFilePath
        }

        ; Organize extracted files/folders
        if (allcount = 1) {
            if InStr(FileExist(onlyitem), "D") {
                SplitPath, onlyitem, onlyfoldername
                FileMoveDir, %onlyitem%, %folder%\%onlyfoldername%
            } else {
                FileMove, %onlyitem%, %folder%
                FileRemoveDir, %tempfolder%, 1
            }
        } else if (firstlevelcount = 1 && InStr(FileExist(firstlevelitem), "D")) {
            SplitPath, firstlevelitem, firstlevelfoldername
            FileMoveDir, %firstlevelitem%, %folder%\%firstlevelfoldername%
        } else {
            FileMoveDir, %tempfolder%, %folder%\%zipname_noext%
        }

        ; Cleanup temp folder if not needed
        FileRemoveDir, %tempfolder%, 1
    }

    ; Refresh Explorer to show extracted files
    If WinActive("ahk_exe explorer.exe")
        Send {F5}
Return

; Copy selected file paths from Explorer
ClipToVar() {
    cliptemp := ClipboardAll
    clipboard =
    Send ^c
    ClipWait, 1
    clip := clipboard
    clipboard := cliptemp
    return clip
}

; Unzip function using Windows Shell COM (for .zip files)
UnzipWithShell(zipfile, folder) {
    psh := ComObjCreate("Shell.Application")

    ; Ensure the destination folder exists
    if !FileExist(folder)
        FileCreateDir, %folder%

    ; Verify Shell COM object is valid
    destFolder := psh.Namespace(folder)
    zipItems := psh.Namespace(zipfile).Items

    if (!destFolder || !zipItems) {
        return false  ; Return failure if extraction cannot proceed
    }

    destFolder.CopyHere(zipItems, 4|16)
    Sleep, 500  ; Give time for extraction process
    return true  ; Return success
}

; Unzip function using 7-Zip (for .7z, .rar, .tar, .gz, .iso)
UnzipWith7Zip(zipfile, folder, sevenZipPath) {
    ; Extract using 7z.exe
    cmd := Format("""{1}"" x -y ""{2}"" -o""{3}""", sevenZipPath, zipfile, folder)
    RunWait, %ComSpec% /C %cmd%,, Hide

    return FileExist(folder)
}

;Smarter unzip sketch

IniFile := "config.ini"
Section := "Settings"
Key := "Hotkey"

; Create INI with default hotkey if missing
if !FileExist(IniFile)
    IniWrite, F1, IniFile, Section, Key

; Read hotkey from INI
IniRead, myHotkey, IniFile, Section, Key, F1

; Set hotkey only when Explorer is active
Hotkey, IfWinActive,  ahk_exe explorer.exe
Hotkey,  %myHotkey%, Main
Return

Main:



    ;get all selected files in Explorer and verify that it is .zip
    clip := ClipToVar()

    ;MsgBox % clip
    sleep, 200

    Loop, parse, clip, `n, `r
    {
    targetFile := A_LoopField

    ; MsgBox % targetFile
    sleep, 200

    if !FileExist(targetFile) or (SubStr(targetFile,-3) != ".zip")
    Continue
    
    ;unzip to temporary folder
    SplitPath, targetFile, , folder, , zipname_noext
    tempfolder := folder "\temp_" A_Now
    FileCreateDir, % tempfolder 
    Unzip(targetFile, tempfolder)
    
    ;count unzipped files/folders recursively
    Loop, Files, % tempfolder "\*.*", FDR
    {
    allcount := A_Index
    onlyitem := A_LoopFilePath
    }
    
    ;count unzipped files/folders at first level only
    Loop, Files, % tempfolder "\*.*", FD
    {
    firstlevelcount := A_Index
    firstlevelitem  := A_LoopFilePath
    }
    
    ;case1: only one file/folder in whole zip
    if (allcount = 1)
    {
    if InStr( FileExist(onlyitem), "D")
    {
        SplitPath, onlyitem, onlyfoldername
        FileMoveDir, % onlyitem, % folder "\" onlyfoldername
    }
    else
        FileMoveDir % tempfolder, % folder "\" zipname_noext ;;;
        FileMove   , % onlyitem, % folder
    }
    
    ;case2: only one folder (and no files) at the first level in zip
    else if (firstlevelcount = 1) and InStr(FileExist(firstlevelitem), "D")
    {
    SplitPath, firstlevelitem, firstlevelfoldername
    FileMoveDir, % firstlevelitem, % folder "\" firstlevelfoldername
    }
    
    ;case3: multiple files/folders at the first level in zip
    else
    {
    FileMoveDir % tempfolder, % folder "\" zipname_noext
    }
    
    ;cleanup temp folder
    FileRemoveDir, % tempfolder, 1
    
    }

    ;refresh Explorer to show results
    If WinActive("ahk_exe Explorer.exe")
    Send {F5}

return
 
 
;function: copy selection to clipboard to var
ClipToVar() {
  cliptemp := clipboardall ;backup
  clipboard = 
  send ^c
  clipwait, 1
  clip := clipboard
  clipboard := cliptemp    ;restore
  return clip
}
 
 
;function: unzip files to already existing folder
;zip file can have subfolders
Unzip(zipfile, folder)
{
  psh := ComObjCreate("Shell.Application")
  psh.Namespace(folder).CopyHere( psh.Namespace(zipfile).items, 4|16)
}
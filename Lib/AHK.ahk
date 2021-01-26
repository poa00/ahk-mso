; AHK Library
; Includes Startup, Compile
AHK_Close(ScriptFullPath){
DetectHiddenWindows, On
WinClose, %ScriptFullPath% ahk_class AutoHotkey
}

; ---------------------------------------------------------------------- 
AHK_IsRunning(ScriptFullPath){
DetectHiddenWindows, On
IfWinExist, %ScriptFullPath%
    return True
return False
}

; ---------------------------------------------------------------------- 

AHK_Compile(ScriptFullPath,FileIcon:="",CompileDir:=""){
; See https://www.autohotkey.com/boards/viewtopic.php?t=60944

SplitPath A_AhkPath,, AhkDir

SplitPath,ScriptFullPath, OutFileName
ExeFileName := StrReplace(OutFileName,".ahk",".exe")
Run, taskkill.exe /F /IM %ExeFileName%
If (FileIcon="") {
    FileIcon := StrReplace(ScriptFullPath,".ahk",".ico")
}

FileRead, sCode, %ScriptFullPath%
sCode := RegExReplace(sCode,"LastCompiled =.*","LastCompiled = " . A_Now)
FileDelete, %ScriptFullPath%
FileAppend, %sCode%, %ScriptFullPath%

; sCmd := "`"" . AhkDir . "\Compiler\Ahk2Exe.exe`" /in `"" . ScriptFullPath . "`""
sCmd = "%AhkDir%\Compiler\Ahk2Exe.exe" /in "%ScriptFullPath%"
If FileExist(FileIcon)
    sCmd = %sCmd% /icon "%FileIcon%"

; Move File to CompileDir

If Not (CompileDir="") {
    If Not InStr(CompileDir,":") { ; relative path
        SplitPath, ScriptFullPath, ScriptFileName, OutDir
        CompileDir = %OutDir%\%CompileDir%
    }
    DestFile = %CompileDir%\%ExeFileName%
    sCmd = %sCmd% /out "%DestFile%"
}
RunWait %sCmd%
}


; ---------------------------------------------------------------------- 

AHK_Exit(ScriptFullPath:=""){
If !ScriptFullPath
    ScriptFullPath=A_ScriptFullPath

SplitPath,ScriptFullPath, FileName, OutDir, Extension
If (Extension = "ahk") {
    DetectHiddenWindows, On
    WinClose, %ScriptFullPath% ahk_class AutoHotkey
} Else {
    ;ExeFileName := StrReplace(OutFileName,".ahk",".exe")
    sCmd = taskkill /f /im "%FileName%"
    Run %sCmd%,,Hide
}
}
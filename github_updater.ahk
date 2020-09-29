; GitHub updater
; kill task, copy .github mirror file and rerun
; called by PTCheckForUpdate
github_updater(A_Args[1])
SetWorkingDir %A_ScriptDir%
github_updater(FileName){
sCmd = taskkill /f /im "%FileName%"
RunWait %sCmd%,,UseErrorLevel Hide
ErrorLevelTaskKill := ErrorLevel
FileCopy, %FileName%.github, %FileName%, 1 ; Overwrite 
If (ErrorLevelTaskKill = 0) ; taskkill was successful->restart file
    Run %FileName% 
FileDelete %FileName%.github
} ; End function
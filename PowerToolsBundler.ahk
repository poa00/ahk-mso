#SingleInstance force

#Include <PowerTools>
#Include <AHK>
; Calls: Lib/ToStartup

LastCompiled = 20200806152811

IcoFile := RegExReplace(A_ScriptFullPath,"\..*",".ico")
If (FileExist(IcoFile)) 
	Menu,Tray,Icon, %IcoFile%

AppList = MO,ConnectionsEnhancer,NWS,OutlookShortcuts,PeopleConnector,TeamsShortcuts

sGuiTitle = PowerTools Bundle
Gui, New,,%sGuiTitle%

; ListView
Gui, Add, ListView, h200 w180 Checked, %sGuiTitle%  ; Create a ListView.
ImageListID := IL_Create(6)  ; Create an ImageList to hold 10 small icons.
LV_SetImageList(ImageListID)  ; Assign the above ImageList to the current ListView.


Loop, Parse, AppList, `,
{  ; Load the ImageList with a series of icons from the DLL.
    If a_iscompiled {
        IL_Add(ImageListID, A_LoopField . ".exe") 
    } Else {
        IcoFile := A_ScriptDir . "\" . A_LoopField . ".ico"
        If FileExist(IcoFile) {
            IL_Add(ImageListID,IcoFile) 
        }
    }
    LV_Add("Icon" . A_Index , A_LoopField)
} ; End Loop     


LV_ModifyCol()  ; Auto-adjust the column widths.

; https://jacksautohotkeyblog.wordpress.com/2019/12/30/use-autohotkey-gui-menu-bar-for-instant-hotkeys/

; MenuBar
Menu, ItemsMenu, Add, Check All`tCtrl+A, SelectAll
Menu, ItemsMenu, Add, Uncheck all`tCtrl+Shift+A, UncheckAll

If !a_iscompiled {
    Menu, ActionsMenu, Add, &Compile`tCtrl+C, Compile
    Menu, ActionsMenu, Add, Compile And &Push`tCtrl+P, CompileAndPush
    Menu, ActionsMenu, Add, Compile Bundler, CompileSelf
    Menu, ActionsMenu, Add, &Developper Mode`tCtrl+D, DevMode
    Menu, ActionsMenu, Add, Exe Mode, ExeMode
} Else {
    Menu, ActionsMenu, Add, Check for &Update/Download`tCtrl+U, CheckForUpdate
}
Menu, ActionsMenu, Add, Add to &Startup`tCtrl+S, AddToStartup
Menu, ActionsMenu, Add, &Run`tCtrl+R, Run
Menu, ActionsMenu, Add, E&xit`tCtrl+X, Exit


Menu, HelpMenu, Add, Open Help`tCtrl+H, OpenHelp
Menu, HelpMenu, Add, Check for Update, CheckForUpdateSelf

Menu, MyMenuBar, Add, &Items, :ItemsMenu 
Menu, MyMenuBar, Add, &Actions, :ActionsMenu 
Menu, MyMenuBar, Add, &Help, :HelpMenu
Gui, Menu, MyMenuBar

Gui, Show

; -------------------------------------------------------------------------------------------------------------------

SelectAll:
LV_Modify(0, "Check")  ; Uncheck all the checkboxes.
return
UncheckAll:
LV_Modify(0, "-Check")  ; Uncheck all the checkboxes.
return
; -------------------------------------------------------------------------------------------------------------------

CheckForUpdateSelf:
PTCheckForUpdate()
return
; -------------------------------------------------------------------------------------------------------------------

CheckForUpdate: ; CFU
RowNumber = 0
Loop {
    RowNumber := LV_GetNext(RowNumber, "Checked")
    if not RowNumber 
	    break
	
    LV_GetText(ItemName, RowNumber, 1)
	PTCheckForUpdate(ItemName)
}

; Update PowerTools.ini - only once
guExe = %A_ScriptDir%\github_updater.exe
sUrl = https://raw.githubusercontent.com/tdalon/ahk/master/PowerTools.ini
UrlDownloadToFile, %sUrl%, PowerTools.ini.github
sCmd = %guExe% PowerTools.ini
RunWait, %sCmd%,,Hide

Run %A_ScriptDir%
return
; -------------------------------------------------------------------------------------------------------------------

OpenHelp:
PTHelp("Bundler")
return
; -------------------------------------------------------------------------------------------------------------------
Exit: ; CFU
RowNumber = 0
Loop {
    RowNumber := LV_GetNext(RowNumber, "Checked")
    if not RowNumber 
	    break
	LV_GetText(ItemName, RowNumber, 1)
	If a_iscompiled 
	    ScriptFullPath = %A_ScriptDir%\%ItemName%.exe
    Else
	    ScriptFullPath = %A_ScriptDir%\%ItemName%.ahk

    AHKExit(ScriptFullPath)
}
return
; -------------------------------------------------------------------------------------------------------------------
Compile:
RowNumber = 0
Loop {
    RowNumber := LV_GetNext(RowNumber, "Checked")
    if not RowNumber 
	    break
	LV_GetText(ItemName, RowNumber, 1)
	ScriptFullPath = %A_ScriptDir%\%ItemName%.ahk
    AHKCompile(ScriptFullPath,,"PowerTools")
	}
Run %A_ScriptDir%\PowerTools
return
; -------------------------------------------------------------------------------------------------------------------

CompileSelf:
AHKCompile(A_ScriptFullPath,,"PowerTools")
return
; -------------------------------------------------------------------------------------------------------------------
CompileAndPush:
RowNumber = 0
FileList =
Loop {
    RowNumber := LV_GetNext(RowNumber, "Checked")
    if not RowNumber 
	    break
	LV_GetText(ItemName, RowNumber, 1)
	ScriptFullPath = %A_ScriptDir%\%ItemName%.ahk
    AHKCompile(ScriptFullPath,,"PowerTools")
    FileList =  %FileList% %ItemName%.exe
} 
RunWait, git add %FileList%, %A_ScriptDir%\PowerTools
RunWait, git commit -m "Update compiled powertools", %A_ScriptDir%\PowerTools
RunWait, git push origin master, %A_ScriptDir%\PowerTools
return
; -------------------------------------------------------------------------------------------------------------------
AddToStartup:
RowNumber = 0
Loop % LV_GetCount()
{
    RowNumber := A_Index
	SendMessage, 4140, RowNumber - 1, 0xF000, SysListView321  ; 4140 is LVM_GETITEMSTATE. 0xF000 is LVIS_STATEIMAGEMASK.
    IsChecked := (ErrorLevel >> 12) - 1  ; This sets IsChecked to true if RowNumber is checked or false otherwise.
    LV_GetText(ItemName, RowNumber, 1)

    If a_iscompiled 
	    ScriptFullPath = %A_ScriptDir%\%ItemName%.exe
    Else
	    ScriptFullPath = %A_ScriptDir%\%ItemName%.ahk
        
    ToStartup(ScriptFullPath,IsChecked)
}
Run %A_Startup%
return


; -------------------------------------------------------------------------------------------------------------------
Run:
RowNumber = 0
Loop % LV_GetCount()
{
    RowNumber := A_Index
	SendMessage, 4140, RowNumber - 1, 0xF000, SysListView321  ; 4140 is LVM_GETITEMSTATE. 0xF000 is LVIS_STATEIMAGEMASK.
    IsChecked := (ErrorLevel >> 12) - 1  ; This sets IsChecked to true if RowNumber is checked or false otherwise.
    LV_GetText(ItemName, RowNumber, 1)

    If a_iscompiled 
	    ScriptFullPath = %A_ScriptDir%\%ItemName%.exe
    Else
	    ScriptFullPath = %A_ScriptDir%\%ItemName%.ahk
    If IsChecked
        Run, %ScriptFullPath%
    Else
        AHKExit(ScriptFullPath)
}
return

; -------------------------------------------------------------------------------------------------------------------
StartMO:
Run %A_ScriptDir%\PowerTools\MO.exe
return
; -------------------------------------------------------------------------------------------------------------------

DevMode:
;DevList = ConNextEnhancer,NWS,OutlookShortcuts,PeopleConnector,TeamsShortcuts
Loop, Parse, AppList, `,
{  
    ScriptFullPath = %A_ScriptDir%\%A_LoopField%.ahk
    Run, %ScriptFullPath%
    ScriptFullPath = %A_ScriptDir%\%A_LoopField%.exe
    SplitPath,ScriptFullPath, FileName
    sCmd = taskkill /f /im "%FileName%"
    Run %sCmd%,,Hide     
} ; End Loop
;Run %A_ScriptDir%\PowerTools\MO.exe
return 

ExeMode:
Loop, Parse, AppList, `,
{  
    ScriptFullPath = %A_ScriptDir%\%A_LoopField%.ahk
    AHKExit(ScriptFullPath)        
    ScriptFullPath = %A_ScriptDir%\PowerTools\%A_LoopField%.exe
    Run, %ScriptFullPath%
} ; End Loop
return
; -------------------------------------------------------------------------------------------------------------------


GuiClose:
ExitApp
GuiEscape:
ExitApp

return    

; You can compile it via running the Ahk2Exe command e.g. D:\Programs\AutoHotkey\Compiler\Ahk2Exe.exe /in "Teamsy.ahk" /icon "icons\Teams.ico"
;#Include <Teams>
;#SingleInstance force ; for running from editor

Teamsy(A_Args[1])
ExitApp

Teamsy(sInput){
    
If (!sInput) { ; empty
    WinId := Teams_GetMainWindow()
    WinActivate, ahk_id %WinId%
    return
}

FoundPos := InStr(sInput," ")  

If FoundPos {
    sKeyword := SubStr(sInput,1,FoundPos-1)
    sInput := SubStr(sInput,FoundPos+1)
} Else {
    sKeyword := sInput
    sInput =
}

Switch sKeyword
{
Case "w": ; Web App
    Switch sInput
    {
    Case "c","cal":
        Teams_OpenWebCal()
    Default:
        Teams_OpenWebApp()
    }
    return
Case "h","-h","help":
    Run, https://tdalon.github.io/ahk/Teamsy
    return
Case "u":
    sKeyword = unread
Case "p":
    sKeyword = pop
Case "c":
    sKeyword = call
Case "f":
    sKeyword = find
Case "free","a":
    sKeyword = available
Case "s","save":
    sKeyword = saved
Case "d":
    sKeyword = dnd
Case "cal":
    WinId := Teams_GetMainWindow()
    WinActivate, ahk_id %WinId%
    SendInput ^4; open calendar
    return
Case "m","meet": ; create a meeting
    WinId := Teams_GetMainWindow()
    WinActivate, ahk_id %WinId%
    WinGetTitle Title, A
    If ! (Title="Calendar | Microsoft Teams") {
            SendInput ^4 ; open calendar
            Sleep, 300
            While ! (Title="Calendar | Microsoft Teams") { 
                WinGetTitle Title, A
                Sleep 500
            }
    }
    SendInput !+n ; schedule a meeting alt+shift+n
    return
Case "l","leave": ; leave meeting
    WinId := Teams_GetMeetingWindow()
    If !WinId ; empty
        return
    WinActivate, ahk_id %WinId%
    SendInput ^+b ; ctrl+shift+b
    return
Case "sh","share":  
    WinId := Teams_GetMeetingWindow()
    
    If !WinId ; empty
        return
    
    WinActivate, ahk_id %WinId%
    SendInput ^+e ; ctrl+shift+e 
    sleep, 1000
    SendInput {Tab}{Enter} ; Select first screen
    return
Case "mu","mute":  
    WinId := Teams_GetMainWindow()
    If !WinId ; empty
        return
    WinActivate, ahk_id %WinId%
    SendInput ^+m ;  ctrl+shift+m
    return
Case "de":  ; decline call
    WinId := Teams_GetMainWindow()
    If !WinId ; empty
        return
    WinActivate, ahk_id %WinId%
    SendInput ^+d ;  ctrl+shift+d 
    return
Case "q","quit": ; quit
    sCmd = taskkill /f /im "Teams.exe"
    Run %sCmd%,,Hide 
    return
Case "r","restart": ; restart
    WinId := Teams_GetMainWindow()
    sCmd = taskkill /f /im "Teams.exe"
    Run %sCmd%,,Hide 
    While WinActive("ahk_id " . WinId)
        Sleep 500
    Teams_GetMainWindow()
    return
Case "n","new","x": ; new expanded conversation 
    WinId := Teams_GetMainWindow()
    WinActivate, ahk_id %WinId%
    SendInput ^{f6} ; Activate posts tab https://support.microsoft.com/en-us/office/use-a-screen-reader-to-explore-and-navigate-microsoft-teams-47614fb0-a583-49f6-84da-6872223e74a0#picktab=windows
    ; workaround will flash the search bar if posts/content panel already selected but works now even if you have just selected the channel on the left navigation panel
    ;SendInput {Esc} ; in case expand box is already opened
    SendInput !+c ;  compose box alt+shift+c: necessary to get second hotkey working (regression with new conversation button)
    sleep, 300
    SendInput ^+x ; expand compose box ctrl+shift+x (does not work anymore immediately)
    sleep, 800
    SendInput +{Tab} ; move cursor back to subject line via shift+tab
    return
Case "v","vi": ; Toggle video with background
    WinId := Teams_GetMeetingWindow()
    If !WinId ; empty
        return
    WinActivate, ahk_id %WinId%
    SendInput ^+o ; toggle video Ctl+Shift+o
    ;SendInput ^+p ; toggle background blur
    return
} ; End Switch

WinId := Teams_GetMainWindow()
WinActivate, ahk_id %WinId%

Send ^e ; Select Search bar
If (SubStr(sKeyword,1,1) = "@") {
    SendInput @
    sleep, 300
    sInput := SubStr(sKeyword,2)
} Else {
    SendInput /
    sleep, 300
    SendInput %sKeyword%
    sleep, 500
    SendInput {enter}
}

If (!sInput) ; empty
    return
sleep, 500

;sLastChar := SubStr(sInput,StrLen(sInput)) 
doBreak := (SubStr(sInput,StrLen(sInput)) == "-")
If (doBreak) {
    sInput := SubStr(sInput,1,StrLen(sInput)-1) ; remove last -
}

SendInput %sInput%
sleep, 800
If !doBreak
    SendInput {enter}
} ; End function     

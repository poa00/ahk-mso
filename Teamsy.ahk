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
    ; MsgBox %sKeyword%`n%sInput% ; DBG
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
Case "u":
    sKeyword = unread
Case "p":
    sKeyword = pop
Case "c":
    sKeyword = call
Case "f","free","a":
    sKeyword = available
Case "s","save":
    sKeyword = saved
Case "d":
    sKeyword = dnd
Case "m","meet": ; create a meeting
    WinId := Teams_GetMainWindow()
    WinActivate, ahk_id %WinId%
    SendInput ^4; open calendar
    Sleep, 300
    SendInput !+n ; schedule a meeting alt+shift+n
    return
Case "l","leave": ; leave meeting
    WinId := Teams_GetMeetingWindow()
    WinActivate, ahk_id %WinId%
    SendInput ^+b  
    return
Case "sh","share":  
    WinId := Teams_GetMeetingWindow()
    WinActivate, ahk_id %WinId%
    SendInput ^+e ; expand compose box ctrl+shift+e ; does not work if no other has joined
    sleep, 1000
    SendInput {Tab}{Enter} 
    return
Case "mu","mute":  
    WinId := Teams_GetMeetingWindow()
    WinActivate, ahk_id %WinId%
    SendInput ^+m ; expand compose box ctrl+shift+m ; does not work if no other has joined
    return
Case "x","n","new": 
    WinId := Teams_GetMainWindow()
    WinActivate, ahk_id %WinId%
    SendInput ^+x ; expand compose box ctrl+shift+x
    sleep, 300
    SendInput +{Tab} ; move cursor back to subject line via shift+tab
    return
Case "v","vi": ; Activate video with background
    WinId := Teams_GetMeetingWindow()
    WinActivate, ahk_id %WinId%
    SendInput ^+o ; toggle video Ctl+Shift+o
    ;SendInput ^+p ; toggle background blur
    return
} ; End Switch

WinId := Teams_GetMainWindow()
WinActivate, ahk_id %WinId%

Send ^e ; Select Search bar
SendInput /
sleep, 300
SendInput %sKeyword%
sleep, 500
SendInput {enter}
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

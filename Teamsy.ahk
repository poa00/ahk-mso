; You can compile it via running the Ahk2Exe command e.g. D:\Programs\AutoHotkey\Compiler\Ahk2Exe.exe /in "Teamsy.ahk" /icon "icons\Teams.ico"
Teamsy(A_Args[1])

Teamsy(sInput){
    
If !WinActive("ahk_exe Teams.exe") {  
    WinActivate, ahk_exe Teams.exe
    If !WinActive("ahk_exe Teams.exe") { 
        TeamsExe = C:\Users\%A_UserName%\AppData\Local\Microsoft\Teams\current\Teams.exe
        Run, %TeamsExe%
    }
    WinWaitActive, ahk_exe Teams.exe
}
If (!sInput) ; empty
    return

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
Case "u":
    sKeyword = unread
Case "f","free","a":
    sKeyword = available
Case "s","save":
    sKeyword = saved
Case "d":
    sKeyword = dnd
Case "m","meet": ; create a meeting
    SendInput ^4; open calendar
    Sleep, 300
    SendInput !+n ; schedule a meeting alt+shift+n
    return
Case "l","leave": ; leave meeting
    SendInput ^+b  
    return
Case "sh","share":     
    SendInput ^+e ; expand compose box ctrl+shift+e ; does not work if no other has joined
    ;MsgBox %sKeyword% ; DBG
    sleep, 1000
    SendInput {Tab}{Enter} 
    return
Case "x","n","new": 
    SendInput ^+x ; expand compose box ctrl+shift+x
    sleep, 300
    SendInput +{Tab} ; move cursor back to subject line via shift+tab
    return
Case "v": ; Activate video with background
    SendInput ^+o ; toggle video Ctl+Shift+o
    SendInput ^+p ; toggle background blur
    return
} ; End Switch

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

     
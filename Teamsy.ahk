Teamsy(A_Args[1])

Teamsy(sInput){
    
If !WinActive("ahk_exe Teams.exe") {  
    EnvGet, userprofile , userprofile
    Run, %userprofile%\AppData\Local\Microsoft\Teams\current\Teams.exe
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
Case "s":
    sKeyword = saved
Case "d":
    sKeyword = dnd
Case "m","meet": ; create a meeting
    SendInput ^4; open calendar
    Sleep, 300
    SendInput !+n ; schedule a meeting alt+shift+n
    return
Case "x","n","new": 
    SendInput ^+x ; expand compose box ctrl+shift+x
    sleep, 300
    SendInput +{Tab} ; move cursor to subject line shift+tab
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

     
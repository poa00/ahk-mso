#Include <WinClipAPI>
#Include <WinClip>
global PT_wc
If !PT_wc
    PT_wc := new WinClip

; Calls: Lib/PasteText

Chromy(A_Args[1],A_Args[2])

Chromy(sInput,sProfile:=""){
FoundPos := InStr(sInput," ")   
sKeyword := SubStr(sInput,1,FoundPos-1)
sInput := SubStr(sInput,FoundPos+1)
;sInput := StrReplace(sInput, "#", "{#}")
If (sProfile == "" ){ ; no Profile
    If WinActive("ahk_exe chrome.exe") {   
        SendInput , ^t%sKeyword%{tab}
        ClipPasteText(sInput)
        SendInput {enter}
    } Else {
        Run, chrome.exe
        WinWaitActive, ahk_exe chrome.exe 
        SendInput , ^l%sKeyword%{tab}
        ClipPasteText(sInput)
        SendInput {enter}
    }
   
} Else { ; Profile passed as argument
 ; remove - too ugly to check the current profile
 ; if Chrome currently used simply run manually without Chromy if you want to stay in the current window; else a new Chrome window is opened
    If 0 { ; WinActive("ahk_exe chrome.exe")  
        sCurProfile := GetCurrentChromeProfile() 
        If (sCurProfile = sProfile) {
            SendInput , ^t%sKeyword%{tab}
            ClipPasteText(sInput)
            SendInput {enter}
            return
        }
    } 
    sCmd = chrome.exe --profile-directory="%sProfile%" "" ; "about:blank" bypass startup page redirect
    Run, %sCmd%
    WinWaitActive, ahk_exe chrome.exe 
    SendInput , ^l%sKeyword%{tab}
    ;sInput = %sKeyword% %sInput%
    ClipPasteText(sInput)
    SendInput {enter}
}
} ; End function


GetCurrentChromeProfile() {
SendInput , ^tchrome://version{enter}
WinWait About Version
ClipSaved := ClipboardAll
Clipboard = 
Send ^a
sleep, 300
Send ^c
ClipWait, 0.5
sAbout := Clipboard
; Restore Clipboard
Clipboard := ClipSaved
Send ^w ; Close About Tab

; Parse Profile Path
RegExMatch(sAbout,"Profile Path\t(.*)",sMatch)
sProfile := RegExReplace(sMatch1,".*\\","")
return sProfile
}

DoesNotWork(){
; https://autohotkey.com/board/topic/74322-winget-returns-wrong-pid-for-chromeexe/
WinGet, PrId, PID, A
    sQuery := "Select * from Win32_Process where ProcessId='" . PrId . "'" 
    for process in ComObjGet("winmgmts:").ExecQuery(sQuery){
        sCommandline := process.Commandline
        ; MsgBox %PrId% %sCommandline%  ; DBG
    }       

    sPat = chrome.exe.*--profile-directory="%sProfile%"    
    If RegExMatch(sCommandline, sPat) {
        SendInput , ^t%sKeyword%{tab}
        ClipPasteText(sInput)
        SendInput {enter}
    } Else {
        sCmd = chrome.exe --profile-directory="%sProfile%"
        Run, %sCmd%
        WinWait, New Tab
        WinActivate , New Tab
        SendInput , ^l%sKeyword%{tab}
        ClipPasteText(sInput)
        SendInput {enter}
     }
     
return
}


; ----------------------------------------------------------------------
ClipPasteText(Text){
global PT_wc
PT_wc.iSetText(Text)
PT_wc.iPaste()
}
; ----------------------------------------------------------------------
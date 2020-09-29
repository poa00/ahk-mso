; Creates an IT HPSM Ticket from a ConNext Forum Entry
; Forum Entry must be opened in current active browser
; See documentation
; Relies on ConNext2HPSM.ini in %A_ScriptDir%
; Calls: GetActiveBrowserUrl
; Does not require setting up Password

#Include <WinactiveBrowser>

#NoTrayIcon
#SingleInstance force ; for running from editor - avoid warning if another instance is running

; Ctrl+Win+T
^#t:: ; <--- [Browser] Create Ticket
; Constants
sTel = 7428

sForumUrl := GetActiveBrowserUrl()
;sForumTitle := ConNextGetForumTitle(sForumUrl)

WinGetActiveTitle, sForumTitle
; Remove trailing - containing program e.g. - Google Chrome
StringGetPos,pos,sForumTitle,%A_space%-,R
if (pos != -1)
    sForumTitle := SubStr(sForumTitle,1,pos)

sDesc := RegExReplace(sForumTitle," - .*","")
; Extract Category from Forum Title
If RegExMatch(sForumTitle,"- (.*) Forum -",sForumShortTitle) {
    sIniFile = %A_ScriptDir%\ConNext2HPSMCat.ini
    needle := sForumShortTitle1 "\t(.*)"
	Loop, Read, %sIniFile%
 		{
			If RegExMatch(A_LoopReadLine, needle, sCat) {
				sCat := sCat1
                Break
			}
		}
}
If !sCat
    sCat = %sForumShortTitle1%

If sCat = ConNext
    sDetails = See details %sForumUrl%. Please root directly to ConNext Support who has access to this page.
Else {
    sDetails = See attached PDF for details. (Source %sForumUrl%)
    Send ^P ; Print
}

Run https://hpsm-web.cw01.contiwan.com/sm/ess.do
MsgBox 0x1001, Wait..., Wait for Submit a Trouble Ticket. Click on the Ticket Form 'Preferred Phone number'. Press OK when ready to be filled.
IfMsgBox Cancel
    return

SendInput {Raw} %sTel%
Send {Tab}
SendInput {Raw} %sCat%
Send {Tab 3}
SendInput {Raw} %A_ComputerName%
Send {Tab 2}
SendInput {Raw} %sDesc%
Send {Tab}
SendInput {Raw} %sDetails%

return

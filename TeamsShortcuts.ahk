; MS Teams Keyboard Shortcuts
; Author: Thierry Dalon
; See user documentation here: https://connext.conti.de/blogs/tdalon/entry/teams_shortcuts_ahk
; Code Project Documentation is available on GitHub here: https://github.com/tdalon/ahk
; Source : https://github.com/tdalon/ahk/blob/master/TeamsShortcuts.ahk
;

LastCompiled = 20210111131721

#Include <Teams>
#Include <PowerTools>
#Include <WinClipAPI>
#Include <WinClip>

#SingleInstance force ; for running from editor

SubMenuSettings := PowerTools_MenuTray()
Menu, SubMenuSettings, Add, Teams PowerShell, MenuCb_ToggleSettingTeamsPowerShell
RegRead, TeamsPowerShell, HKEY_CURRENT_USER\Software\PowerTools, TeamsPowerShell
If (TeamsPowerShell) 
  Menu,SubMenuSettings,Check, Teams PowerShell
Else 
  Menu,SubMenuSettings,UnCheck, Teams PowerShell

Menu, SubMenuSettings, Add, Teams Personalize Mentions, MenuCb_ToggleSettingTeamsMentionPersonalize
TeamsMentionPersonalize := PowerTools_RegRead("TeamsMentionPersonalize")

If (TeamsMentionPersonalize) 
  Menu,SubMenuSettings,Check, Teams Personalize Mentions
Else 
  Menu,SubMenuSettings,UnCheck, Teams Personalize Mentions

Menu, SubMenuSettings, Add, Update Personal Information, GetMe

Menu,Tray,NoStandard
Menu,Tray,Add,Add to Teams Favorites, Link2TeamsFavs

Menu, SubMenuCustomBackgrounds, Add, Open Custom Backgrounds Folder, OpenCustomBackgrounds
Menu, SubMenuCustomBackgrounds, Add, Open GUIDEs Shared Backgrounds, OpenGUIDEsCustomBackgrounds

Menu, Tray, Add, Custom Backgrounds, :SubMenuCustomBackgrounds

Menu, Tray,Add,Start Second Instance, Teams_OpenSecondInstance
Menu, Tray,Add,Open Web App, Teams_OpenWebApp
Menu, Tray,Add
Menu, Tray,Add,Export Team Members, Users2Excel
Menu, Tray,Add,Refresh Teams List, Teams_ExportTeams
Menu, Tray,Add

Menu, SubMenuMeeting, Add, Open Teams Web Calendar, Teams_OpenWebCal

; Add Cursor Highlighter
Menu, SubMenuMeeting, Add, Cursor Highlighter, PowerTools_CursorHighlighter

Menu, SubMenuVLC, Add, Start VLC, VLCStart
Menu, SubMenuVLC, Add, Set Play Mode, VLCPlayMode
Menu, SubMenuVLC, Add, Reset, VLCReset

Menu, SubMenuMeeting, Add, VLC, :SubMenuVLC
Menu, Tray, Add, Meeting, :SubMenuMeeting
Menu, Tray,Add
Menu, Tray,Standard

; Tooltip
If !a_iscompiled {
	FileGetTime, LastMod , %A_ScriptFullPath%
} Else {
	LastMod := LastCompiled
}
FormatTime LastMod, %LastMod% D1 R

sTooltip = Teams Shortcuts %LastMod%`nUse 'Win+T' to open main menu in Teams.`nRight-Click on icon to access other functionalities.
Menu, Tray, Tip, %sTooltip%

; -------------------------------------------------------------------------------------------------------------------
Menu, TeamsShortcutsMenu, add, Smart &Reply (Alt+R), Teams_SmartReply
Menu, TeamsShortcutsMenu, add, &Quote Conversation (Alt+Q), QuoteConversation
Menu, TeamsShortcutsMenu, add, &New Expanded Conversation (Alt+N), NewConversation
Menu, TeamsShortcutsMenu, add, Create E&mail with link to current conversation (Win+M), ShareByMail
Menu, TeamsShortcutsMenu, add, Send Mentions (Win+Q), SendMentions
Menu, TeamsShortcutsMenu, add, Personalize &Mention (Win+1), PersonalizeMention
Menu, TeamsShortcutsMenu, add, View &Unread (Win+U), ViewUnread
Menu, TeamsShortcutsMenu, add, View &Saved (Win+S), ViewSaved
Menu, TeamsShortcutsMenu, add, &Pop-out Chat (Win+P), Pop
Menu, TeamsShortcutsMenu, add, Add to &Favorites, Link2TeamsFavs
; -------------------------------------------------------------------------------------------------------------------

; Reset Main WinId at startup because of some possible hwnd collision
PowerTools_RegWrite("TeamsMainWinId","")

#IfWinActive,ahk_exe Teams.exe


#1:: ; <--- Personalize Mention
PersonalizeMention:
Teams_PersonalizeMention()
return

;--- Compose in Expand mode
; Alt + N
!n::  ; <--- New Expanded Conversation
NewConversation:
	SendInput ^{f6}
    SendInput !+c ;  compose box alt+shift+c: necessary to get second hotkey working (regression with new conversation button)
    sleep, 300
    SendInput ^+x ; expand compose box ctrl+shift+x (does not work anymore immediately)
    sleep, 800
    SendInput +{Tab} ; move cursor back to subject line via shift+tab
return


; View Unread
; Win+U
#u:: ; <--- View Unread
ViewUnread:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/teams_shortcuts_ahk" 
	return
}
Send ^e ; Select Search bar
SendInput /unread 
Sleep 500
Send {Enter}
return	

; View Saved
; Win+S
#s:: ; <--- View Saved
ViewSaved:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/teams_shortcuts_ahk" 
	return
}
Send ^e ; Select Search bar
SendInput /saved 
Sleep 500
Send {Enter}
return

; Pop-out chat
; Win+P
#p:: ; <--- Pop-out chat
Pop:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/teams_shortcuts_ahk" 
	return
}
Send ^e ; Select Search bar
SendInput /pop  
Sleep 500
Send {Enter}
return

; Alt+e
!e:: ; <--- Edit
Send {Enter}
Send {Tab}
Send {Enter}
Sleep 500
Send {Down 1}
Send {Enter}
return	

; Alt+.
!.:: ; <--- (...) Actions menu
Send {Enter}
Send {Tab}
Send {Enter}
return	

#t::
Menu, TeamsShortcutsMenu, Show
return

; Win+C
#c:: ; <--- Copy Link
TeamsCopyLink()
return	
; -------------------------------------------------------------------------------------------------------------------
; Alt+R - 
!r:: ; <--- Smart Reply with quotation and link to current thread
Teams_SmartReply()
return

; -------------------------------------------------------------------------------------------------------------------
; Alt+Q
!q:: ; <--- Quote conversation
QuoteConversation:
Teams_SmartReply(False)
return
; -------------------------------------------------------------------------------------------------------------------

!m:: ; <--- Create eMail with link to current conversation
ShareByMail:
If GetKeyState("Ctrl") {
	;Run, "https://connext.conti.de/blogs/tdalon/entry/teams_smart_reply"
	return
}
rc := TeamsCopyLink()
If (rc =false)
	return

sHTMLBody = Hello<br>Following <a href="%clipboard%">this conversation</a> in Teams:
; Create Email using ComObj
Try
	MailItem := ComObjActive("Outlook.Application").CreateItem(0)
Catch
	MailItem := ComObjCreate("Outlook.Application").CreateItem(0)
;MailItem.BodyFormat := 2 ; olFormatHTML

MailItem.Subject := linktext
MailItem.HTMLBody := sHTMLBody
;****************************** 
;~ MailItem.Attachments.Add(NewFile)
MailItem.Display ;Make email visible
;~ mailItem.Close(0) ;Creates draft version in default folder
;MailItem.Send() ;Sends the email

; Select email body
Send {Tab 3} 
Send {PgDn}
Send {Enter}
Send ^v
return

; ----------------------------------------------------------------------
#q::
SendMentions:
If GetKeyState("Ctrl") {
	Run, "https://tdalon.blogspot.com/teams-shortcuts-send-mentions"
	return
}
sEmailList := Clipboard
doPerso := PowerTools_RegRead("TeamsMentionPersonalize")
Teams_Emails2Mentions(sEmailList, doPerso)

return

; ----------------------------------------------------------------------
Link2TeamsFavs:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/teams_shortcuts_favorites"
	return
}
sUrl := clipboard
Link2TeamsFav(sUrl)
return

; ----------------------------------------------------------------------
Users2Excel:
TeamLink := Clipboard
sPat = \?groupId=([^&]*)
If (RegExMatch(TeamLink,sPat,sId)) {
	sGroupId := sId1
	Teams_Users2Excel(sGroupId)
} Else
	Teams_Users2Excel()
return

; ----------------------------------------------------------------------
OpenCustomBackgrounds:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/teams_custom_background"
	return
}
Run, %A_AppData%\Microsoft\Teams\Backgrounds\Uploads
return

OpenGUIDEsCustomBackgrounds:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/teams_custom_background"
	return
}
Run, "https://continental.sharepoint.com/:f:/r/teams/team_10000035/Shared Documents/Collaborate/TEAMS Background Pictures"
return
; ----------------------------------------------------------------------

GetMe:
If GetKeyState("Ctrl") {
	Run, "https://tdalon.blogspot.com/2020/11/teams-shortcuts-smart-reply.html#getme"
	return
}
suc := People_GetMe()
If (suc) {
	TrayTipAutoHide("Personal information updated!","Email, OfficeUid, Display Name were stored to the registry.")
}
return

; ----------------------------------------------------------------------

CursorHighliter:
Run %CHFile%
return

; ######################################################################
PersonalizeMentions:
; Does not work!
Send ^a
sHtml := GetSelection("html")
sPat = Us)<span .* itemtype="http://schema.skype.com/Mention".*>(.*)</span>
sNewHtml := sHtml
Pos = 1 
While Pos := RegExMatch(sHtml,sPat,sFullName,Pos+StrLen(sFullName)){
    sFullName1 := RegExReplace(sFullName1," (.*)","")
	FirstName := RegExReplace(sFullName1,".*, ","")
    sNewHtml := StrReplace(sNewHtml,sFullName1,FirstName)
}
;sNewHtml := RegExReplace(sNewHtml,"s).*<html>","<html>")
WinClip.SetHTML(sNewHtml)
WinClip.Paste()

return

; ---------------------------------------------------------------------- SUBFUNCTIONS
MeetingSetup(){
; Open Teams Calendar in the Browser
Teams_OpenWebCal()
	
}


VLCStart() {
If !WinActive("ahk_exe vlc.exe") {  
    WinActivate, ahk_exe vlc.exe
    If !WinActive("ahk_exe vlc.exe") { 
        VLCExe := PowerTools_RegGet("VLCExe")
		If (VLCExe = "") 
			return
        
		Run, %VLCExe%
    }
    WinWaitActive, ahk_exe vlc.exe
}

Send ^c ; Configure Capture Device
} ; eofun
; ----------------------------------------------------------------------

VLCPlayMode(){
If !WinActive("ahk_exe vlc.exe") {  
    WinActivate, ahk_exe vlc.exe
    If !WinActive("ahk_exe vlc.exe") {
	    VLCStart()
		return
	}
}
SendInput ^h ; Minimal Interface
WinSet, AlwaysOnTop , On, ahk_exe vlc.exe
WinSet, Style, -0xC00000, ahk_exe vlc.exe ; remove title bar
} ; eofun
; ----------------------------------------------------------------------

VLCReset(){
If !WinActive("ahk_exe vlc.exe") {  
    WinActivate, ahk_exe vlc.exe
    If !WinActive("ahk_exe vlc.exe") {
	    VLCStart()
		return
	}
}
WinSet, Style, -0xC00000, ahk_exe vlc.exe ; toggle title bar
SendInput ^h ; Minimal Interface
WinSet, AlwaysOnTop , Off, ahk_exe vlc.exe
} ; eofun
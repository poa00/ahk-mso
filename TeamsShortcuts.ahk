; MS Teams Keyboard Shortcuts
; Author: Thierry Dalon
; See user documentation here: https://connext.conti.de/blogs/tdalon/entry/teams_shortcuts_ahk
; Code Project Documentation is available on ContiSource GitHub here: http://github.conti.de/ContiSource/ahk
; Source : http://github.conti.de/ContiSource/ahk/blob/master/TeamsShortcuts.ahk
;

LastCompiled = 20200806155137

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

Menu,Tray,NoStandard
Menu,Tray,Add,Add to Teams Favorites, Link2TeamsFavs

Menu, SubMenuCustomBackgrounds, Add, Open Custom Backgrounds Folder, OpenCustomBackgrounds
Menu, SubMenuCustomBackgrounds, Add, Open GUIDEs Shared Backgrounds, OpenGUIDEsCustomBackgrounds

Menu, Tray, Add, Custom Backgrounds, :SubMenuCustomBackgrounds

Menu,Tray,Add,Start Second Instance, Teams_OpenSecondInstance
Menu,Tray,Add,Open Web App, Teams_OpenWebApp
Menu,Tray,Add
Menu,Tray,Add,Export Team Members, Users2Excel
Menu,Tray,Add,Refresh Teams List, Teams_ExportTeams
Menu,Tray,Add
Menu,Tray,Standard

; Tooltip
If !a_iscompiled {
	FileGetTime, LastMod , %A_ScriptFullPath%
} Else {
	LastMod := LastCompiled
}
FormatTime LastMod, %LastMod% D1 R

sTooltip = Teams Shortcuts %LastMod%`nUse 'Win+T' to open main Menu. Ctrl+Click on menu item to open help.`nRight-Click on icon to access Help and other functionalities.
Menu, Tray, Tip, %sTooltip%

; -------------------------------------------------------------------------------------------------------------------
Menu, TeamsShortcutMenu, add, Smart &Reply (Win+R), Teams_SmartReply
Menu, TeamsShortcutMenu, add, Reply with &Quote from Clipboard (Win+Q), ReplyWithQuote
Menu, TeamsShortcutMenu, add, Create E&mail with link to current conversation (Win+M), ShareByMail
Menu, TeamsShortcutMenu, add, &New Expanded Conversation (Win+N), NewConversation
Menu, TeamsShortcutMenu, add, Personalize &Mention (Win+1), PersonalizeMention
Menu, TeamsShortcutMenu, add, Personalize &Mention with uid (Win+2), PersonalizeMention2
Menu, TeamsShortcutMenu, add, View &Unread (Win+U), ViewUnread
Menu, TeamsShortcutMenu, add, View &Saved (Win+S), ViewSaved
Menu, TeamsShortcutMenu, add, &Pop-out Chat (Win+P), Pop
Menu, TeamsShortcutMenu, add, Add to &Favorites, Link2TeamsFavs
; -------------------------------------------------------------------------------------------------------------------


#IfWinActive,ahk_exe Teams.exe

#1:: ; <--- Personalize Mention
PersonalizeMention:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/personalize_mention_powertool"
	return
}
SendInput ^{Left}{Backspace}
return

#2:: ; <--- Personalize Mention
PersonalizeMention2:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/personalize_mention_powertool"
	return
}
SendInput {Backspace}^{Left}{Backspace}
return

;--- Compose in Expand mode
; Win + N
#n::  ; <--- New Expanded Conversation
NewConversation:
	SendInput ^+x ; Ctrl+Shift+x Expand
	sleep, 300
    SendInput +{Tab} ; move cursor to subject line
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
Menu, TeamsShortcutMenu, Show
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
; Alt+m - 
!m:: ; <--- New meeting
SendInput ^4; open calendar
Sleep, 300
SendInput !+n ; schedule a meeting alt+shift+n
return

#q:: ; <--- Reply with quote from clipboard
ReplyWithQuote:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/teams_smart_reply#reply_with_quote"
	return
}
sHtml := WinClip.getHTML()
Teams_SmartReply(sHtml,False)
return
; -------------------------------------------------------------------------------------------------------------------

#m:: ; <--- Create eMail with link to current conversation
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


; ######################################################################
PersonalizeMentions:
; Does not work!
Send ^a
sHtml := GetSelection("html")
sPat = Us)<span .* itemtype="http://schema.skype.com/Mention".*>(.*)</span>
sNewHtml := sHtml
Pos = 1 
While Pos := RegExMatch(sHtml,sPat,sFullName,Pos+StrLen(sFullName)){
    FirstName := RegExReplace(sFullName1,".*, ","")
    sNewHtml := StrReplace(sNewHtml,sFullName1,FirstName)
}
;sNewHtml := RegExReplace(sNewHtml,"s).*<html>","<html>")
WinClip.SetHTML(sNewHtml)

Sleep 500
Send ^v ; replace

return

; ---------------------------------------------------------------------- SUBFUNCTIONS

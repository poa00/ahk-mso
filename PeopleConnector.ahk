; Author: Thierry Dalon
; Standalone AHK File also delivered as compiled Standalone EXE file
; See help/homepage: https://connext.conti.de/blogs/tdalon/entry/people_connector

; Calls: ExtractEmails, TrayTipAutoHide, ToStartup
LastCompiled = 20200806155136

#SingleInstance force ; for running from editor

#Include <WinClipAPI>
#Include <WinClip>
#Include <People>
#Include <Connections>
#Include <Teams>
#Include <PowerTools>
#Include <WinActiveBrowser>

PT_Config := PowerTools_GetConfig()
RegRead, PT_TeamsOnly, HKEY_CURRENT_USER\Software\PowerTools, TeamsOnly
PowerTools_ConnectionsRootUrl := PowerTools_RegRead("ConnectionsRootUrl")
SubMenuSettings := PowerTools_MenuTray()

; -------------------------------------------------------------------------------------------------------------------
; SETTINGS
Menu, SubMenuSettings, Add,Set Password, Login_SetPassword
Menu, SubMenuSettings, Add, Notification at Startup, MenuCb_ToggleSettingNotificationAtStartup

RegRead, SettingNotificationAtStartup, HKEY_CURRENT_USER\Software\PowerTools, NotificationAtStartup
If (SettingNotificationAtStartup = "")
	SettingNotificationAtStartup := True ; Default value
If (SettingNotificationAtStartup) {
    Menu, SubMenuSettings, Check, Notification at Startup
}
Else {
    Menu, SubMenuSettings, UnCheck, Notification at Startup
}

Menu, SubMenuSettings, Add, Teams PowerShell, MenuCb_ToggleSettingTeamsPowerShell
RegRead, TeamsPowerShell, HKEY_CURRENT_USER\Software\PowerTools, TeamsPowerShell
If (TeamsPowerShell) 
  Menu,SubMenuSettings,Check, Teams PowerShell
Else 
  Menu,SubMenuSettings,UnCheck, Teams PowerShell


; -------------------------------------------------------------------------------------------------------------------
If !a_iscompiled {
	FileGetTime, LastMod , %A_ScriptFullPath%
	FormatTime LastMod, %LastMod% D1 R
} Else {
	LastMod := LastCompiled
}
sTextMenuTip = Double Tap 'Shift' to open menu.`nRight-Click on icon to access Help.
Menu, Tray, Tip, People Connector - %LastMod%`n%sTextMenuTip%
sText = Double Tap 'Shift' to open menu after selection.`nRight-Click on icon to access Help/ Support/ Check for Updates.

If (SettingNotificationAtStartup)
    TrayTip "People Connector is running!", %sText%
;TrayTipAutoHide("People Connector is running!",sText)

Menu, MainMenu, add, Teams &Chat, TeamsChat
Menu, MainMenu, add, Teams &Pop out Chat, TeamsPop
Menu, MainMenu, add, &Teams Call, TeamsCall
Menu, MainMenu, add, Create Teams Meeting, TeamsMeet
Menu, MainMenu, add, Add to Teams &Favorites, Emails2TeamsFavs
Menu, MainMenu, add, Add Users to Team, Emails2TeamUsers
Menu, MainMenu, add, Teams Chat - Copy Link, TeamsChatCopyLink
Menu, MainMenu,Add ; Separator
If !(PT_TeamsOnly)
    Menu, MainMenu, add, &Skype Chat, SkypeChat
Menu, MainMenu, add, Soft Phone Call, Tel
Menu, MainMenu,Add ; Separator
Menu, MainMenu, add, Create &Email, MailTo
Menu, MainMenu, add, Create Outlook &Meeting, MeetTo
If (PowerTools_ConnectionsRootUrl != "") {
    Menu, MainMenu, add, Copy Connections &at-Mentions, CNEmail2Mention
    Menu, MainMenu,Add ; Separator
    Menu, MainMenu, add, &Open Connections Profile, CNOpenProfile
}

Menu, MainMenu,Add ; Separator
Menu, MainMenu, add, Open o365/&Delve Profile, DelveProfile
Menu, MainMenu, add, &Bing Search, BingSearch
Menu, MainMenu, add, &LinkedIn Search By Name, LinkedInSearchByName
Menu, MainMenu, add, Stream Profile, StreamProfile
If !(PT_Config=Conti)
    Menu, MainMenu, add, People&View OrgChart (MySuccess), PeopleView
Menu, MainMenu,Add ; Separator
Menu, MainMenu, add, Copy Office &Uids, Emails2Uids
Menu, MainMenu, add, Copy Windows Uids, Emails2WinUids
Menu, MainMenu, add, Uids to Emails, winUids2Emails
Menu, MainMenu, add, Copy Emails, CopyEmails
Menu, MainMenu,Add ; Separator
If (PowerTools_ConnectionsRootUrl != "") {
    Menu, SubMenuCNMentions, add, &Emails, ConNextMentions2Emails
    Menu, SubMenuCNMentions, add, &Teams Chat, ConNextMentions2TeamsChat
    Menu, SubMenuCNMentions, add, &Mentions (extract), ConNextMentions2Emails
    Menu, MainMenu, add, (Connections) Mentions to, :SubMenuCNMentions
}
Menu, MainMenu, add, (Outlook) Addressees to Excel, Outlook2Excel
return

Shift:: ; (Double Press) <--- Open People Connector Menu
If (A_PriorHotKey = A_ThisHotKey and A_TimeSincePriorHotkey < 500) {
    sSelection := GetSelection("html")
    If !(sSelection) ; empty
        sSelection := GetSelection()
    If !(sSelection) { ; empty
        sSelection := Clipboard
        ;TrayTipAutoHide("People Connector warning!","You need first to select something!")   
        ;return
    } 
    ;MsgBox %sSelection% ; DBG

    If WinActive("ahk_exe OUTLOOK.exe")
        Menu, MainMenu, Enable, (Outlook) Addressees to Excel
    Else
        Menu, MainMenu, Disable, (Outlook) Addressees to Excel
    
    If (PowerTools_ConnectionsRootUrl != "") {
        If WinActiveBrowser()
            Menu, MainMenu, Enable, (Connections) Mentions to
        Else
            Menu, MainMenu, Disable, (Connections) Mentions to
    }
    
    Menu, MainMenu, Show
}
return


; ----------------------------  Menu Callbacks -------------------------------------
TeamsChat: 
sEmailList := GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
Teams_Emails2Chat(sEmailList)

;If InStr(sEmailList,";") ; multiple Emails
;    Teams_Emails2ChatDeepLink(sEmailList, askOpen:= true)
;Else {    
 ;   EnvGet, userprofile , userprofile
;    Run,  %userprofile%\AppData\Local\Microsoft\Teams\current\Teams.exe sip:%sEmailList%
;    Run,  msteams://teams.microsoft.com/l/chat/0/0?users=%sEmailList%
;}
return

; ------------------------------------------------------------------
TeamsMeet: 
sEmailList := GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
Teams_Emails2Meeting(sEmailList)

return

; ------------------------------------------------------------------
TeamsPop: 
sEmailList := GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
Teams_Pop(sEmailList)

return

; ------------------------------------------------------------------
TeamsChatCopyLink:
sEmailList := GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
Teams_Emails2ChatDeepLink(sEmailList)
return

TeamsCall:
sEmailList := GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
If InStr(sEmailList,";") { ; multiple Emails
    TrayTipAutoHide("People Connector warning!","Feature does not work for multiple users!")   
    return
} Else {
    EnvGet, userprofile , userprofile
    Run,  %userprofile%\AppData\Local\Microsoft\Teams\current\Teams.exe callto:%sEmailList%
}
return

; ------------------------------------------------------------------
SkypeChat:
sEmailList := GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
If InStr(sEmailList,";") { ; multiple Emails
    TrayTipAutoHide("People Connector warning!","Feature does not work for multiple users!")   
    return
} Else {
    sSip :=RegExReplace(sEmailList, "@.*","@continental.com")
    sCmd = "C:\Program Files (x86)\Microsoft Office\root\Office16\lync.exe" sip:%sSip%
    Run,  %sCmd%
}
return

; ------------------------------------------------------------------
Tel:
sEmailList := People_GetEmailList(sSelection)
If !sEmailList { ; empty
    If !RegexMatch(sSelection,"[1-9\(\)-\s]*")  {
        TrayTipAutoHide("People Connector warning!","You shall have an email or phone number selected!")   
    } Else {
        sTelNum := StrReplace(sSelection, " ","")
        sTelNum := StrReplace(sTelNum, "-","")
        ; If number starts with 0 prepend a 0
        If SubStr(sTelNum, 1, 1) = "0" {
	        sTelNum := SubStr(sTelNum, 2)
	        sTelNum = +49%sTelNum%
        }
        Run tel:%sTelNum%
    }
    return
}
If InStr(sEmailList,";") { ; multiple Emails
    TrayTipAutoHide("People Connector warning!","Feature does not work for multiple users!")   
    return
} Else {
    Run tel:%sEmailList%
}
return

Im:
sEmailList := People_GetEmailList(sSelection)
If !sEmailList { ; empty    
    return
}
sEmailList := StrReplace(sEmailList, ";",",")
MsgBox %sEmailList%
Run im:%sEmailList%

return

CNEmail2Mention:
sEmailList := GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
sHtmlMentions := Connections_Emails2Mentions(sEmailList)
WinClip.SetHtml(sHtmlMentions)
TrayTipAutoHide("People Connector","Mentions were copied to clipboard in RTF!")   
return


ConNextMentions2Mentions:
sHtml := GetSelection("html")
sHtml := CNMentions2Mentions(sHtml)
WinClip.SetHtml(sHtml)
TrayTipAutoHide("Copy Mentions", "Mentions were copied to the clipboard in RTF!")
return

ConNextMentions2TeamsChat:
sHtml := GetSelection("html")
sEmailList := CNMentions2Emails(sHtml)
Teams_Emails2ChatDeepLink(sEmailList)
return

ConNextMentions2Emails:
sHtml := GetSelection("html")
sEmailList := CNMentions2Emails(sHtml)
WinClip.SetText(sEmailList)
TrayTipAutoHide("Copy Emails", "Emails were copied to the clipboard!")
return

; ------------------------------------------------------------------
CopyEmails:
sEmailList := GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
WinClip.SetText(sEmailList)
TrayTipAutoHide("Copy Emails", "Emails were copied to the clipboard!")
return

Emails2Uids:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/people_connector_get_userid"
	return
}
sEmailList := GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
sUidList := Emails2Uids(sEmailList)
clipboard := sUidList
TrayTipAutoHide("Copy Uid","Uids " . sUidList . " were copied to the clipboard!")   
return

Emails2WinUids:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/people_connector_get_userid"
	return
}
sEmailList := GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
sUidList := Emails2Uids(sEmailList,"sAMAccountName")
clipboard := sUidList
TrayTipAutoHide("Copy Uid","Uids " . sUidList . " were copied to the clipboard!")   
return

; ------------------------------------------------------------------
winUids2Emails:
sEmailList := winUids2Emails(sSelection)
clipboard := sEmailList
TrayTipAutoHide("Copy Emails","Emails " . sEmailList . " were copied to the clipboard!")   
return

; ------------------------------------------------------------------
Emails2TeamsFavs:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/people_connector_teams_favorites"
	return
}
sEmailList := GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
Teams_Emails2Favs(sEmailList)
return

; ------------------------------------------------------------------
Emails2TeamUsers:
sEmailList := GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
Teams_Emails2Users(sEmailList)

return

Outlook2Excel:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/people_connector_ol2xl"
	return
}
OL2XL(sSelection)
return

MailTo:
sEmailList := GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}

Try
	MailItem := ComObjActive("Outlook.Application").CreateItem(0)
Catch
	MailItem := ComObjCreate("Outlook.Application").CreateItem(0)
;MailItem.BodyFormat := 2 ; olFormatHTML

MailItem.To := sEmailList
MailItem.Display ;Make email visible
return

MeetTo:
sEmailList := GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
Try
	oItem := ComObjActive("Outlook.Application").CreateItem(1)
Catch
	oItem := ComObjCreate("Outlook.Application").CreateItem(1)
;MailItem.BodyFormat := 2 ; olFormatHTML
oItem.MeetingStatus := 1
Loop, parse, sEmailList, ";"
{
    oItem.Recipients.Add(A_LoopField) 
}	
oItem.Display ;Make email visible
return

; ------------------------------------------------------------------
DelveProfile:
sEmailList := GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
Loop, parse, sEmailList, ";"
{
   Run,  https://continental-my.sharepoint.com/person.aspx?user=%A_LoopField%&v=profiledetails
}	
return

; ------------------------------------------------------------------
CNOpenProfile:
People_ConnectionsOpenProfile(sSelection)
return

; ------------------------------------------------------------------
StreamProfile:
sEmailList := GetEmailList(sSelection)
If (sEmailList != "") {
    Loop, parse, sEmailList, ";"
    {
    Run, https://web.microsoftstream.com/browse?q=%A_LoopField%&view=people
    }	
} Else {
    sName := People_GetName(sSelection)
    Run, https://web.microsoftstream.com/browse?q=%sName%&view=people
}
return

; ------------------------------------------------------------------
LinkedInSearchByName:
sName := People_GetName(sSelection)
sName := RegExReplace(sName,"\d*","") ; remove any numbers
Run, https://www.linkedin.com/search/results/people/?keywords=%sName%
return

BingSearch:
sSelection := People_GetName(sSelection)
Run, http://www.bing.com/search?q=%sSelection%#,Person 
return

PeopleView:
People_PeopleView(sSelection)
return

; ----------------------------------------------------------------------
ConNextAuth:
CNAuth()
return

; ----------------------------------------------------------------------
AddShortcutToADPhonebook:
EnvGet userprofile, userprofile
sLnk = %userprofile%\Desktop\Phonebook.lnk
sFile = \\cw01.contiwan.com\root\Loc\sbas\did43653\Public\SMT2009Tools\PusADTel\Release\PusADTel_OpenScape\PusADTel.exe
FileCreateShortcut, %sFile%, %sLnk%
TrayTipAutoHide("AD Phonebook","Shortcut was copied to your profile desktop")  
return

; ----------------------------------------------------------------------

; ---------------------------------------------------------------------- STARTUP -------------------------------------------------
MenuCb_ToggleSettingNotificationAtStartup:

If (SettingNotificationAtStartup := !SettingNotificationAtStartup) {
  Menu, SubMenuSettings, Check, Notification at Startup
}
Else {
  Menu, SubMenuSettings, UnCheck, Notification at Startup
}
PowerTools_RegWrite("NotificationAtStartup",SettingNotificationAtStartup)
return
; ------------------------------- SUBFUNCTIONS ----------------------------------------------------------


GetEmailList(sSelection){
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "")  { 
    sInput := GetSelection("html")
    sInput := StrReplace(sInput,"%40","@") ; for connext profile links
    sEmailList := People_GetEmailList(sInput)
}
return sEmailList
}
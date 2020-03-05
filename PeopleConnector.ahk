; Author: Thierry Dalon
; Standalone AHK File also delivered as compiled Standalone EXE file
; See help/homepage: https://connext.conti.de/blogs/tdalon/entry/people_connector

; Calls: ExtractEmails, TrayTipAutoHide, ToStartup
LastCompiled = 20200304150629

#SingleInstance force ; for running from editor

#Include <WinClipAPI>
#Include <WinClip>
#Include <People>
#Include <ConNext>
#Include <Teams>
#Include <PowerTools>
#Include <WinActiveBrowser>

SubMenuSettings := PTMenuTray()

; -------------------------------------------------------------------------------------------------------------------
; SETTINGS
Menu, SubMenuSettings, Add,Set Password, PasswordSet
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
Menu, MainMenu, add, &Teams Call, TeamsCall
Menu, MainMenu, add, Add to Teams &Favorites, Emails2TeamsFavs
Menu, MainMenu, add, Add Members to Team, Emails2TeamMembers
Menu, MainMenu,Add ; Separator
Menu, MainMenu, add, &Skype Chat, SkypeChat
Menu, MainMenu, add, Soft &Phone Call, Tel
Menu, MainMenu,Add ; Separator
Menu, MainMenu, add, &Create Email, MailTo
Menu, MainMenu, add, Create Outlook &Meeting, MeetTo
Menu, MainMenu, add, Copy ConNext &at-Mentions, ConNextEmail2Mention
Menu, MainMenu,Add ; Separator
Menu, MainMenu, add, &Open ConNext Profile, ConNextProfile
Menu, MainMenu, add, ConNext Profile Search By Name, ConNextProfileSearchByName
Menu, MainMenu,Add ; Separator
Menu, MainMenu, add, Open o365/&Delve Profile, DelveProfile
Menu, MainMenu, add, &LinkedIn Search By Name, LinkedInSearchByName
Menu, MainMenu, add, Stream Profile, StreamProfile
Menu, MainMenu,Add ; Separator
Menu, MainMenu, add, Copy Office &Uids, Emails2Uids
Menu, MainMenu, add, Copy Windows Uids, Emails2WinUids
Menu, MainMenu, add, Uids to Emails, winUids2Emails
Menu, MainMenu, add, Copy &Emails, CopyEmails
Menu, MainMenu,Add ; Separator
Menu, SubMenuCNMentions, add, &Emails,ConNextMentions2Emails
Menu, SubMenuCNMentions, add, &Teams Chat,ConNextMentions2TeamsChat
Menu, SubMenuCNMentions, add, &Mentions (extract),ConNextMentions2Emails
Menu, MainMenu, add, (ConNext) Mentions to, :SubMenuCNMentions
Menu, MainMenu, add, (Outlook) Addressees to Excel, Outlook2Excel
return

;<^>!c:: ; <--- Open People Connector Menu
Shift:: ; (Double Press) <--- Open People Connector Menu
If (A_PriorHotKey = A_ThisHotKey and A_TimeSincePriorHotkey < 500) {
    
    If (WinActiveBrowser()) {
        sUrl := GetActiveBrowserUrl()
	     If InStr(sUrl,"://connext.conti.de/homepage/web/updates/") {
            sSelection := GetSelection("html")
         } Else    
            sSelection:= GetSelection()
    } Else {
        sSelection:= GetSelection()
    }
    
    If !(sSelection) { ; empty
        TrayTipAutoHide("People Connector warning!","You need first to select something!")   
        return
    } 

    If WinActive("ahk_exe OUTLOOK.exe")
        Menu, MainMenu, Enable, (Outlook) Addressees to Excel
    Else
        Menu, MainMenu, Disable, (Outlook) Addressees to Excel
    
    If WinActiveBrowser()
        Menu, MainMenu, Enable, (ConNext) Mentions to
    Else
        Menu, MainMenu, Disable, (ConNext) Mentions to
    
    
    Menu, MainMenu, Show

}
return


; ----------------------------  Menu Callbacks -------------------------------------
TeamsChat: 
sEmailList := GetEmailList(sSelection)
If (sEmailList = "")
    return ; empty
 Emails2TeamsChatDeepLink(sEmailList, askOpen:= true)

;If InStr(sEmailList,";") ; multiple Emails
;    Emails2TeamsChatDeepLink(sEmailList, askOpen:= true)
;Else {    
 ;   EnvGet, userprofile , userprofile
;    Run,  %userprofile%\AppData\Local\Microsoft\Teams\current\Teams.exe sip:%sEmailList%
;    Run,  msteams://teams.microsoft.com/l/chat/0/0?users=%sEmailList%
;}
return

TeamsCall:
sEmailList := GetEmailList(sSelection)
If (sEmailList = "")
    return ; empty
If InStr(sEmailList,";") { ; multiple Emails
    TrayTipAutoHide("People Connector warning!","Feature does not work for multiple users!")   
    return
} Else {
    EnvGet, userprofile , userprofile
    Run,  %userprofile%\AppData\Local\Microsoft\Teams\current\Teams.exe callto:%sEmailList%
}
return

SkypeChat:
sEmailList := GetEmailList(sSelection)
If (sEmailList = "")
    return ; empty
If InStr(sEmailList,";") { ; multiple Emails
    TrayTipAutoHide("People Connector warning!","Feature does not work for multiple users!")   
    return
} Else {
    sSip :=RegExReplace(sEmailList, "@.*","@continental.com")
    sCmd = "C:\Program Files (x86)\Microsoft Office\root\Office16\lync.exe" sip:%sSip%
    Run,  %sCmd%
}
return

Tel:
sEmailList := ExtractEmails(sSelection)
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

ConNextProfile:
sEmailList := GetEmailList(sSelection)
If (sEmailList = "")
    return ; empty

Loop, parse, sEmailList, ";"
{
     Run,  https://connext.conti.de/profiles/html/profileView.do?email=%A_LoopField%
}	; End Loop Parse Clipboard
return

ConNextProfileSearchByName:
sSelection:= StrReplace(sSelection,",","")
Run, https://connext.conti.de/profiles/html/simpleSearch.do?searchBy=name&searchFor=%sSelection%
return


ConNextEmail2Mention:
sEmailList := GetEmailList(sSelection)
If (sEmailList = "")
    return ; empty
sHtmlMentions := CNEmails2Mentions(sEmailList)
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
Emails2TeamsChatDeepLink(sEmailList)
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
If (sEmailList = "")
    return ; empty
WinClip.SetText(sEmailList)
TrayTipAutoHide("Copy Emails", "Emails were copied to the clipboard!")
return

Emails2Uids:
GetKeyState, KeyState, Ctrl
If (KeyState = "D"){
	Run, "https://connext.conti.de/blogs/tdalon/entry/people_connector_get_userid"
	return
}
sEmailList := GetEmailList(sSelection)
If (sEmailList = "")
    return ; empty
sUidList := Emails2Uids(sEmailList)
clipboard := sUidList
TrayTipAutoHide("Copy Uid","Uids " . sUidList . " were copied to the clipboard!")   
return

Emails2WinUids:
GetKeyState, KeyState, Ctrl
If (KeyState = "D"){
	Run, "https://connext.conti.de/blogs/tdalon/entry/people_connector_get_userid"
	return
}
sEmailList := GetEmailList(sSelection)
If (sEmailList = "")
    return ; empty
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
GetKeyState, KeyState, Ctrl
If (KeyState = "D") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/people_connector_teams_favorites"
	return
}
sEmailList := GetEmailList(sSelection)
If (sEmailList = "")
    return ; empty
Emails2TeamsFavs(sEmailList)
return

; ------------------------------------------------------------------
Emails2TeamMembers:
GetKeyState, KeyState, Ctrl
If (KeyState = "D") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/emails2teammembers"
	return
}
RegRead, TeamsPowerShell, HKEY_CURRENT_USER\Software\PowerTools, TeamsPowerShell
If (TeamsPowerShell := !TeamsPowerShell) {
    sUrl := "https://connext.conti.de/wikis/home/wiki/Wc4f94c47297c_42c8_878f_525fd907cb68/page/Teams%20PowerShell%20Setup"
    Run, "%sUrl%" 
    MsgBox 0x1024,People Connector, Have you setup Teams PowerShell on your PC?
	IfMsgBox No
		return
    OfficeUid := AD_GetUserField("sAMAccountName=" . A_UserName, "mailNickname") ; mailNickname - office uid 
    RegWrite, REG_SZ, HKEY_CURRENT_USER\Software\PowerTools, OfficeUid, %OfficeUid%
    ;sWinUid := AD_GetUserField("mail=" . sEmail, "sAMAccountName")  ;- login uid

    RegWrite, REG_SZ, HKEY_CURRENT_USER\Software\PowerTools, TeamsPowerShell, %TeamsPowerShell%
    Menu,SubMenuSettings,Check, Teams PowerShell

}
sEmailList := GetEmailList(sSelection)
If (sEmailList = "")
    return ; empty
Emails2TeamMembers(sEmailList)

return

Outlook2Excel:
GetKeyState, KeyState, Ctrl
If (KeyState = "D") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/people_connector_ol2xl"
	return
}
OL2XL(sSelection)
return

MailTo:
sEmailList := GetEmailList(sSelection)
If (sEmailList = "")
    return ; empty

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
If (sEmailList = "")
    return ; empty
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
If (sEmailList = "")
    return ; empty
Loop, parse, sEmailList, ";"
{
   Run,  https://continental-my.sharepoint.com/person.aspx?user=%A_LoopField%&v=profiledetails
}	
return

StreamProfile:
sEmailList := GetEmailList(sSelection)
If (sEmailList = "") {
    sName := Trim(sSelection)
    If Not (InStr(sName, ",")) {
        arr := StrSplit(sName, " ")
        sName := arr[2] . ", " arr[1]
    }
    Run, https://web.microsoftstream.com/browse?q=%sName%&view=people
    return ; empty
}
    
Loop, parse, sEmailList, ";"
{
  Run, https://web.microsoftstream.com/browse?q=%A_LoopField%&view=people
}	
return

LinkedInSearchByName:
sName := StrReplace(sSelection,",","")
sName := RegExReplace(sName, "<.*>","") ; Strip email adress from Outlook
Run, https://www.linkedin.com/search/results/people/?keywords=%sName%
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
PasswordSet:
SetPassword()
return
; ---------------------------------------------------------------------- STARTUP -------------------------------------------------
MenuCb_ToggleSettingNotificationAtStartup:

If (SettingNotificationAtStartup := !SettingNotificationAtStartup) {
  Menu, SubMenuSettings, Check, Notification at Startup
}
Else {
  Menu, SubMenuSettings, UnCheck, Notification at Startup
}

RegWrite, REG_SZ, HKEY_CURRENT_USER\Software\PowerTools, NotificationAtStartup, %SettingNotificationAtStartup%
return
; ------------------------------- SUBFUNCTIONS ----------------------------------------------------------

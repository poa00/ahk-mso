; Author: Thierry Dalon
; Documentation: https://tdalon.github.io/ahk/NWS-PowerTool
; Code Project Documentation is available on ContiSource GitHub here: https://github.com/tdalon/ahk
;  Source: https://github.com/tdalon/ahk/blob/master/NWS.ahk

SetWorkingDir %A_ScriptDir%

#Include <WinClipAPI>
#Include <WinClip>
#Include <Connections>
#Include <Jira>
#Include <Confluence>
#Include <Login>
#Include <PowerTools>
#Include <WinActiveBrowser>
#Include <IntelliPaste>
#Include <Teams>
#Include <SharePoint>
#Include <Explorer>
#Include <People>

LastCompiled = 20201109213900

global PowerTools_ConnectionsRootUrl
If (PowerTools_ConnectionsRootUrl="") {
	If FileExist("PowerTools.ini") {
		IniRead, PowerTools_ConnectionsRootUrl, PowerTools.ini, Main, ConnectionsRootUrl
		If (PowerTools_ConnectionsRootUrl="ERROR")
			PowerTools_ConnectionsRootUrl = connext.conti.de
	} Else
		PowerTools_ConnectionsRootUrl = connext.conti.de
}

global PowerTools_DocRootUrl
If FileExist("PowerTools.ini") {
	IniRead, PowerTools_DocRootUrl, PowerTools.ini, Main, DocRootUrl
} Else
	PowerTools_DocRootUrl = https://connext.conti.de/blogs/tdalon/entry/


; AutoExecute Section must be on the top of the script
;#NoEnv
;SetWorkingDir %A_ScriptDir%
#Warn All, OutputDebug

GroupAdd, Explorer, ahk_class CabinetWClass         
GroupAdd, Explorer, ahk_class ExploreWClass         
GroupAdd, Explorer, ahk_exe FreeCommander.exe
GroupAdd, Explorer, ahk_exe TOTALCMD.EXE

GroupAdd, OpenLinks, ahk_exe outlook.exe
GroupAdd, OpenLinks, ahk_exe powerpoint.exe
GroupAdd, OpenLinks, ahk_exe onenote.exe
GroupAdd, OpenLinks, ahk_exe word.exe
GroupAdd, OpenLinks, ahk_exe winword.exe
GroupAdd, OpenLinks, ahk_exe Teams.exe
GroupAdd, OpenLinks, ahk_exe lync.exe ; Skype


GroupAdd, MSOffice, ahk_exe outlook.exe
GroupAdd, MSOffice, ahk_exe powerpoint.exe
GroupAdd, MSOffice, ahk_exe POWERPNT.exe
GroupAdd, MSOffice, ahk_exe onenote.exe
GroupAdd, MSOffice, ahk_exe word.exe
GroupAdd, MSOffice, ahk_exe WINWORD.exe
GroupAdd, MSOffice, ahk_exe excel.exe
GroupAdd, MSOffice, ahk_exe teams.exe
GroupAdd, MSOffice, ahk_exe lync.exe

GroupAdd, NoIntelliPasteIns, ahk_exe XMind.exe
GroupAdd, NoIntelliPasteIns, ahk_exe freemind.exe

; Excel does not work

#SingleInstance force ; for running from editor - avoid warning another instance is running

Config := PowerTools_GetConfig()

SubMenuSettings := PowerTools_MenuTray()
Menu,Tray,Insert,Settings,Power Tools Bundler, PowerTools_RunBundler

; -------------------------------------------------------------------------------------------------------------------
; SETTINGS
Menu, SubMenuSettings, Add,Set Password, Login_SetPassword
Menu, SubMenuSettings, Add, Notification at Startup, MenuCb_ToggleSettingNotificationAtStartup

RegRead, SettingNotificationAtStartup, HKEY_CURRENT_USER\Software\PowerTools, NotificationAtStartup
If (SettingNotificationAtStartup = "")
	SettingNotificationAtStartup := True ; Default value
If (SettingNotificationAtStartup) {
  Menu, SubMenuSettings, Check, Notification at Startup
} Else {
  Menu, SubMenuSettings, UnCheck, Notification at Startup
}

; IntelliPaste Hotkey setting
Menu, SubMenuSettingsIntelliPaste, Add, &Hotkey, IntelliPaste_HotkeySet
Menu, SubMenuSettingsIntelliPaste, Add, &Refresh Teams List and SPSync.ini, IntelliPaste_Refresh
Menu, SubMenuSettingsIntelliPaste, Add, Help, IntelliPaste_Help
Menu, SubMenuSettings, Add, IntelliPaste, :SubMenuSettingsIntelliPaste

Menu, SubMenuSettings, Add, Phone Number, SetTelNr
Menu, SubMenuSettings, Add, Teams PowerShell, MenuCb_ToggleSettingTeamsPowerShell
RegRead, TeamsPowerShell, HKEY_CURRENT_USER\Software\PowerTools, TeamsPowerShell
If (TeamsPowerShell) 
	Menu,SubMenuSettings,Check, Teams PowerShell
Else 
  	Menu,SubMenuSettings,UnCheck, Teams PowerShell

; -------------------------------------------------------------------------------------------------------------------
; Setting - IntelliPasteHotkey
RegRead, IntelliPasteHotkey, HKEY_CURRENT_USER\Software\PowerTools, IntelliPasteHotkey
If ErrorLevel { ; regkey not set-> take default
	IntelliPasteHotkey = Insert
	PowerTools_RegWrite("IntelliPasteHotkey",IntelliPasteHotkey)
}

If (IntelliPasteHotkey == "Insert") {
	Hotkey, IfWinNotActive, ahk_group NoIntelliPasteIns
	Hotkey, %IntelliPasteHotkey%, IntelliPaste, On 
	Hotkey, IfWinNotActive,
} Else
	Hotkey, %IntelliPasteHotkey%, IntelliPaste, On 
         

; -------------------------------------------------------------------------------------------------------------------
; Tooltip
If !a_iscompiled {
	FileGetTime, LastMod , %A_ScriptFullPath%
} Else {
	LastMod := LastCompiled
}
FormatTime LastMod, %LastMod% D1 R
sTooltip = NWS PowerTool %LastMod%.`nRight-Click on icon to access Help and Settings.
sTrayTip = Right-Click on icon to access Help.
Menu, Tray, Tip, %sTooltip%

If (SettingNotificationAtStartup)
	TrayTip NWS PowerTool is running! , %sTrayTip%

; -------------------------------------------------------------------------------------------------------------------
; Add Custom Menus to MenuTray
Menu,Tray,NoStandard
If (Config = "Conti") {
	Menu,SubMenuFavs,Add, Open NWS Search, PowerTools_NWSSearch
	Menu,SubMenuFavs,Add, Create Ticket (ESS), SysTrayCreateTicket
	Menu,SubMenuFavs,Add, KSSE, KSSE
}
Menu,SubMenuFavs,Add, Cursor Highlighter, PowerTools_CursorHighlighter
Menu, Tray, Add, Tools, :SubMenuFavs

Menu,Tray,Add, Toggle AlwaysOnTop (Ctrl+Shift+Space), SysTrayToggleAlwaysOnTop
Menu,Tray,Add, Toggle Title Bar, SysTrayToggleTitleBar
Menu, SubMenuODB, Add, Open Permissions Settings,ODBOpenPermissions
Menu, SubMenuODB, Add, Open Document Library in Classic View,ODBOpenDocLibClassic
Menu, Tray, Add, OneDrive, :SubMenuODB


Menu, SubMenuODM, Add, ODM Set Path, ODMSetPath
Menu, SubMenuODM, Add, ODM Edit, ODMEdit
Menu, SubMenuODM, Add, ODM Run, ODMRun

Menu, SubMenuODM, Add, ODM AutoStart, MenuCb_ToggleODMAutoStart
RegRead, ODMAutoStart, HKEY_CURRENT_USER\Software\PowerTools, ODMAutoStart
If (ODMAutoStart) 
  	Menu,SubMenuODM,Check, ODM AutoStart
Else 
  Menu,SubMenuODM,UnCheck, ODM AutoStart

Menu, Tray, Add, OneDrive Mapper, :SubMenuODM

If ODMAutoStart {
	RegRead, ODMPath, HKEY_CURRENT_USER\Software\PowerTools, ODMPath
    If ODMPath {
		RunWait, PowerShell.exe -ExecutionPolicy Bypass -Command %ODMPath% ,, Hide
	} Else {
		TrayTipAutoHide("ODM Wrong Setup","ODM AutoStart is set but .ps1 file is not set.")
	}
}

Menu,Tray,Add ; Separator
Menu,Tray,Standard
; -------------------------------------------------------------------------------------------------------------------
; NWS Menu (Shown with Win+F1 hotkey in Browser)
Menu, NWSMenu, add, (Browser) Intelli &Copy current Url (Ctrl+Shift+C), IntelliCopyActiveUrl
Menu, NWSMenu, add, (Browser) Share by E&mail current Url (Ctrl+Shift+M), EmailShareActiveUrl
Menu, NWSMenu, add, (Browser) Create IT &Ticket, CreateTicket
Menu, NWSMenu, add, (Browser) Quick &Search (Win+F), QuickSearch
Menu, NWSMenu, add, (Browser) Share Url to Teams, TeamsShareActiveUrl
; -------------------------------------------------------------------------------------------------------------------

; EDIT : SCRIPT PARAMETERS
DefExplorerExe := "explorer.exe" ;*[NWS]
;DefExplorerExe := "D:\DSUsers\uid41890\PortableApps\FreeCommanderPortable\FreeCommanderPortable.exe"
IfNotExist, %DefExplorerExe%
	DefExplorerExe := "explorer.exe"
;DefConnextEditorBrowserExe := "chrome.exe" not used because handled by BrowserSelect

If (A_UserName = "uid41890") { ; only for me
	;SetCapsLockState , AlwaysOff
	SetNumLockState , AlwaysOn
}

; Start VPN (only for Conti config)
If (Config = "Conti") {
	If ! (Login_IsContiNet())
		VPNConnect()
}

;================================================================================================
;  CapsLock processing.  Must double tap CapsLock to toggle CapsLock mode on or off.
; https://www.howtogeek.com/446418/how-to-use-caps-lock-as-a-modifier-key-on-windows/
;================================================================================================
; Must double tap CapsLock to toggle CapsLock mode on or off.
CapsLock:: ; <--- Must double tap CapsLock to toggle CapsLock mode on or off.
    KeyWait, CapsLock                                                   ; Wait forever until Capslock is released.
    KeyWait, CapsLock, D T0.2                                           ; ErrorLevel = 1 if CapsLock not down within 0.2 seconds.
    if ((ErrorLevel = 0) && (A_PriorKey = "CapsLock") )                 ; Is a double tap on CapsLock?
        {
        SetCapsLockState, % GetKeyState("CapsLock","T") ? "Off" : "On"  ; Toggle the state of CapsLock LED
        }
return


^+Space:: WinSet, AlwaysOnTop, Toggle, A
return
;================================================================================================
; Hotkeys with CapsLock modifier.  See https://autohotkey.com/docs/Hotkeys.htm#combo
;================================================================================================

CapsLock & d:: ;  <--- Get DEFINITION of selected word.
sSelection:= GetSelection()
Run, http://www.google.com/search?q=define+%sSelection%     ; Launch with contents of clipboard
Return

CapsLock & b:: ;  <--- Bing search.
sSelection:= GetSelection()
If RegExMatch(sSelection,"(.*), (.*) <(.*)>",sMatch) ; From Outlook contact
	sSelection = %sMatch2% %sMatch1%#,Person ; Transform Firstname Lastname
Run, http://www.bing.com/search?q=%sSelection%     ; Launch with contents of clipboard
Return


CapsLock & g:: ; <--- GOOGLE the selected text.
sSelection:= GetSelection()
Run, http://www.google.com/search?q=%sSelection%             ; Launch with contents of clipboard
Return

;CapsLock & t:: ; <--- Do THESAURUS of selected text
;sSelection:= GetSelection()
;Run http://www.thesaurus.com/browse/%sSelection%             ; Launch with contents of clipboard
;Return


CapsLock & w:: ; <--- Do WIKIPEDIA of selected text
sSelection:= GetSelection()
Run, https://en.wikipedia.org/wiki/%sSelection%              ; Launch with contents of clipboard
Return

CapsLock & n:: ; <--- Open NWS Search or trigger NWS Search with selected text
PowerTools_NWSSearch()
return

CapsLock & y:: ; <--- YouTube search of selected text
sSelection:= GetSelection()
Run, https://www.youtube.com/results?search_query=%sSelection%              ; Launch with contents of clipboard
Return


; -------------------------------------------------------------------------------------------------------------------
;   All Applications
; -------------------------------------------------------------------------------------------------------------------
; Ctrl+Alt+V
^!v:: ; <--- VPN Connect
VPNConnect()
return

; -------------------------------------------------------------------------------------------------------------------
; Ctrl+F12
^F12:: ; <--- Paste clean url with url decoded
; Bypass for Lotus Notes/ Sametime: always encode/ no broken links
If WinActive("ahk_exe notes2.exe")
	PasteCleanUrl(false)	
Else
	PasteCleanUrl(true)
return	

; -------------------------------------------------------------------------------------------------------------------
; Ctrl+Ins: paste clean url without uridecode - unbroken link
; useful to paste unbroken link e.g. in ConNext comments
^Ins:: ; <--- Paste Clean Url
PasteCleanUrl(false)	
return	

; -------------------------------------------------------------------------------------------------------------------
#IfWinActive, ahk_group OpenLinks
; Open in Default Browser (incl. Office applications) - see OpenLink function
; Middle Mouse Click
MButton:: ; <--- Open in Preferred Browser
; If target window is not under focus, e.g. MButton on Chrome Tab
Clip_All := ClipboardAll  ; Save the entire clipboard to a variable
Clipboard =  ; Empty the clipboard to allow ClipWait work

sleep, 200 ;(wait in ms) leave time to release the Shift 
SendEvent {RButton} ;Click Right does not work in Outlook embedded tables
sleep, 200 ;(wait in ms) give time for the menu to popup	

If WinActive("ahk_exe onenote.exe")
	Sendinput i ; Copy Link
Else
	Sendinput c ; Copy Link

ClipWait, 2

sURL := Clipboard

If sURL ; Not empty
	OpenLink(sURL)
Else {
	;MsgBox OpenLinks: Empty URL/ Error
	Send {MButton}
}		

Clipboard := Clip_All ; Restore the original clipboard
return


; -------------------------------------------------------------------------------------------------------------------
;   Exclude Excel
; -------------------------------------------------------------------------------------------------------------------

; OpenIssue Exclude Excel
; Ctrl+Shift+I
#IfWinNotActive, ahk_exe EXCEL.EXE
^+i:: ; <--- Open Issue (LL, RAI, Jira, IMS)
ClipSaved := ClipboardAll

; Try first if selection is manually set (Triple click doesn't work in Outlook #35)
Clipboard := ""		;Clear the clipboard
Send, ^c			;Copy (Ctrl+C)
ClipWait, 2			;Wait for the clipboard to contain data
if (ErrorLevel){	;If ClipWait failed
	return
}

If (clipboard = "")
{
	Click 3 ; Click 2 won't get the word because "-" split it. Select the line
	; Does not work in Outlook
	Clipboard := ""		;Clear the clipboard
	Send, ^c			;Copy (Ctrl+C)
	ClipWait, 2			;Wait for the clipboard to contain data
	if (ErrorLevel){	;If ClipWait failed
		return
	}
}

sWord := clipboard

Click ; Remove word selection
Clipboard := ClipSaved ; restore clipboard

If RegExMatch(sWord,"\s*LL\-([\d]{1,})",sId) {
	sURL = https://aws1.conti.de/sites/LessonsLearned/Lists/LL/DispForm.aspx?ID=%sId1%
	Run, iexplore.exe %sURL%
}
Else If RegExMatch(sWord,"\s*RAI\-([\d]{1,})",sId) {
	sURL = https://aws1.conti.de/sites/LessonsLearned/Lists/RAI/DispForm.aspx?id=%sId1%
	Run, iexplore.exe %sURL%
}	
; JIRA Issue Keys
; JIRA IT EA: Implemented Project Keys: CFTI, CFTPT
Else If RegExMatch(sWord,"\s*/(CFTI|CFTPT)\-([\d]{1,})",sKeys) {
	sURL = http://it-ea-agile.conti.de:7090/browse/%sKeys1%-%sKeys2%
	Run, "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --profile-directory="Profile 3" %sURL%
}
; JIRA I BS: implemented project keys: EEXCTCV, EEXMGT
Else If RegExMatch(sWord,"\s*/(EEXCTCV|EEXMGT)\-([\d]{1,})",sKeys) {
	sURL = http://jira-bs-1:8080/browse/%sKeys1%-%sKeys2%
	Run, "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --profile-directory="Profile 3" %sURL%
}
Else If RegExMatch(sWord,"\s*/(JSDCTCV)\-([\d]{1,})",sKeys) {
	sURL = http://jira-bs-1.cw01.contiwan.com:8080/servicedesk/customer/portal/1/%sKeys1%-%sKeys2%
	;http://jira-bs-1.cw01.contiwan.com:8080/servicedesk/customer/portal/1/JSDCTCV-538
	Run, "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --profile-directory="Profile 3" %sURL%
}

Else If RegExMatch(sWord,"\s*IMS\-([\d]{1,}).*",sId) {
	sURL = integrity://ims-bs:7002/im/viewissue?selection=%sId1%
	Run, %sURL%
}
; HPSM ESS SD Tickets
Else If RegExMatch(sWord,"\s*(SD[\d]{1,}).*",sId) {
	sURL = https://hpsm-web.cw01.contiwan.com/sm/ess.do?ctx=docEngine&file=incidents&query=incident.id="%sId1%"
	sURL := StrReplace(sURL,"""","%22")
	Run, %sURL%
}
; Default IMS	
Else If RegExMatch(sWord,"\s*([\d]{1,}).*",sId) {
	sURL = integrity://ims-bs:7002/im/viewissue?selection=%sId1%	
	Run, %sURL%
}
return

; Ctrl+Shift+V  Paste in plain text/ removing rich-text formatting like links
; http://stackoverflow.com/a/132826/2043349
; https://lifehacker.com/better-paste-takes-the-annoyance-out-of-pasting-formatt-5388814
; Exclude for Excel for AddIn Ctrl+Shift+V: open IMS issue in document - use MS Office paste option instead
#IfWinNotActive, ahk_exe EXCEL.EXE
^+v:: ; <--- Paste in plain text format
WinClip.Paste(Clipboard)
return

; -------------------------------------------------------------------------------------------------------------------
;   BROWSER Group
; -------------------------------------------------------------------------------------------------------------------
#IfWinActive ahk_group Browser
/*
	; Ctrl + Alt + V - remove quotes for MySuccess copy/paste of goals from Excel
	#IfWinActive,ahk_group Browser
	^!v:: 
	ClipSaved := ClipboardAll
	sURL := clipboard
	sURL := StrReplace(sURL,"""","")
	;MsgBox Clean url:`n%sURL%
	clipboard = ; Empty the clipboard
	Clipboard := sURL
	ClipWait, 0.5
	Send ^v
	Sleep 100 ; pause necessary because of lag in browser (no problem in Notepad e.g.)- next command restore clipboard runs asynchron before paste
	; https://autohotkey.com/board/topic/37029-good-practices-with-clipboard/#entry233156	
	Clipboard := ClipSaved ; restore clipboard
	
	return
*/

; Win+F 
#f:: ; <--- [Browser] Run Quick Search (ConNext, Confluence, Jira)
QuickSearch()
return


; Ctrl+E - like Explorer or Edit - from Browser
; Do not use Alt key because of issue with IE
^e:: ; <--- [Browser] Edit ConNext or Open SharePoint in File Explorer
;!e:: ; Alt+E because Ctrl+E is used and can not be overwritten with Windows 10 and IE/Edge Browser Universal App
sURL := GetActiveBrowserUrl()
If Connections_IsConnectionsUrl(sURL) { 
	Connections_Edit(sURL)
	return
} Else If SharePoint_IsSPUrl(sURL) { ; SharePoint
	newurl:= SharePoint_CleanUrl(sURL) ; returns wihout ending /
	; For o365 SharePoint check if file is synced in SPsync.ini
	If InStr(newurl,"continental.sharepoint.com") or InStr(newurl,"https://mspe.conti.de/") { ; mspe can also offers Sync
		EnvGet, sOneDriveDir , onedrive
		sOneDriveDir := StrReplace(sOneDriveDir,"OneDrive - ","")
		sIniFile = %sOneDriveDir%\SPsync.ini

		If Not FileExist(sIniFile)
		{
			TrayTip, NWS PowerTool, File %sIniFile% does not exist! File was created in "%sOneDriveDir%". Fill it following user documentation.

			FileAppend, REM See documentation https://connext.conti.de/blogs/tdalon/entry/onedrive_sync_ahk#Setup`n, %sIniFile%
			FileAppend, REM Use a TAB to separate local root folder from SharePoint root url`n, %sIniFile%
			FileAppend, REM Replace #TBD by the SharePoint root url. Url shall not end with /`n, %sIniFile%
			FileAppend, %syncDir%%A_Tab%#TBD`n, %sIniFile%
			Run https://connext.conti.de/blogs/tdalon/entry/onedrive_sync_ahk#Setup
			sCmd = Edit "%sIniFile%"
			Run %sCmd%
    		return
		}

		If FileExist(sIniFile){
			If RegExMatch(newurl,".*/(?:Shared|Project) Documents",rooturl) { ; ?: non capturing group
				;MsgBox %newurl% %rooturl%
				needle := "(.*)\t" rooturl "(.*)"
				needle := StrReplace(needle," ","(?:%20| )")
				Loop, Read, %sIniFile%
				{
				If RegExMatch(A_LoopReadLine, needle, match) {	
					;MsgBox %rooturl% 1: %match1% 2: %match2%		
					sFile := StrReplace(newurl, rooturl . match2,Trim(match1) ) ; . "/"
					sFile := StrReplace(sFile, "/", "\")
					Run %DefExplorerExe% "%sFile%"
					; MsgBox %A_LoopReadLine% %needle% 
					return 
					}
				}
			}
		}
	}
	; new SharePoint open in file explorer/ Dav does not work
	If InStr(newurl,"continental.sharepoint.com") {
		TrayTipAutoHide("NWS PowerTool","SharePoint is not Sync'ed or OneDrive SPSync.ini File is not properly configured! (WebDav access does not work for newer SharePoint.)",3,0x3)
		sCmd = Edit "%sIniFile%"
		Run %sCmd%
		return
	}
	; Old and new SharePoints
	newurl:=StrReplace(newurl,"https:","")
	newurl:=StrReplace(newurl,"+"," ") ; strange issue with blank converted to +
	newurl:=StrReplace(newurl,"/","\")
	newurl:=StrReplace(newurl,"conti.de","conti.de@ssl\DavWWWroot")	; without @ssl it takes too long to open
	newurl:=StrReplace(newurl,"continental.sharepoint.com","continental.sharepoint.com@SSL\DavWWWroot")
	Run %DefExplorerExe% "%newurl%" 	
}
Else {
	TrayTipAutoHide("NWS PowerTool",sURL . " did not match a ConNext or SharePoint url!",,0x2)
}
return


#F1:: ; <--- [Browser] Open NWS PowerTool Menu
; Win+F1
Menu, NWSMenu, Show
return

; IntelliCopyActiveURL Ctrl+Shift+C
#IfWinActive,ahk_group Browser

^+c:: ; <--- [Browser] Intelli Copy Active Url
IntelliCopyActiveUrl:
If GetKeyState("Ctrl") and !GetKeyState("Shift") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/intelli_copy_active_url"
	return
}
sLink := GetActiveBrowserUrl()
sLink := CleanUrl(sLink,false)
WinGetActiveTitle, linktext
; Remove trailing - containing program e.g. - Google Chrome
StringGetPos,pos,linktext,%A_space%-,R
if (pos >=0)
	linktext := SubStr(linktext,1,pos)
If Connections_IsConnectionsUrl(sLink,"blog") { ; ConNext Blog
	linktext := StrReplace(linktext,"Blog Blog","Blog")
}
Else If Jira_IsUrl(sLink){
	; [CFTI-276] Document the metrics produced by hadoop - jira.conti.de
	StringGetPos,pos,linktext,%A_space%-,R
	if (pos >=0)
		linktext := SubStr(linktext,1,pos)
}
sHTML =	<a href="%sLink%">%linktext%</a>   
;MsgBox %linktext%
WinClip.SetHTML(sHTML)
TrayTipAutoHide("NWS PowerTool","Link was copied to the clipboard!")

return

; IntelliSharebyEmailActiveURL Ctrl+Shift+M
#IfWinActive, ahk_group Browser

^+m:: ; <--- [Browser] Share by eMail active url
EmailShareActiveUrl:
If GetKeyState("Ctrl") and !GetKeyState("Shift") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/share_link_by_email"
	return
}
sLink := GetActiveBrowserUrl()
sLink := CleanUrl(sLink,false)
WinGetActiveTitle, linktext
; Remove trailing - containing program e.g. - Google Chrome
StringGetPos,pos,linktext,%A_space%-,R
if (pos != -1)
	linktext := SubStr(linktext,1,pos)

If Connections_IsConnectionsUrl(sLink,"blog")  { ; ConNext Blog
	linktext := StrReplace(linktext,"Blog Blog","Blog")
}
sHTMLBody = Hello<br>I thought you might be interested in this post: <a href="%sLink%">%linktext%</a>.<br>
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

return


; -------------------------------------------------------------------------------------------------------------------
;     #CHROME BROWSER
; -------------------------------------------------------------------------------------------------------------------
#IfWinActive ahk_exe chrome.exe
;https://autohotkey.com/board/topic/84792-opening-a-link-in-non-default-browser/
; Shift+ middle mouse button
+MButton:: ;  <--- [Chrome] Open Link in IE

Clip_All := ClipboardAll  ; Save the entire clipboard to a variable
Clipboard := ""  ; Empty the clipboard to allow ClipWait work

Click Right ; Click Right mouse button
sleep, 100 ;(wait in ms) give time for the menu to popup
Sendinput e ; Send the underlined key that copies the link from the right click menu. see https://productforums.google.com/forum/#!topic/chrome/CPi4EmhqHPE
ClipWait, 2

sURL := Clipboard
If (sURL = "") {
	Exit
}

Run, iexplore.exe %sURL%
;Sleep 1000
;WinActivate, ahk_exe, iexplore.exe
;Run, C:\Users\%A_UserName%\AppData\Local\Google\Chrome\Application\chrome.exe %clipboard% ; Open in Google Chrome
;Run, %A_ProgramFiles%\Mozilla Firefox\firefox.exe %clipboard%; open in Firefox

Clipboard := Clip_All ; Restore the original clipboard
return

; Shift Mouse click
+LButton:: ; <--- [Chrome] Open link in preferred browser
; Calls OpenLink

Clip_All := ClipboardAll  ; Save the entire clipboard to a variable
Clipboard := ""  ; Empty the clipboard to allow ClipWait work

Click Right ; Click Right mouse button
sleep, 100 ;(wait in ms) give time for the menu to popup
SendInput e ; Send the underlined key that copies the link from the right click menu. see https://productforums.google.com/forum/#!topic/chrome/CPi4EmhqHPE
ClipWait, 2

sURL := Clipboard
If (sURL = "") {
	Exit
}

sURL := CleanUrl(sURL)
OpenLink(sURL)

Clipboard := Clip_All ; Restore the original clipboard
return

; Ctrl+Right mouse button
^RButton:: ; <--- [Chrome] Open link in File Explorer

Clip_All := ClipboardAll  ; Save the entire clipboard to a variable
Clipboard := ""  ; Empty the clipboard to allow ClipWait work

Click Right ; Click Right mouse button
sleep, 100 ;(wait in ms) give time for the menu to popup
SendInput e ; Send the underlined key that copies the link from the right click menu. see https://productforums.google.com/forum/#!topic/chrome/CPi4EmhqHPE
ClipWait, 2

sURL := Clipboard
If !sURL
	Exit

If SharePoint_IsSPUrl(sURL) {
	newurl:=SharePoint_CleanUrl(sURL)
	newurl:=StrReplace(newurl,"https:","")
	newurl:=StrReplace(newurl,"+"," ") ; strange issue with blank converted to +
	newurl:=StrReplace(newurl,"/","\")
	newurl:=StrReplace(newurl,"conti.de","conti.de@ssl\DavWWWroot")	; without @ssl it takes too long to open
	Run %DefExplorerExe% "%newurl%" 
}	

Clipboard := Clip_All ; Restore the original clipboard
return


; -------------------------------------------------------------------------------------------------------------------
;    EXPLORER Group
; -------------------------------------------------------------------------------------------------------------------
#ifWinActive,ahk_group Explorer              ; Set hotkeys to work in explorer only
; Open file With Notepad++ from Explorer using Alt+N hotkey
; https://autohotkey.com/board/topic/77665-open-files-with-portable-notepad/


; -------------------------------------------------------------------------------------------------------------------
; Alt+N 
!n:: ; <--- [Explorer] Open file in Notepad++
ClipSaved := ClipboardAll
Clipboard := ""
SendInput ^c
ClipWait, 0.5
file := Clipboard
Clipboard := ClipSaved
Run, notepad++.exe "%file%" 
return

; -------------------------------------------------------------------------------------------------------------------
; Override Delete key for Sync location
$Del:: ; <--- [Explorer] Safeguard Delete if in  ODB Sync location
sFile := Explorer_GetSelection()
; if no file selected in File Explorer
If (!sFile) ; file empty
	return
EnvGet, sOneDriveDir , onedrive
sOneDriveDir := StrReplace(sOneDriveDir,"OneDrive - ","")	
If InStr(sFile,sOneDriveDir . "\") { 
	MsgBox 0x14, Delete?,Are you sure you want to delete in your Sync location?`nIt might also delete the file in the SharePoint / not only locally for you, if sync is active.
	IfMsgBox, No
		return
}
Send {Delete}
return
; -------------------------------------------------------------------------------------------------------------------
; Ctrl+E Open SharePoint File from mapped Document Library or Sync location in Default Browser
; Calls: GetFileLink, GetExplorerSelection
^e:: ;	<--- [Explorer] Open SharePoint file selection in IE Browser	
sFile := Explorer_GetSelection()
; if no file selected in File Explorer
If !sFile ; empty
{
	MsgBox "You need to select a file!" 		
	return
}

; For multi-section take the last one
sFile := RegExReplace(sFile,"`n.*","")

sFile := GetFileLink(sFile)
If (!sFile) ; file empty
	return
If InStr(sFile,"#TBD")
{
	sIniFile := SharePoint_GetSyncIniFile()
	TrayTip, NWS PowerTool, File %sIniFile% is not configured! Replace #TBD by the SharePoint root url.
	Run https://connext.conti.de/blogs/tdalon/entry/onedrive_sync_ahk#Setup
	sCmd = Edit "%sIniFile%"
	Run, %sCmd%
    return
}

SplitPath, sFile, OutFileName, OutDir	
If InStr(OutFileName,".")  ; then a file is selected (Last part in Path containing "." for file extension)
	Run, "%OutDir%" ; Open parent directory
Else
	Run, "%sFile%"

return
; -------------------------------------------------------------------------------------------------------------------
; Ctrl+O
; Calls: GetFileLink, GetExplorerSelection, OpenFile
^o:: ; <--- [Explorer] Open file
sFiles := Explorer_GetSelection()
; if no file selected in File Explorer
If sFiles =
{
	MsgBox "You need to select a file!" 		
	return
}

Loop, parse, sFiles, `n, `r
{
	sFile := A_LoopField	
	If InStr(sFile,".xlsx") or  InStr(sFile,".docx") or InStr(sFile,".pptx") or InStr(sFile,".xlsm") or  InStr(sFile,".docm") or InStr(sFile,".pptm") {
		sFile := GetFileLink(sFile)
		Run, iexplore.exe "%sFile%" ; BUG: Edge can not open file links
	} Else
		Run, Open "%sFile%"
}

return

; -------------------------------------------------------------------------------------------------------------------
; Ctrl+K
; Calls: GetFileLink, GetExplorerSelection, OpenFile
^k:: ; <--- [Explorer] Copy File Link (OneDrive)
Send +{F10} ; Shift+F10
Send s 
Send {Enter}
Sleep 2000 ; Time to load the UI
Send {Tab 3} 
Send {Enter}
Send {Esc}
return

; -------------------------------------------------------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------
; FUNCTIONS
; -------------------------------------------------------------------------------------------------------------------

; -------------------------------------------------------------------------------------------------------------------
;IsIELink(url)
; true if link shall be opened with Internet Explorer rather than another browser e.g. Chrome because of incompatibility
IsIELink(sURL){
	If SharePoint_IsSPUrl(sURL) || InStr(sURL,"file://") || InStr(sURL,"/pkit/") || InStr(sURL,"/BlobIT/") || InStr(sURL,"/openscapeuc/dial/") || InStr(sURL,"://tts.fr.ge.conti.de")
		return true
	Else 	
		return false	
}


; -------------------------------------------------------------------------------------------------------------------
PasteCleanUrl(decode){
	; encode: True/False
	; calls: CleanUrl
	; called by Hotkey Ctrl+Ins [decode=false] and Ctrl+F12 [decode=true]
	ClipSaved := ClipboardAll
	sURL := clipboard
	sURL := CleanUrl(sURL,decode)
	
	;MsgBox Clean url:`n%sURL%
	;sendInput % sURL
	
	clipboard = ; Empty the clipboard
	Clipboard := sURL
	ClipWait
	Send ^v
	Sleep 400 ; pause necessary because of lag in browser or Outlook (no problem in Notepad e.g.)- next command restore clipboard runs asynchron before paste
	; https://autohotkey.com/board/topic/37029-good-practices-with-clipboard/#entry233156	
	Clipboard := ClipSaved ; restore clipboard
	return
}


; -------------------------------------------------------------------------------------------------------------------
; Function OpenLink
; Open Link in Default browser
OpenLink(sURL) {
	If IsIELink(sURL) {
		;Run, %A_ProgramFiles%\Internet Explorer\iexplore.exe %sURL%
		Run, iexplore.exe "%sURL%"
		;Sleep 1000
		;WinActivate, ahk_exe, iexplore.exe
	} Else If IsConfluenceUrl(sURL)  || Jira_IsUrl(sURL)  { ; JIRA or Confluence urls							
		;Run, C:\Users\%A_UserName%\AppData\Local\Google\Chrome\Application\chrome.exe %sURL%
		Run, "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --profile-directory="Profile 3" "%sURL%"
	} Else If InStr(sURL,"https://teams.microsoft.com/") { ; Teams url=>open by default in App
		;Run, StrReplace(sURL, "https://","msteams://")
		Run, "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --profile-directory="Profile 4" "%sURL%"
	} Else {
		;sURL := CleanUrl(sURL) ; No need to clean because handled by Redirector
		Run, "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --profile-directory="Profile 1" "%sURL%"
	} ; End If
} ; End Function OpenLink

; ----------------------------------------------------------------------
QuickSearch(){
If GetKeyState("Ctrl") and !GetKeyState("Shift") {
	sUrl := "https://connext.conti.de/wikis/home/wiki/Wc4f94c47297c_42c8_878f_525fd907cb68/page/NWS%20PowerTool?section=%28Win%2BF%29_Quick_Search_%28ConNext%2C_Jira%2C_Confluence%29"
	Run, "%sUrl%"
	return
}
sUrl := GetActiveBrowserUrl()
If !sUrl { ; empty
	MsgBox Cannot get URL ; DBG
	return
}
If Connections_IsConnectionsUrl(sUrl) 
	ConnectionsSearch(sUrl)
Else If IsConfluenceUrl(sUrl) {
	ConfluenceSearch(sUrl)
} Else If Jira_IsUrl(sUrl)
	JiraSearch(sUrl)
}	
; ----------------------------------------------------------------------
SysTrayCreateTicket:
SendInput, !{Esc}
CreateTicket()
return

SysTrayToggleAlwaysOnTop:
SendInput, !{Esc}
WinSet, AlwaysOnTop, Toggle, A
return

SysTrayToggleTitleBar:
SendInput, !{Esc}
WinSet, Style, ^0xC00000, A ; toggle title bar
return

CreateTicket()
{
If GetKeyState("Ctrl") {
	sUrl := "https://connext.conti.de/blogs/tdalon/entry/connext2ticket_ahk"
	Run, "%sUrl%"
	return
}
sTel := GetTelNr()
sForumUrl := GetActiveBrowserUrl()
;sForumTitle := ConNextGetForumTitle(sForumUrl)
If !(sForumUrl =="") {
WinGetActiveTitle, sForumTitle

; Remove trailing - containing program e.g. - Google Chrome
StringGetPos,pos,sForumTitle,%A_space%-,R
if (pos != -1)
    sForumTitle := SubStr(sForumTitle,1,pos)
sDesc := RegExReplace(sForumTitle," - .*","")
; Extract Category from Forum Title
If RegExMatch(sForumTitle,"- (.*) Forum -",sForumName) {
    sCat = %sForumName1%
	sIniFile = %A_ScriptDir%\ConNext2HPSMCat.ini
	If FileExist(sIniFile) {
    	needle := sCat "\t(.*)"
		Loop, Read, %sIniFile%
 		{
			If RegExMatch(A_LoopReadLine, needle, sCat) {
				sCat := sCat1
                Break
			}
		}
	} Else {
		Switch sCat
		{
		Case "SharePoint Online":
			sCat = SharePoint Online Platform
		Case "SharePoint Transition":
			sCat = SharePoint Online Migration
		Case "MS Teams":
			sCat = Teams
		Case "ConNext","Conny ConNext Assistant suggestions":
			sCat = ConNext
		Case "Outlook":
			sCat = Outlook
		} ; End Switch
	}
} ; End If RegExMatch(sForumTitle,"- (.*) Forum -",sForumShortTitle)
If (sCat == "ConNext")
    sDetails = See details %sForumUrl%.`nPlease root directly to ConNext Support who has access to this page.
Else {
    MsgBox 0x1001, Print to PDF, Do you want to Print the Browser page to PDF? Press OK when ready.
	IfMsgBox OK
    	sDetails = See attached PDF for details. (Source %sForumUrl%)
}
} ; End if sForumUrl
;MsgBox %sCat%`n%sDetails% ; DBG

Run https://hpsm-web.cw01.contiwan.com/sm/ess.do
MsgBox 0x1001, Wait..., Wait for Submit a Trouble Ticket. Click on the Ticket Form 'Preferred Phone number'. Press OK when ready to be filled.
IfMsgBox Cancel
    return

SendInput {Raw} %sTel%
Send {Tab}
If sCat
	SendInput {Raw} %sCat%
Send {Tab 3}
SendInput {Raw} %A_ComputerName%
Send {Tab 2}

If sDesc
	SendInput {Raw} %sDesc%
Else
	return
Send {Tab}
SendInput {Raw} %sDetails%

} 

; ----------------------------------------------------------------------
GetTelNr(){
RegRead, sTel, HKEY_CURRENT_USER\Software\PowerTools, TelNr
If (sTel=""){
	sTel := SetTelNr()
	If (sTel="") ; cancel
		return
}
return sTel
}
; ----------------------------------------------------------------------
SetTelNr(){
If GetKeyState("Ctrl") {
	sUrl := "https://connext.conti.de/blogs/tdalon/entry/connext2ticket_ahk"
	Run, "%sUrl%"
	return
}
RegRead, sTel, HKEY_CURRENT_USER\Software\PowerTools, TelNr
InputBox, sTel, TelNr, Enter preferred Phone number for IT tickets, , 200, 150, , , , , %sTel%
If ErrorLevel
	return
PowerTools_RegWrite("TelNr",sTel)
return sTel
}

; ----------------------------------------------------------------------
KSSE(){
Run %A_AppData%\Microsoft\Windows\Start Menu\Programs\Special Applications\KSSE.appref-ms
WinWaitActive, Anmeldung

sPersNum := Login_GetPersonalNumber()
If (sPersNum="")
	return
SendInput %sPersNum%{Tab}

sPassword := Login_GetPassword()
SendInput %sPassword%
;Send,{Enter}

} ; eofun


; ----------------------------------------------------------------------
ODBOpenPermissions(){
If GetKeyState("Ctrl") {
	sUrl := "https://connext.conti.de/blogs/tdalon/entry/onedrive_open_persmissions_settings_powertool" 
	Run, "%sUrl%"
	return
}
OfficeUid := People_GetMyOUid()
Run https://continental-my.sharepoint.com/personal/%OfficeUid%_contiwan_com/_layouts/15/user.aspx	
}
; ----------------------------------------------------------------------
ODBOpenDocLibClassic(){
If GetKeyState("Ctrl") {
	sUrl := "https://connext.conti.de/blogs/tdalon/entry/onedrive_alert#Shortcut_/_PowerTool_way"
	Run, "%sUrl%"
	return
}
OfficeUid := People_GetMyOUid()
sUrl := "https://continental-my.sharepoint.com/personal/" . OfficeUid . "_contiwan_com/Documents/Forms/All.aspx?ShowRibbon=true&InitialTabId=Ribbon%2ELibrary&VisibilityContext=WSSTabPersistence"
Run, "%sUrl%"	
}

; ----------------------------------------------------------------------
TeamsShareActiveUrl:
If GetKeyState("Ctrl") and !GetKeyState("Shift") {
	Run, "https://tdalon.blogspot.com/share-to-teams"
	return
}
sLink := GetActiveBrowserUrl()
sLink := CleanUrl(sLink,false)
sLink = https://teams.microsoft.com/share?href=%sLink%
Run %sLink%
return

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

; ----------------------------------------------------------------------



MenuCb_ToggleODMAutoStart(ItemName, ItemPos, MenuName){
If GetKeyState("Ctrl") {
    sUrl := "https://connext.conti.de/blogs/tdalon/entry/OneDrive_Mapper"
    Run, "%sUrl%"
	return
}
RegRead, ODMAutoStart, HKEY_CURRENT_USER\Software\PowerTools, ODMAutoStart
ODMAutoStart := !ODMAutoStart
If (ODMAutoStart) {
 	RegRead, ODMPath, HKEY_CURRENT_USER\Software\PowerTools, ODMPath
    If !ODMPath {
		ODMPath := ODMSetPath()
		If ODMPath ; If no path entered cancel
			return
	}
	Menu,%MenuName%,Check, %ItemName%
	TrayTipAutoHide("OneDrive Mapper","OneDrive Mapper will auto-start with this script.")	

} Else {
    Menu,%MenuName%,UnCheck, %ItemName%	 
	TrayTipAutoHide("OneDrive Mapper","OneDrive Mapper auto-start was switched OFF.")
}
PowerTools_RegWrite("ODMAutoStart",ODMAutoStart)
}



ODMSetPath(){
If GetKeyState("Ctrl") {
    sUrl := "https://connext.conti.de/blogs/tdalon/entry/OneDrive_Mapper"
    Run, "%sUrl%"
	return
}
	
FileSelectFile, ODMPath , 1, OneDriveMapper.ps1, Browse for your OneDriveMapper.ps1 location
If (ODMPath = "") or !InStr(ODMPath,"OneDriveMapper.ps1") {
	TrayTipAutoHide("OneDrive Mapper Setup","OneDriveMapper.ps1 wasn't selected!")
	ODMPath =
} Else {
	PowerTools_RegWrite("ODMPath",ODMPath)
}
return ODMPath
}


ODMEdit(){
If GetKeyState("Ctrl") {
    sUrl := "https://connext.conti.de/blogs/tdalon/entry/OneDrive_Mapper"
    Run, "%sUrl%"
	return
}
RegRead, ODMPath, HKEY_CURRENT_USER\Software\PowerTools, ODMPath
If !ODMPath {
	TrayTipAutoHide("OneDrive Mapper Setup","OneDriveMapper.ps1 wasn't selected! Set ODM Path.")
	return
}
sCmd = Edit "%ODMPath%"
Run %sCmd%
}

ODMRun(){
If GetKeyState("Ctrl") {
    sUrl := "https://connext.conti.de/blogs/tdalon/entry/OneDrive_Mapper"
    Run, "%sUrl%"
	return
}
RegRead, ODMPath, HKEY_CURRENT_USER\Software\PowerTools, ODMPath
If ODMPath {
	RunWait, PowerShell.exe -ExecutionPolicy Bypass -Command %ODMPath% ,, Hide
} Else {
	TrayTipAutoHide("OneDrive Mapper Run","OneDriveMapper.ps1 wasn't selected! Set ODM Path.")
}
}

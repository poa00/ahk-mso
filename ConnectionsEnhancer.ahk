; Connections Enhancer
; See documentation here: https://github.com/tdalon/ahk/wiki/Connections-Enhancer
; Author: Thierry Dalon
; Code Project Documentation is available on GitHub here: https://github.com/tdalon/ahk
;  Source https://github.com/tdalon/ahk/blob/master/ConnectionsEnhancer.ahk
;
; Run in ConNext Editor in Edit Mode
; Hotkey: Win+C


; Includes: ConNextMentionsMenu.ahk, ConNextMentionsMenuCallbacks.ahk
; Calls Lib/UnHTM
;       see #Includes


; AutoExecute Section must be on the top of the script
;#NoEnv
LastCompiled = 20201126214328
SetWorkingDir %A_ScriptDir%

#Include <WinClipAPI>
#Include <WinClip>
#Include <Connections>
#Include <Teams>
#Include <People>  
; for GetEmailList and Emails2Chat
#Include <PowerTools>
#Include <WinActiveBrowser>

global PowerTools_ConnectionsRootUrl
PowerTools_ConnectionsRootUrl := Connections_GetRootUrl()
If !PowerTools_ConnectionsRootUrl {
	TrayTipAutoHide("Error!","Connections Root Url is not defined!",2000,3)
}

GroupAdd, ChromeBasedBrowser, ahk_exe chrome.exe
GroupAdd, ChromeBasedBrowser, ahk_exe vivaldi.exe
GroupAdd, ChromeBasedBrowser, ahk_exe msedge.exe	

global PT_wc
If !PT_wc
	PT_wc := new WinClip

SetTitleMatchMode 2
#SingleInstance force ; for running from editor

SubMenuSettings := PowerTools_MenuTray()

Config := PowerTools_GetConfig()
If (Config ="Conti") {
	ConnextEnhancerIcon := "HBITMAP:*" . Create_ConnextEnhancer_ico()
	Menu, Tray, Icon, %ConnextEnhancerIcon%
}

; -------------------------------------------------------------------------------------------------------------------
; SETTINGS
Menu, SubMenuSettings, Add,Set Connections Root Url, Connections_SetRootUrl
Menu, SubMenuSettings, Add,Set Password, Login_SetPassword

Menu, SubMenuSettings, Add, Notification at Startup, MenuCb_ToggleSettingNotificationAtStartup

SettingNotificationAtStartup := PowerTools_RegRead("NotificationAtStartup")
If (SettingNotificationAtStartup = "")
	SettingNotificationAtStartup := True ; Default value
If (SettingNotificationAtStartup) {
  Menu, SubMenuSettings, Check, Notification at Startup
}
Else {
  Menu, SubMenuSettings, UnCheck, Notification at Startup
}

Menu, SubMenuSettings, Add, TOC Style, Connections_SettingSetTocStyle

; -------------------------------------------------------------------------------------------------------------------

If !a_iscompiled {
	FileGetTime, LastMod , %A_ScriptFullPath%
} Else {
	LastMod := LastCompiled
}
FormatTime LastMod, %LastMod% D1 R
sTooltip = Connections Enhancer %LastMod%`nUse 'Win+C' to open the Connections Enhancer Menu.`nRight-Click on icon to access Help.
sTrayTip = Use 'Win+C' to open the Connections Enhancer Menu from a ConNext Page in the Browser.`nRight-Click on icon to access Help.
Menu, Tray, Tip, %sTooltip%
If (SettingNotificationAtStartup)
	TrayTip "Connections Enhancer is running!", %sTrayTip%
;TrayTipAutoHide("Connections Enhancer is running!",sText)
; -------------------------------------------------------------------------------------------------------------------

#include *i %A_ScriptDir%\ConNextMentionsMenu.ahk ; only include if it exists
; -> Output SubMenuMentions
;Menu, ConNextMenu, add, Groups, :ConNextGroupSubmenu

; Display Menus
Menu, ConnectionsEnhancerMenu, add, &Quote, Quote
Menu, ConnectionsEnhancerMenu, add, Quote2, Quote2
Menu, ConnectionsEnhancerMenu, add, QuoteImg, QuoteImg
Menu, ConnectionsEnhancerMenu, Add ; Separator

Menu, SubMenuBoxes, add, &Idea, Boxes
Menu, SubMenuBoxes, add, &Bug, Boxes
Menu, SubMenuBoxes, add, Brake, Boxes
Menu, SubMenuBoxes, add, &Speed, Boxes
Menu, SubMenuBoxes, add, &Cancel, Boxes

Menu, SubMenuBoxesTable, add, &IdeaTable, Boxes
Menu, SubMenuBoxesTable, add, &BugTable, Boxes
Menu, SubMenuBoxesTable, add, BrakeTable, Boxes
Menu, SubMenuBoxesTable, add, &SpeedTable, Boxes
Menu, SubMenuBoxesTable, add, &CancelTable, Boxes

Menu, ConnectionsEnhancerMenu, add, Boxes, :SubMenuBoxes
Menu, ConnectionsEnhancerMenu, add, &BoxesTable, :SubMenuBoxesTable

; Kudos
Menu, SubMenuKudosTable, add, Thank_you (Table), Kudos
Menu, SubMenuKudosTable, add, Awesome (Table), Kudos
Menu, SubMenuKudosTable, add, Great_Idea (Table), Kudos
Menu, SubMenuKudosTable, add, Visionary (Table), Kudos
Menu, SubMenuKudosTable, add, Made_My_Day (Table), Kudos
Menu, SubMenuKudosTable, add, Team_Player (Table), Kudos

Menu, SubMenuKudos, add, Thank_you, Kudos
Menu, SubMenuKudos, add, Awesome, Kudos
Menu, SubMenuKudos, add, Great_Idea, Kudos
Menu, SubMenuKudos, add, Visionary, Kudos
Menu, SubMenuKudos, add, Made_My_Day, Kudos
Menu, SubMenuKudos, add, Team_Player, Kudos

Menu, SubMenuKudosParent, add, Kudos (Centered), :SubMenuKudos
Menu, SubMenuKudosParent, add, Kudos (Table), :SubMenuKudosTable
Menu, ConnectionsEnhancerMenu, add, Kudos, :SubMenuKudosParent
; End Kudos

Menu,ConnectionsEnhancerMenu,Add ; Separator
Menu, SubMenuMsg, add, &Info, Msg
Menu, SubMenuMsg, add, &Success, Msg
Menu, SubMenuMsg, add, &Warning, Msg
Menu, SubMenuMsg, add, &Error, Msg
Menu, ConnectionsEnhancerMenu, add, &Message Boxes, :SubMenuMsg

Menu, ConnectionsEnhancerMenu, add,  Collapsible Section, MenuItemCollapsibleSection

Menu, ConnectionsEnhancerMenu, add, Buttons, AddButtons
Menu, ConnectionsEnhancerMenu, add, Add Table Images, AddTableImages
; Menu, ConnectionsEnhancerMenu, add, Buttons Rounded, AddButtonsRounded ; does not work
Menu, ConnectionsEnhancerMenu,Add ; Separator
Menu, ConnectionsEnhancerMenu, Add, Intelli Paste (Ins), MenuItemIntelliPaste
Menu, ConnectionsEnhancerMenu, Add, Embed Video (small), MenuItemEmbedVideoSmall
Menu, ConnectionsEnhancerMenu, Add ; Separator
Menu, ConnectionsEnhancerMenu, add, --- Requires textbox.io or switch to HTML Code ---, CEHelpEditor
Menu, SubMenuCleaner, Add, &Clean All (Win+Shift+C), MenuItemCleanAll
Menu, SubMenuCleaner, Add ; Separator
Menu, SubMenuCleaner, add, Clean-up links, MenuItemCleanLinks
Menu, SubMenuCleaner, add, Clean-up Html, MenuItemCleanCode
Menu, SubMenuCleaner, add, Auto-Fix Heading Ids, MenuItemAutoFixHeadingIds
Menu, SubMenuCleaner, add, Clean-up Tables, MenuItemCleanTable
Menu, ConnectionsEnhancerMenu, Add, Cleaner, :SubMenuCleaner

Menu, ConnectionsEnhancerMenu, Add, Format Images (Win+Shift+I), FormatImg
Menu, ConnectionsEnhancerMenu, Add, View HTML Source (Win+Shift+U), CNViewSource

; TOC
Menu, SubMenuToc, Add, Generate &TOC (Win+Shift+T), MenuItemGenerateToc
Menu, SubMenuToc, Add, (Wiki) Generate &SubPages TOC, MenuItemGenerateWikiToc
Menu, SubMenuToc, Add, &Refresh TOC (Win+F5), RefreshToc
Menu, SubMenuToc, Add, &Delete TOC, DeleteToc
Menu, SubMenuToc, Add, TOC Style, Connections_SettingSetTocStyle

Menu, ConnectionsEnhancerMenu, Add, &Table of Contents, :SubMenuToc

Menu, ConnectionsEnhancerMenu, Add, Update/Copy User Table, MenuItemUpdateUserTable
Menu, ConnectionsEnhancerMenu, Add, Expand Mentions With Profile &Picture, ExpandMentionsWithProfilePic
Menu, ConnectionsEnhancerMenu, Add, Center IFrames, MenuItemCenterIframe

If FileExist("ConNextMentionsMenu.ahk") {
	Menu, ConnectionsEnhancerMenu, Add ; Separator
	Menu, ConnectionsEnhancerMenu, Add, Mentions, :SubMenuMentions
}
	

; ####################################################################
; Menu specific for actions in read mode / not in edit mode

; Social Functions

Menu, SubMenuContactMentions, add, &Mentions (extract), MenuItemMentions2Mentions
Menu, SubMenuContactMentions, add, &Emails, MenuItemMentions2Emails
Menu, SubMenuContactMentions, add, &Teams Chat, MenuItemMentions2Chat


Menu, SubMenuContactEmails, add, &Emails (extract), MenuItemEmails2Emails
Menu, SubMenuContactEmails, add, &Mentions, MenuItemEmails2Mentions
Menu, SubMenuContactEmails, add, &Send Mentions, MenuItemEmails2SendMentions
Menu, SubMenuContactEmails, add, &Teams Chat, MenuItemEmails2Chat

Menu, SubMenuContact, add, &Mentions to, :SubMenuContactMentions
Menu, SubMenuContact, add, &Emails to, :SubMenuContactEmails

Menu, SubMenuContactLikers, add, &Emails, MenuItemLikers2Emails
Menu, SubMenuContactLikers, add, E&mail to Likers, MenuItemLikers2Email
Menu, SubMenuContactLikers, add, &Teams Chat, MenuItemLikers2Chat
Menu, SubMenuContactLikers, add, &Mentions, MenuItemLikers2Mentions

Menu, SubMenuEvent, add, &Meeting, Event2Meeting
Menu, SubMenuEvent, add, &Calendar, Event2Calendar
Menu, SubMenuEvent, add, &Teams Chat, MenuItemAttendees2Chat
Menu, SubMenuEvent, add, E&mail, Event2Email
Menu, SubMenuEvent, add, Emails, MenuItemAttendees2Emails

Menu, ConnectionsEnhancerReadMenu, add, &Contact, :SubMenuContact
Menu, ConnectionsEnhancerMenu,Add ; Separator
Menu, ConnectionsEnhancerMenu, add, &Contact, :SubMenuContact
Menu, ConnectionsEnhancerMenu, add, Personalize Mentions (Win+1), PersonalizeMentions

Menu, ConnectionsEnhancerReadMenu, add, (&Blog) Likers to, :SubMenuContactLikers
Menu, ConnectionsEnhancerReadMenu, add, &Event to, :SubMenuEvent

Menu, ConnectionsEnhancerReadMenu, add, (Profile Search) Copy &Mentions, MenuItemProfileSearchToMentions
Menu, ConnectionsEnhancerReadMenu, add, (Profile Search) Copy &Table of Mentions with Profile pictures, MenuItemCopyUserTable
Menu, ConnectionsEnhancerReadMenu, add, (Profile Search) Copy Emails, MenuItemProfileSearch2Emails
Menu, ConnectionsEnhancerReadMenu, Add ; Separator
Menu, ConnectionsEnhancerReadMenu, add, Create New... (Win+N), CNCreateNew
Menu, ConnectionsEnhancerReadMenu, add, Download Html, Connections_DownloadHtml
return


; ####################################################################
; Hotkeys

; Only in Browser Window
#IfWinActive,ahk_group Browser
; #If IsWinConNext()

#1:: ; <--- Personalize mentions
PersonalizeMentions()	
;SendInput ^{Backspace}{Backspace}{Space}
return

; ----------------------------------------------------------------------

;Win+Shift+I : for images
#+i::  ; <--- [Browser] Format images
FormatImg:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_quick_images_formatting"
	return
}
TransformHtml("Connections_FormatImg")
return

#+u::  ; <--- [Browser] View HTML Source
CNViewSource()
return

;Win+Shift+T 
#+t:: ; <--- [Browser] Generate TOC/ Table of Contents
GoTo, MenuItemGenerateToc 
return


;Win+F5
#F5:: ; <--- [Browser] Refresh TOC
RefreshToc()
return

;Win+N
#N:: ; <--- [Browser] Create new (Forum topic, Blog post or Wiki page)
CNCreateNew()
return

; Win+NumpadLeft
#NumpadLeft:: ; <--- [Browser] Previous page (Wiki)
sUrl := GetActiveBrowserUrl()
CNWikiGoTo(sUrl,"prev")
return

; Win+NumpadRight
#NumpadRight:: ; <--- [Browser] Next page (Wiki)
sUrl := GetActiveBrowserUrl()
CNWikiGoTo(sUrl,"next")
return

; Win+NumpadUp
#NumpadUp:: ; <--- [Browser] Parent page (Wiki)
sUrl := GetActiveBrowserUrl()
CNWikiGoTo(sUrl,"parent")
return

; ----------------------------------------------------------------------

; Only in Browser Window
;Win+C
#c:: ; <--- [Browser] Open Connections Enhancer Menu
;^RButton:: ; Ctrl+Right Mouse Click (Win key does not work) -> will also open normal context menu7 overlap
;MsgBox % IsWinConNextEdit()
sUrl := GetActiveBrowserUrl()
If IsWinConNextEdit(sUrl) 
{
	If Connections_IsConnectionsUrl(sUrl,"wiki-edit") 
		Menu, SubMenuToc, Enable, (Wiki) Generate &SubPages TOC
	Else
		Menu, SubMenuToc, Disable, (Wiki) Generate &SubPages TOC
	
	Menu, ConnectionsEnhancerMenu, Show
}

Else {
	If Connections_IsConnectionsUrl(sUrl,"blog") 
		Menu, ConnectionsEnhancerReadMenu, Enable, (&Blog) Likers to
	Else
		Menu, ConnectionsEnhancerReadMenu, Disable, (&Blog) Likers to

	If InStr(sUrl,"&eventInstUuid=") 
		Menu, ConnectionsEnhancerReadMenu, Enable, &Event to
	Else
		Menu, ConnectionsEnhancerReadMenu, Disable, &Event to

	If InStr(sUrl, PowerTools_ConnectionsRootUrl . "/profiles/html/") {
		Menu, ConnectionsEnhancerReadMenu, Enable, (Profile Search) Copy &Mentions 
		Menu, ConnectionsEnhancerReadMenu, Enable, (Profile Search) Copy &Table of Mentions with Profile pictures 
		Menu, ConnectionsEnhancerReadMenu, Enable, (Profile Search) Copy Emails
	}
	Else {
		Menu, ConnectionsEnhancerReadMenu, Disable, (Profile Search) Copy &Mentions 
		Menu, ConnectionsEnhancerReadMenu, Disable, (Profile Search) Copy &Table of Mentions with Profile pictures
		Menu, ConnectionsEnhancerReadMenu, Disable, (Profile Search) Copy Emails
	}

	Menu, ConnectionsEnhancerReadMenu, Show
}
return 	

#+c:: ; <--- [Browser] Clean all
GoSub MenuItemCleanAll
return


#IfWinActive, ahk_group ChromeBasedBrowser
; Ctrl+S - save
; Do not use Alt key because of issue with IE
^s:: ; <--- [Chrome] Save current post or page
;!e:: ; Alt+E because Ctrl+E is used and can not be overwritten with Windows 10 and IE/Edge Browser Universal App
sUrl := GetActiveBrowserUrl()
If Connections_IsConnectionsUrl(sUrl) { 
	Connections_Save(sURL)
	return
} Else {
	SendInput ^s
}

; ----------------------------------------------------------------------
CEHelpEditor:
Run, https://connext.conti.de/blogs/tdalon/entry/connext_enhancer#ConNext_Editor
return


; ----------------------------------------------------------------------
MenuItemIntelliPaste:
Run, "https://connext.conti.de/blogs/tdalon/entry/intelli_paste"
return
; ----------------------------------------------------------------------
MenuItemCollapsibleSection:
input := MultiLineInputBox("Enter text to be wrapped in a collapsible section. `nFirst line will become the summary. The rest the details.", "", "Connections Enhancer: Collapsible Section")
If ErrorLevel ; Cancel
	return
sPat = Us).*\n ; Line break in GUI Control are represented by linefeed instead of `r`n

If (RegExMatch(input,sPat,sMatch)) {
	summary := sMatch
	text := StrReplace(input,sMatch,"")
} else {
	summary:=input
}
sHtml = <details><summary><b>%summary%</b></summary>%text%</details>
ClipPasteHtml(sHtml)
return


; ----------------------------------------------------------------------
MenuItemEmbedVideoSmall:
InputBox, sLink, Paste Link (.mp4 or _player.html),,,,100
if ErrorLevel
	return		

If InStr(sLink,".mp4") {
	sHtml = <p dir="ltr" style="text-align: center;"><video controls="controls"  width="95_percent" tabindex="-1" class="HTML5Video"><source src="%sLink%"/></video></p>
	sHtml := StrReplace(sHtml,"_percent","%")
} 
Else If InStr(sLink,"_player.html") { ; Camtasia player file
	If WinActive("Blog")  
		sHtml = <p dir="ltr" style="text-align: center;"><iframe  allowfullscreen="true"  frameborder="0" src="%sLink%" width="700" height="394"></iframe></p>
	Else
		sHtml = <p dir="ltr" style="text-align: center;"><iframe  allowfullscreen="true"  frameborder="0" src="%sLink%" width="800" height="450"></iframe></p>
}
ClipPasteHtml(sHtml)	
return
; ----------------------------------------------------------------------


; ----------------------------------------------------------------------
; TransformHtml

MenuItemGenerateWikiToc:
sUrl := GetActiveBrowserUrl()
;If !RegExMatch(sUrl,".*/connext.conti.de/wikis/.*/edit") {
;	MSgBox 0x10, Connections Enhancer: Error, GenTocWiki Function only works for a Wiki page in Edit mode!
;	return
;}
sPat = s)<div id="toc-wiki.*?".*?>.*?</div>
If !RegExMatch(sHtml,sPat,sMatch) { ; insert new
	sInsLoc := GetInsLoc()
	If (sInsLoc = "cursor")
		SendInput $CURSORLOC$
}

sHtml := CNGetHtmlEditor()

If !sHtml {
	Send {Escape} 
	return
}
sMatch :=
sTocHtml := GenerateWikiTocSubPages(sUrl,sMatch,sMatch)
If !sTocHtml
	return

If RegExMatch(sHtml,sPat,sMatch)
	sHtml := RegExReplace(sHtml, sPat, sTocHtml)
Else {
	If (sInsLoc = "cursor") {
		sLoc = <p dir="ltr">$LOCMARKER$</p>
		sHtml := StrReplace(sHtml,sLoc,sTocHtml)
		sHtml := StrReplace(sHtml,"$CURSORLOC$",sTocHtml)
	} else if (sInsLoc = "top") { 
		sHtml := sTocHtml . "`n" . sHtml
	} else if (sInsLoc = "bottom") { 
		sHtml := sHtml . "`n" . sTocHtml
	}
}    	

ClipPasteText(sHtml)

; Select and Press the OK button by Tab
Send {Tab} 
Send {Tab} 
Send {Enter}

return

MenuItemGenerateToc:
sUrl := GetActiveBrowserUrl()
If Not IsWinConNextEdit(sUrl)
	return

sInsLoc := GetInsLoc()

If (sInsLoc = "cursor")
	SendInput $CURSORLOC$

sHtml := CNGetHtmlEditor() ; also select all

; Wiki page
If InStr(sUrl, PowerTools_ConnectionsRootUrl . "/wikis/") {
	sHtml := InsertWikiToc(sHtml,sInsLoc)
} Else { ; Other than wiki
	sMatch :=
	sHtml  := GenerateToc(sHtml,sMatch,sMatch,sInsLoc)
	If !sHtml {
		Send {Escape} 
		return
	}
}

ClipPasteText(sHtml)

; Select and Press the OK button by Tab
Send {Tab} 
Send {Tab} 
Send {Enter}

return

; ----------------------------------------------------------------------


; ----------------------------------------------------------------------
MenuItemCenterIframe:
	sHtml := CNGetHtmlEditor()
	sHtml  := CenterIframe(sHtml)
	ClipPasteText(sHtml)

	; Select and Press the OK button by Tab
	Send {Tab} 
	Send {Tab} 
	Send {Enter}

return
; ----------------------------------------------------------------------
MenuItemCleanAll:
	sHtml := CNGetHtmlEditor()
	sHtml := CleanLinks(sHtml)
	sHtml := CleanCode(sHtml)
	sHtml := AutoFixHeadingIds(sHtml)
	sHtml := CleanTable(sHtml)
	ClipPasteText(sHtml)

	; Select and Press the OK button by Tab
	Send {Tab} 
	Send {Tab} 
	Send {Enter}
return


MenuItemAutoFixHeadingIds:
	sHtml := CNGetHtmlEditor()
	sHtml := AutoFixHeadingIds(sHtml)
	ClipPasteText(sHtml)

	; Select and Press the OK button by Tab
	Send {Tab} 
	Send {Tab} 
	Send {Enter}
return


MenuItemCleanLinks:
	sHtml := CNGetHtmlEditor()
	sHtml := CleanLinks(sHtml)
	ClipPasteText(sHtml)

	; Select and Press the OK button by Tab
	Send {Tab} 
	Send {Tab} 
	Send {Enter}

return
; ----------------------------------------------------------------------
MenuItemCleanTable:
	sHtml := CNGetHtmlEditor()	
	sHtml := CleanTable(sHtml)
	ClipPasteText(sHtml)

	; Select and Press the OK button by Tab
	Send {Tab} 
	Send {Tab} 
	Send {Enter}

return
; ----------------------------------------------------------------------
; ----------------------------------------------------------------------
MenuItemCleanCode:
	sHtml := CNGetHtmlEditor()	
	sHtml := CleanCode(sHtml)
	ClipPasteText(sHtml)

	; Select and Press the OK button by Tab
	Send {Tab} 
	Send {Tab} 
	Send {Enter}

return
; ----------------------------------------------------------------------

Quote:		
	input := MultiLineInputBox("Enter text to be wrapped in a quote block:", "", "Connections Enhancer: Quote")
	If ErrorLevel ; Cancel
		return
	input := StrReplace(input,"`n","<br>")
	sHtml=<div style="background: rgb(255, 255, 255); border-width: 0em 0em 0em 0.3em; border-style: solid; border-color: rgb(255, 165, 0); padding: 0.2em 2em; border-image: none; width: auto; overflow: auto;" dir="ltr"> %input%</div>
	ClipPasteHtml(sHtml)
	
return

; ----------------------------------------------------------------------
Quote2:
	input := MultiLineInputBox("Enter text to be wrapped in a quote2 block:", "", "Connections Enhancer: Quote2")
	If ErrorLevel ; Cancel
		return

	input := StrReplace(input,"`n","<br>")
	sHtml = <div><span style="color:  rgb(255, 165, 0);">  <span style="font-size: 25px;"><strong>&#x201C</strong></span> %input%</span></div>
	ClipPasteHtml(sHtml)
		
return

; ----------------------------------------------------------------------
QuoteImg:
	input := MultiLineInputBox("Enter text to be wrapped in a quote block:", "", "Connections Enhancer: QuoteImg")
	If ErrorLevel ; Cancel
		return

	input := StrReplace(input,"`n","<br>")

	sHtml = <div dir="ltr" style="direction:ltr"><table border="1" cellpadding="0" cellspacing="0" style="border:1pt solid rgb( 163 , 163 , 163 );border-collapse:collapse;direction:ltr" summary="" title=""><tbody><tr><td style="padding:4pt;border:1pt solid rgb( 255 , 255 , 255 );width:0.667in;vertical-align:top;background-color:white">
    <p style="margin:0in;font-family:'calibri';font-size:11pt"><img src="https://connext.conti.de/files/form/anonymous/api/library/ffd279a0-2764-46bb-a8ec-b1aa3c713072/document/64fda1e5-6fe9-4fa5-b90a-8ab8255e977f/media/wind_sock.png"/></p>
    </td>
        <td style="padding:4pt;border:1pt solid rgb( 255 , 255 , 255 );width:9.857in;vertical-align:top;background-color:rgb( 255 , 204 , 153 )">
          <p style="margin:0in;color:black;font-size:21pt">&nbsp;</p>
          <p style="margin:0in;font-family:'calibri';font-size:21pt;font-style:italic">%input%</p>
          <p style="margin:0in;font-family:'calibri';font-size:9pt">Author</p>
        </td></tr></tbody></table></div>
		
	ClipPasteHtml(sHtml)
return


; ----------------------------------------------------------------------
AddTableImages:
input := MultiLineInputBox("Enter ImageLink1, ImageLink2 ENTER....", "", "Connections Enhancer: TableImages")
If ErrorLevel ; Cancel
	return
	
sHtml= <table border="0" style="border-color:rgb( 105 , 105 , 105 );width:95_percent;">
Loop, parse, input, `n, `r
{
	sTr =
	Loop, parse, A_LoopField, `, 
	{
		sSrc := Trim(A_LoopField)
		sTd = <td style="border-color:rgb( 105 , 105 , 105 )"><img width="95_percent" style="margin:0px auto;display:block" src="%sSrc%" /></td> 
		sTd := StrReplace(sTd,"_percent","%")
		sTr = %sTr%%sTd%
	}
	sHtml = %sHtml%<tr>%sTr%</tr>
}	; End Loop
sHtml = %sHtml%</table>	
	
ClipPasteHtml(sHtml)

return	
; ----------------------------------------------------------------------
; ----------------------------------------------------------------------
AddButtons:
	input := MultiLineInputBox("Enter Button1Title , Link ENTER Button2Title , Link ENTER.... `nN.B.: Save and Close to view the buttons.", "", "Connections Enhancer: Buttons")
	If ErrorLevel ; Cancel
		return
	
	sHtml=
	Loop, parse, input, `n, `r
{
	
sPat= (.*),\s*(.*)\s*
sRep = <a class="lotusBtn" href="$2">$1</a>
; also center
sBtn := RegExReplace(A_LoopField,sPat,sRep)
sHtml= %sHtml%%sBtn%
	
	
}	; End Loop
	
sHtml = <div class="lotusBtnContainer">%sHtml%</div>
	
ClipPasteHtml(sHtml)

return	
; ----------------------------------------------------------------------
; ----------------------------------------------------------------------
AddButtonsRounded:
	input := MultiLineInputBox("Enter Button1Title , Link ENTER Button2Title , Link ENTER.... N.B.: Save and Close to view the buttons.", "", "Connections Enhancer: Rounded Buttons")
	If ErrorLevel ; Cancel
		return
	
	sHtml=
	Loop, parse, input, `n, `r
{
	
sPat= (.*),(.*)
sRep = <a class="lotusBtn " href="$2">$1</a>
; also center
sBtn := RegExReplace(A_LoopField,sPat,sRep)
sHtml= %sHtml%%sBtn%
	
	
}	; End Loop
	
sHtml = <div class="lotusBtnContainer" style="border-radius: 10px;">%sHtml%</div>
ClipPasteHtml(sHtml)

return	
; ----------------------------------------------------------------------
Boxes:
	StringLower, Keyword, A_ThisMenuItem
	AddBox(Keyword) 
return
; ----------------------------------------------------------------------
Kudos:
	StringLower, Keyword, A_ThisMenuItem
	AddKudos(Keyword) 
return
; ----------------------------------------------------------------------
Msg:
	StringLower, Keyword, A_ThisMenuItem
	AddMsg(Keyword) 
return



; -------------------------------------------------------------------------------------------------------------------
; Contact Functions
MenuItemLikers2Email:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_enhancer_reach_out#From_a_Blog_Post"
	return
}
; Calls: CNGetTitle
sUrl := GetActiveBrowserUrl()
sEmailList := CNLikers2Emails(sUrl)
Try
	MailItem := ComObjActive("Outlook.Application").CreateItem(0)
Catch
	MailItem := ComObjCreate("Outlook.Application").CreateItem(0)

sTitle := CNGetTitle(sUrl)
;MailItem.BodyFormat := 2 ; olFormatHTML
sHTMLBody =	Hello<br>You have liked this post: <a href="%sUrl%">%sTitle%</a>.<br> 
MailItem.HTMLBody := sHTMLBody
MailItem.To := sEmailList
MailItem.Display ;Make email visible
return

MenuItemLikers2Emails:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_enhancer_reach_out#From_a_Blog_Post"
	return
}
sUrl := GetActiveBrowserUrl()
sEmailList := CNLikers2Emails(sUrl)
WinClip.SetText(sEmailList)
TrayTipAutoHide("Copy Emails", "Emails were copied to the clipboard!")
return

MenuItemLikers2Chat:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_enhancer_reach_out#From_a_Blog_Post"
	return
}
sUrl := GetActiveBrowserUrl()
sEmailList := CNLikers2Emails(sUrl)
Teams_Emails2ChatDeepLink(sEmailList)
return

MenuItemLikers2Mentions:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_enhancer_reach_out#From_a_Blog_Post"
	return
}
sUrl := GetActiveBrowserUrl()
sUrl := RegExReplace(sUrl,"\?.*","") ; clean-up url: remove section and lang tag
Likers2Mentions(sUrl)

return

; ----------------------------------------------------------------------
MenuItemMentions2Emails:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_enhancer_reach_out#From_Mentions"
	return
}
sHtml := CNGetHtml()
If !sHtml
	return

sEmailList := CNMentions2Emails(sHtml)
WinClip.SetText(sEmailList)
TrayTipAutoHide("Copy Emails", "Emails were copied to the clipboard!")
return

; ----------------------------------------------------------------------
MenuItemMentions2Mentions:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_enhancer_reach_out#From_Mentions"
	return
}
sHtml := CNGetHtml()
If !sHtml {
	TrayTipAutoHide("Copy Mentions", "No html found!")
	return
}
sHtml := CNMentions2Mentions(sHtml)
If !sHtml {
	TrayTipAutoHide("Copy Mentions", "No mentions found!")
	return
}
WinClip.SetHtml(sHtml)
TrayTipAutoHide("Copy Mentions", "Mentions were copied to the clipboard in RTF!")
return

; ----------------------------------------------------------------------
MenuItemMentions2Chat:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_enhancer_reach_out#From_Mentions"
	return
}
sHtml := CNGetHtml()
If !sHtml
	return
sEmailList := CNMentions2Emails(sHtml)
Teams_Emails2ChatDeepLink(sEmailList)
return

; ----------------------------------------------------------------------
MenuItemEmails2Emails:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_enhancer_reach_out#From_Emails"
	return
}
sHtml := CNGetHtml()
sEmailList := People_GetEmailList(sHtml)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
WinClip.SetText(sEmailList)
TrayTipAutoHide("Copy Emails", "Emails were copied to the clipboard!")
return

; ----------------------------------------------------------------------
MenuItemEmails2Mentions:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_enhancer_reach_out#From_Emails"
	return
}
sHtml := CNGetHtml()
sEmailList := People_GetEmailList(sHtml)
If (sEmailList = "")
    return ; empty
sHtmlMentions := Connections_Emails2Mentions(sEmailList)
WinClip.SetHtml(sHtmlMentions)
TrayTipAutoHide("People Connector","Mentions were copied to clipboard in RTF!")   
return

; ----------------------------------------------------------------------
MenuItemEmails2SendMentions:
If GetKeyState("Ctrl") {
	Run, "https://tdalon.blogspot.com/2020/11/connections-enhancer-send-mentions.html"
	return
}
Connections_SendMentions(Clipboard)
return

; ----------------------------------------------------------------------
MenuItemEmails2Chat:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_enhancer_reach_out#From_Emails"
	return
}
sHtml := CNGetHtml()
sEmailList := Html2Emails(sHtml)
Teams_Emails2ChatDeepLink(sEmailList)
return

; ----------------------------------------------------------------------
MenuItemProfileSearchToMentions:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_enhancer_reach_out#From_Profile_Search"
	return
}
sUrl := GetActiveBrowserUrl()
sHtml := CNGet(sUrl) ; TBC do not display sporadic error

sHtml := ProfileSearchHtml2Mentions(sHtml)
WinClip.SetHtml(sHtml)
TrayTipAutoHide("Copy Mentions", "Mentions were copied to the clipboard!")
return

; ----------------------------------------------------------------------
MenuItemProfileSearchToEmails:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_enhancer_reach_out#From_Profile_Search"
	return
}
sUrl := GetActiveBrowserUrl()


sEmailList := Connections_ProfileSearch2Emails(sUrl)
WinClip.SetHtml(sHtml)
TrayTipAutoHide("Copy Mentions", "Mentions were copied to the clipboard!")
return

; ----------------------------------------------------------------------
MenuItemCopyUserTable:
sUrl := GetActiveBrowserUrl()
sHtml := CNGet(sUrl) ; TBC

sHtml := GenerateUserTable(sHtml)
WinClip.SetHtml(sHtml)
TrayTipAutoHide("Copy User Table", "Html Table was copied to the clipboard!")
return

; ----------------------------------------------------------------------

MenuItemProfileSearch2Emails:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_enhancer_generate_nice_user_table" ; TODO
	return
}
sUrl := GetActiveBrowserUrl()
sEmailList := Connections_ProfileSearch2Emails(sUrl)
If (sEmailList ="")
	return
Clipboard := sEmailList
TrayTipAutoHide("Copy Emails", "List of Emails was copied to the clipboard!")
return

; ----------------------------------------------------------------------
MenuItemUpdateUserTable:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_enhancer_generate_nice_user_table"
	return
}
sUrl := Clipboard
ReConnectionsRootUrl := StrReplace(PowerTools_ConnectionsRootUrl,".","\.")
If !RegExMatch(sUrl,"^https?://" . ReConnectionsRootUrl .  "/profiles/html/") {
	MsgBox 0x1031,Connections Enhancer, Copy the search url to your clipboard, go back to the page to edit and press ok!
	IfMsgBox Cancel
		return
}
sUrl := Clipboard
If !RegExMatch(sUrl,"^https?://" . ReConnectionsRootUrl . "/profiles/html/") {
	;MsgBox 0x10,Connections Enhancer, Clipboard does not match url with 'connectionsroot/profiles/html/'!
	TrayTip, Update User Table, Clipboard does not match url with '<connectionsroot>/profiles/html/'!,,0x3
	return
}
sHtml := CNGetHtmlEditor()
If !sHtml {
	Send {Escape} 
	return
}
sProfileSearchHtml := CNGet(sUrl) ; TBC do not display error
sTableHtml := GenerateUserTable(sProfileSearchHtml)

sPat = Us)<table.*//connext\.conti\.de/profiles/photo\.do.*class="vcard".*</table>
Pos = 1 
While Pos := RegExMatch(sHtml,"Us)<table.*</table>",sMatch,Pos+StrLen(sMatch)){
	If RegExMatch(sMatch,sPat,sTableUserMatch)
		Break
}
If !sTableUserMatch {
	WinClip.SetHtml(sTableHtml)
	; Close source code editor
	Send {Escape}
	; MsgBox
	MsgBox 0x30,Connections Enhancer, No User Table was found!`nTable was copied to the clipboard. Paste in the rich-text editor at your desired location.
	return
}

sHtml := StrReplace(sHtml, sMatch, sTableHtml)

ClipPasteText(sHtml)

; Select and Press the OK button by Tab
Send {Tab} 
Send {Tab} 
Send {Enter}
return

; ----------------------------------------------------------------------
MenuItemAttendees2Emails:
	; Test https://connext.conti.de/communities/service/html/communityview?communityUuid=1f40ae3f-215e-48b2-86f5-7b43b7229d2c#fullpageWidgetId=W2a25ef83ef5c_4f39_b65b_578387091606&eventInstUuid=ed9fccde-6c68-4705-aa05-104f5dda28ee 
	; more than 100
	; MsgBox 0x10, Connections Enhancer: Error, You need to be in Edit mode!	
	sUrl := GetActiveBrowserUrl() 
	sEmailList := CNEvent2Emails(sUrl)
	If (sEmailList = "")
    	return ; empty
	Clipboard := sEmailList
	TrayTipAutoHide("Copy Emails", "List of emails from attendees was copied to the clipboard!")
return

; ----------------------------------------------------------------------
Event2Meeting:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_event_to_calendar"
	return
}
sUrl := GetActiveBrowserUrl() 
CNEvent2Meeting(sUrl)
return
; ----------------------------------------------------------------------

Event2Calendar:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_event_to_calendar"
	return
}
sUrl := GetActiveBrowserUrl() 
CNEvent2Meeting(sUrl,False) ; no Meeting
return
; ----------------------------------------------------------------------
Event2Email:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_event_update_handling"
	return
}
sUrl := GetActiveBrowserUrl() 
CNEvent2Email(sUrl) 
return

; ----------------------------------------------------------------------
MenuItemAttendees2Chat:
sUrl := GetActiveBrowserUrl() 
sEmailList := CNEvent2Emails(sUrl)
If (sEmailList = "")
	return ; empty
Teams_Emails2ChatDeepLink(sEmailList)	
return

; ----------------------------------------------------------------------
#include *i %A_ScriptDir%\ConNextMentionsMenuCallbacks.ahk ; only include if it exists


; -------------------------------------------------------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------
; FUNCTIONS
; -------------------------------------------------------------------------------------------------------------------

TransformHtml(sFuncName){
If Not IsFunc(sFuncName) {
	MsgBox Function %sFuncName% not defined!
	return
}
If Not IsWinConNextEdit()
{
	TrayTipAutoHide("Connections Enhancer", "This feature requires to be in Edit mode!")
	return
}
sHtml := CNGetHtmlEditor()

sHtml  := %sFuncName%(sHtml) 
If !sHtml {
	Send {Escape} 
	return
}

ClipPasteText(sHtml)

; Select and Press the OK button by Tab
Send {Tab} 
Send {Tab} 
Send {Enter}
}

; -------------------------------------------------------------------------------------------------------------------
ExpandMentionsWithProfilePic(){
; Expand selected mention or all mentions if no selection
If Not IsWinConNextEdit()
{
	TrayTipAutoHide("Connections Enhancer", "This feature requires to be in Edit mode!")
	return
}
global PT_wc
PT_wc.iCopy
sSelection := PT_wc.iGetHTML()
If (!sSelection) {
	sPat = U)<[^<]*class="vcard".*\?userid=([^"]*)".*</a>
	If (RegExMatch(sSelection, sPat, sSearch))
		sUid := sSearch1
}

sHtml := CNGetHtmlEditor()
sHtml  := Connections_ExpandMentionsWithProfilePicture(sHtml,sUid) 
If !sHtml {
	Send {Escape} 
	return
}

ClipPasteText(sHtml)

; Select and Press the OK button by Tab
Send {Tab} 
Send {Tab} 
Send {Enter}
}


; -------------------------------------------------------------------------------------------------------------------
PersonalizeMentions(){
; Personalize selected mention or all mentions if no selection
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_enhancer_personalize_mentions"
	return
}
sUrl := GetActiveBrowserURL()
If Not IsWinConNextEdit(sUrl)
{
	TrayTipAutoHide("Connections Enhancer", "This feature requires to be in Edit mode!")
	return
}
global PT_wc
PT_wc.iCopy
sSelection := PT_wc.iGetHTML()
If (!sSelection) {
	sPat = U)<[^<]*class="vcard".*\?userid=([^"]*)".*</a>
	If (RegExMatch(sSelection, sPat, sSearch))
		sUid := sSearch1
}

sHtml := CNGetHtmlEditor()
sHtml  := Connections_PersonalizeMentions(sHtml,sUid) 
If !sHtml {
	Send {Escape} 
	return
}

ClipPasteText(sHtml)

; Select and Press the OK button by Tab
Send {Tab} 
Send {Tab} 
Send {Enter}
}



; Conti color rgb(255, 165, 0)


; -------------------------------------------------------------------------------------------------------------------
;   BOXES
; -------------------------------------------------------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------
AddBox(keyword){
input := MultiLineInputBox("Enter text to be wrapped in a box block with icon:", "","Connections Enhancer: Box" )
If ErrorLevel ; Cancel
	return

input := StrReplace(input,"`n","<br>")
	
If (InStr(keyword,"brake"))
	imgsrc =http://connext.conti.de/files/form/anonymous/api/library/ffd279a0-2764-46bb-a8ec-b1aa3c713072/document/0d29e0f6-8f9b-4435-85bd-2ec696f60384/media/brake_warning_orange_32.png
Else If (InStr(keyword,"cancel"))
	imgsrc =http://connext.conti.de/files/form/anonymous/api/library/ffd279a0-2764-46bb-a8ec-b1aa3c713072/document/826829e1-2826-4f8c-ade3-8e04683ef787/media/cancel__red_32.png
Else If (InStr(keyword,"speed"))
	imgsrc =http://connext.conti.de/files/basic/anonymous/api/library/ffd279a0-2764-46bb-a8ec-b1aa3c713072/document/c6cf0424-67a7-4cdc-bc19-1f433ec3dd05/media/speed.png	
Else If (InStr(keyword,"idea"))
	imgsrc =http://connext.conti.de/files/basic/anonymous/api/library/ffd279a0-2764-46bb-a8ec-b1aa3c713072/document/eadf74fd-dea3-47e6-9d4a-8cfdfd8caaf0/media/idea_orange_32.png
Else If (InStr(keyword,"bug"))
	imgsrc = http://connext.conti.de/files/form/anonymous/api/library/ffd279a0-2764-46bb-a8ec-b1aa3c713072/document/75d06b4d-6e60-4718-a00e-6d549042f004/media/bug_orange_32.png

If (InStr(keyword,"table")) {
	sHtml = <div style="background: rgb(255, 255, 255); border-width: 0.1em 0.1em 0.1em 0.8em; border-style: solid; border-color: rgb(255, 165, 0); padding: 0.2em 0.6em; border-image: none; width: auto;" dir="ltr"><table width="95_percent"><tr ><td style="vertical-align: middle;table-layout: fixed; width: 40px;"><img src="%imgsrc%"/></td><td>%input%</td></tr></table></div>
	sHtml := StrReplace(sHtml,"_percent","%")
	}
Else
	sHtml = <div style="background: rgb(255, 255, 255); border-width: 0.1em 0.1em 0.1em 0.8em; border-style: solid; border-color: rgb(255, 165, 0); padding: 0.2em 0.6em; border-image: none; width: auto; " dir="ltr"><img src="%imgsrc%"/> %input%</div>
; #TODO overflow: auto; or hidden
ClipPasteHtml(sHtml)

return
}
; -------------------------------------------------------------------------------------------------------------------
;   KUDOS
; -------------------------------------------------------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------
AddKudos(keyword){

sHtml := ConNext_Kudos2Html(StrReplace(keyword," (Table)",""))
If (InStr(keyword,"table")) {

	input := MultiLineInputBox("Enter text to be wrapped in a box block with icon:", "","Connections Enhancer: Kudos" )
	If ErrorLevel
		return

	input := StrReplace(input,"`n","<br>")

	sHtml = <div style="background: rgb(255, 255, 255); border-width: 0.1em 0.1em 0.1em 0.8em; border-style: solid; border-color: rgb(255, 165, 0); padding: 0.2em 0.6em; border-image: none; width: auto;" dir="ltr"><table width="95_percent" cellpadding="10"><tr ><td style="vertical-align: middle;table-layout: fixed; width: 60px;">%sHtml%</td><td><span style="font-family: Arial; font-size: 14px;">%input%</span></td></tr></table></div>
	sHtml := StrReplace(sHtml,"_percent","%")
	; Change image size
	sHtml := StrReplace(sHtml,"<img ","<img width='50'")
	}
Else
	sHtml = <p style="text-align: center;">%sHtml%<p>

ClipPasteHtml(sHtml)
	
return
}

; -------------------------------------------------------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------
AddMsg(keyword){

input := MultiLineInputBox("Enter text to be wrapped in a Message block.`nN.B.: Box will only be visible after Save and Close.", "","Connections Enhancer: Message Box")
If ErrorLevel ; Cancel
	return

input := StrReplace(input,"`n","<br>")


; keyword can be success, info, warning or error
If (keyword="info")
	lotusIcon = lotusIconMsgInfo
Else If (keyword="success")
	lotusIcon = lotusIconMsgSuccess
Else If (keyword="warning")
	lotusIcon = lotusIconMsgWarning
Else If (keyword ="error")
	lotusIcon = lotusIconMsgError

sHtml = <div class="lotusMessage lotusWarning" dir="ltr" role="alert"><img class="lotusIcon %lotusIcon%" src="http://connext.conti.de/connections/resources/web/com.ibm.lconn.core.styles.oneui3/images/blank.gif" /><div class="lotusMessageBody"> %input%</div></div>

ClipPasteHtml(sHtml)	
return
}

; -------------------------------------------------------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------
; Returns true if current window can be an opened connext editor
IsWinConNextEdit(sUrl := "") {
If !sUrl ; empty
	sUrl := GetActiveBrowserUrl()
ReConnectionsRootUrl := StrReplace(PowerTools_ConnectionsRootUrl,".","\.")
If InStr(sUrl,PowerTools_ConnectionsRootUrl . "/wikis/")
	return RegExMatch(sURL, ReConnectionsRootUrl . "/wikis/.*/edit") || RegExMatch(sURL, ReConnectionsRootUrl . "/wikis/.*/create")
Else If InStr(sUrl, PowerTools_ConnectionsRootUrl . "/blogs/")
	return InStr(sUrl, PowerTools_ConnectionsRootUrl . "/blogs/roller-ui/authoring/weblog.do") ; method edit or create
Else
	return InStr(sUrl, PowerTools_ConnectionsRootUrl . "forums/html/topic?") or InStr(sUrl, PowerTools_ConnectionsRootUrl . "/forums/html/threadTopic?") or  RegExMatch(sURL, ReConnectionsRootUrl . "/forums/html/forum\?id=.*showForm=true")
}

; -------------------------------------------------------------------------------------------------------------------
GenerateWikiTocSubPages(sUrl,nDepth,sTocStyle) {
; Generate Table of Contents based on headings
; Syntax:
;   sTocHtml := GenerateWikiTocSubPages(sUrl,nDepth,sTocStyle)

; Test: "https://connext.conti.de/wikis/home?lang=en-us#!/wiki/W1d0f7e400e73_4328_9930_7e375562b15c/page/PMT%20Guidelines"

; also support /edit links

If RegExMatch(sUrl,"/wikis/.*/wiki/(.*?)/page/([^/]*)",sMatch) {
	sWikiLabel := sMatch1
	sPageLabel := sMatch2
	sWikiPageId := WikiGetPageId(sWikiLabel,sPageLabel)
} Else If RegExMatch(sUrl,"/wikis/.*/wiki/(.*?)/draft/([^/]*)",sMatch) {
	MSgBox 0x10, Connections Enhancer: Error, GenTocWiki Function does not work on a draft page!
	return
} Else {
	MSgBox 0x10, Connections Enhancer: Error, GenTocWiki Function only works for a Wiki page in Edit mode!
	return
}
; Get Navigation Json Feed
sFeedUrl := RegExReplace(sUrl, "/wikis/.*/wiki/(.*?)/.*","/wikis/basic/anonymous/api/wiki/$1/navigation/feed")

sNav := CNGet(sFeedUrl)
If (sNav ~= "Error.*") 
	return

If !sTocStyle 
	sTocStyle:= PowerTools_RegRead("CNTocStyle")
	;RegRead, sTocStyle, HKEY_CURRENT_USER\Software\PowerTools, CNTocStyle
If !sTocStyle 
	sTocStyle := "?"
If (sTocStyle = "?") {
	OnMessage(0x44, "OnTocStyleMsgBox")
	MsgBox 0x23, Toc Style, Select your Table of Content style:
	OnMessage(0x44, "")
	sTocStyle = bullet-white ; Default
	IfMsgBox, No
		sTocStyle = none-yellow	
	IfMsgBox, Cancel
        sTocStyle = num-white
}


If !nDepth {
	; Input nDepth
	InputBox, nDepth , Wiki TOC SubPages Depth, Enter the depth of your SubPages TOC.`nLeave empty for all., , 300, 150
}
If ErrorLevel or (nDepth = "")
	nDepth = Inf


If (sTocStyle = "bullet-white") {
	sLiStyle =
} 
Else If (sTocStyle = "none-yellow")  { 
	sLiStyle = style="list-style-type:none;"
	sTocDiv = style="border-radius:6px;margin:8px;padding:4px;display:block;width:auto;background-color:#ffc"
}
Else If (sTocStyle = "num-white")  {
}
Else {
	MsgBox "Error: wrong Toc Style!"
	return
}
sTocDiv = <div id="toc-wiki_%nDepth%_%sTocStyle%" %sTocDiv%>

If RegExMatch(sTocStyle,"^num-") {
	sLiStyle1 =<ol %sLiStyle%>
	sLiStyle2 =</ol>
}
Else	{
	sLiStyle1 = <ul %sLiStyle%>
	sLiStyle2 =</ul>
}

sTocHtml := GetTocHtmlChild(sWikiPageId,sWikiLabel,sNav,nDepth,sLiStyle1,sLiStyle2)

sTocHtml = %sTocDiv%<b>Table of Contents (SubPages)</b>  %sTocHtml%</div>
; Clipboard := sTocHtml
return sTocHtml
}	; eofun
; -------------------------------------------------------------------------------------------------------------------------------------

GetTocHtmlChild(sParentId,sWikiLabel,sNav,nDepth,sLiStyle1,sLiStyle2) {
sPat = {"parent":"%sParentId%","id":"(.*?)".*?"title":"(.*?)".*?}
Pos=1
While Pos :=    RegExMatch(sNav, sPat, sMatch,Pos+StrLen(sMatch))  {   
    sWikiUrl = https://%PowerTools_ConnectionsRootUrl%/wikis/home/wiki/%sWikiLabel%/page/%sMatch1%
	sHtml = %sHtml%<li><a href="%sWikiUrl%">%sMatch2%</a>
	If (nDepth="Inf") 
		sHtml := sHtml . GetTocHtmlChild(sMatch1,sWikiLabel,sNav,nDepth,sLiStyle1,sLiStyle2) 
	Else If (nDepth -1 >0)
		sHtml := sHtml . GetTocHtmlChild(sMatch1,sWikiLabel,sNav,nDepth-1,sLiStyle1,sLiStyle2) 
	
	sHtml = %sHtml%</li>
} ; End While
If !sHtml ; no children
	return
sHtml := sLiStyle1 . sHtml . sLiStyle2
return sHtml
}
; -------------------------------------------------------------------------------------------------------------------------------------

AutoFixHeadingIds(sHtml){
sOutHtml := sHtml
sPat = sU)<h(\d)(.*)>(.*)</h\d> ; ungreedy dotall regexp
Pos=1
sPatId = id="ii\d*?"
While Pos :=    RegExMatch(sHtml, sPat, sMatch,Pos+StrLen(sMatch))  {
    If (RegExMatch(sMatch2, sPatId,sMatchId))  { ; id default value ii* 
        sId := UnHTM(sMatch3) ; remove Html formatting
		sId := StrReplace(sId,"&nbsp;","_")
		sId := StrReplace(sId," ","_") 
		sQuote = "
		sId := StrReplace(sId,sQuote,"")
		sId := StrReplace(sId," ","_") 
		sId = id="%sId%"
        sHeader := StrReplace(sMatch,sMatchId,sId)
		sOutHtml := StrReplace(sOutHtml, sMatch , sHeader,,1) ; Replace only the first one
    }   
}
return sOutHtml
} ; eof

; -------------------------------------------------------------------------------------------------------------------------------------
GenerateToc(sHtml,nDepth,sTocStyle,sInsLoc:="top"){
; Generate Table of Contents based on headings
; Syntax:
;   sHtml := GenerateToc(sHtml,nDepth,sTocStyle)
; Heading structure shall not jump any level!

If !nDepth {
	; Input nDepth
	InputBox, nDepth , TOC Depth, Enter till which Heading Level you want to generate your TOC.`nLeave empty for all., , 300, 160
}
If ErrorLevel or (nDepth = "")
	nDepth = Inf
	

If !sTocStyle 
	sTocStyle:= PowerTools_RegRead("CNTocStyle")
If !sTocStyle 
	sTocStyle := "?"
If (sTocStyle = "?") {
	OnMessage(0x44, "OnTocStyleMsgBox")
	MsgBox 0x23, Toc Style, Select your Table of Content style:
	OnMessage(0x44, "")
	sTocStyle = bullet-white ; Default
	IfMsgBox, No
		sTocStyle = none-yellow	
	IfMsgBox, Cancel
		sTocStyle = num-white
}

;sTocDiv = <div id="toc" data-depth="%nDepth%" data-toc-style="%sTocStyle%"
sTocDiv = <div id="toc_%nDepth%_%sTocStyle%"
If (sTocStyle = "bullet-white") or (sTocStyle = "num-white") {
	sLiStyle =
} 
Else If (sTocStyle = "none-yellow")  { 
	sLiStyle = style="list-style-type:none;"
	sTocDiv = %sTocDiv% style="border-radius:6px;margin:8px;padding:4px;display:block;width:auto;background-color:#ffc"
}
Else {
	MsgBox "Error: wrong Toc Style!"
	return
}

sTocDiv = %sTocDiv%>

sPat = sU)<h(\d)(.*)>(.*)</h\d> ; ungreedy dotall regexp

sOutHtml := sHtml
Pos=1
nPrev = 0 
While Pos:= RegExMatch(sHtml, sPat, sMatch,Pos+StrLen(sMatch)) {       
    nCur := +sMatch1 ; convert string to number
	If (nDepth <> "Inf") and (nCur > nDepth)
		Continue
	
	sPatId = id="(.*?)"

	; Clean-up Heading
	sHeadingDisp := sMatch3
	; Clean-up <br />
	sHeadingDisp := StrReplace(sHeadingDisp,"<br />","")
	sHeadingDisp := RegExReplace(sHeadingDisp,"^\s*","") ;  remove starting blanks
	sHeadingDisp := RegExReplace(sHeadingDisp,"\s*$","") ;  remove trailing blanks

    If (RegExMatch(sMatch2, sPatId,sId))  { ; id already defined 
		sHeader = <h%sMatch1%%sMatch2%>%sHeadingDisp%</h%sMatch1%>
		sId := sId1
	}
    Else {
        sId := UnHTM(sMatch3) ; remove Html formatting
		sId := StrReplace(sId,"&nbsp;","_")
		sId := StrReplace(sId," ","_") 
		sId := StrReplace(sId,"'","")
        sHeader = <h%sMatch1% id="%sId%" %sMatch2%>%sHeadingDisp%</h%sMatch1%>
    }

    sOutHtml := StrReplace(sOutHtml, sMatch , sHeader)
  
    If (nCur> nPrev) {
		; Opening sublist
		sTocHtml = %sTocHtml%<ul %sLiStyle%>
	} Else If (nCur< nPrev) { ; closing sublist
		nCount := nPrev - nCur
		Loop  %nCount% {
			sTocHtml = %sTocHtml%</ul>
		}
	} Else {
		sTocHtml = %sTocHtml%</li>
	}
		
	sTocHtml = %sTocHtml%<li><a href="#%sId%">%sHeadingDisp%</a>
        
	nPrev := nCur 
} ; End While

; Closing Lists
; Closing sublists 
Loop %nCur% {
	sTocHtml = %sTocHtml%</li>
}
sTocHtml = %sTocHtml%</ul>

; convert unordered list to ordered list
If RegExMatch(sTocStyle,"^num-") {
	sTocHtml := StrReplace(sTocHtml,"<ul ","<ol ")
	sTocHtml := StrReplace(sTocHtml,"</ul> ","</ol>")
}

sTocHtml = %sTocDiv%<b>Table of Contents</b>  %sTocHtml%</div>

sPat = sU)<div id="toc_.*".*>.*</div>
If RegExMatch(sOutHtml,sPat,sMatch) { ; Toc Already Exist
	sOutHtml := RegExReplace(sOutHtml, sPat, sTocHtml)
} Else { ; create new
	If (sInsLoc = "cursor") {
		sLoc = <p dir="ltr">$LOCMARKER$</p>
		sOutHtml := StrReplace(sOutHtml,sLoc,sTocHtml)
		sOutHtml := StrReplace(sOutHtml,"$CURSORLOC$",sTocHtml)
	} else if (sInsLoc = "top") { 
		sOutHtml := sTocHtml . "`n" . sOutHtml
	} else if (sInsLoc = "bottom") { 
		sOutHtml := sOutHtml . "`n" . sTocHtml
	} else 
		sOutHtml = %sTocHtml%%sOutHtml%
}
return sOutHtml
}	; eofun

; -------------------------------------------------------------------------------------------------------------------------------------
InsertWikiToc(sHtml,sInsLoc:="top"){
; Generate Table of Contents based on headings
; Syntax:
;   sHtml := GenerateToc(sHtml,nDepth,sTocStyle)
; Heading structure shall not jump any level!
sTocHtml = <div style="border-radius:6px;margin:8px;padding:4px;display:block;width:auto;background-color:#ffc" name="intInfo" contenteditable="false" dir="ltr">Table of Contents:</div>
	sHtml := AutoFixHeadingIds(sHtml) ; Clean HeadingIds
	If (sInsLoc = "cursor") {
		sLoc = <p dir="ltr">$LOCMARKER$</p>
		sHtml := StrReplace(sHtml,sLoc,sTocHtml)
		sHtml := StrReplace(sHtml,"$CURSORLOC$",sTocHtml)
	} else if (sInsLoc = "top") { 
		sHtml := sTocHtml . "`n" . sHtml
	} else if (sInsLoc = "bottom") { 
		sHtml := sHtml . "`n" . sTocHtml
	}
return sHtml
}	; eofun


OnTocStyleMsgBox() {
    DetectHiddenWindows, On
    Process, Exist
    If (WinExist("ahk_class #32770 ahk_pid " . ErrorLevel)) {
        ControlSetText Button1, bullet-white
		ControlSetText Button2, none-yellow
		ControlSetText Button3, num-white
    }
}

OnWikiNewPageMsgBox() {
    DetectHiddenWindows, On
    Process, Exist
    If (WinExist("ahk_class #32770 ahk_pid " . ErrorLevel)) {
        ControlSetText Button1, child
		ControlSetText Button2, peer
		ControlSetText Button3, root
    }
}

OnInsLocMsgBox() {
    DetectHiddenWindows, On
    Process, Exist
    If (WinExist("ahk_class #32770 ahk_pid " . ErrorLevel)) {
        ControlSetText Button1, top
		ControlSetText Button2, cursor
		ControlSetText Button3, bottom
    }
}

; ----------------------------------------------------------------------
CNGetHtml(){
; Syntax: sHtml := CNGetHtml()
; Calls: ConNextGetHtmlEditor, GetActiveBrowserUrl,ConNextGetSource
If IsWinConNextEdit() 
{
	sHtml := CNGetHtmlEditor()
	Send {Esc}  ; Close Html Source Window
} Else {
	; MsgBox 0x10, Connections Enhancer: Error, You need to be in Edit mode!	
	sHtml := GetSelection("html")
	If !sHtml {
		sUrl := GetActiveBrowserUrl() 
		sHtml := CNGet(sUrl)
	}
}
return sHtml
}	

; ----------------------------------------------------------------------
CNGetHtmlEditor2(){
	; Get Html from ConNext Editor using Hotkey
	; only works with textbox.io
	Send ^+u ; Open code view via Ctrl+Shift+U Hotkey
	; Copy all source code to clipboard
	Send ^a
	sleep, 300
	global PT_wc
	PT_wc.iCopy()
	sHtml := PT_wc.iGetText()
	return sHtml
}

CNGetHtmlEditor(){
	; Get Html from ConNext Editor using Hotkey
	; only works with textbox.io
	Send ^+u ; Open code view via Ctrl+Shift+U Hotkey
	; Copy all source code to clipboard
	ClipSaved := ClipboardAll
	Clipboard = 
	Send ^a
	sleep, 300
	Send ^c
	ClipWait, 0.5
	sHtml := Clipboard
	; Restore Clipboard
	Clipboard := ClipSaved
	return sHtml
} ; TODO duplicate in Lib

BrowserGetPageSource(){
	; Window will flash. Alternative use HttpGet but requires password setting for ConNext
	Send ^u ; Open code view via Ctrl+U Hotkey (Chrome)
	; Copy all source code to clipboard
	ClipSaved := ClipboardAll
	Clipboard = 
	sleep, 300
	Send ^a
	sleep, 300
	Send ^c
	ClipWait, 0.5
	sHtml := Clipboard
	Send ^w
	; Restore Clipboard
	Clipboard := ClipSaved
	return sHtml
}

; ----------------------------------------------------------------------
CNViewSource(){
; only works with textbox.io
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_view_html" 
	return
}

If !IsWinConNextEdit() 
{
	TrayTipAutoHide("View Source", "View Source is only possible in Edit mode!")
	return
}
SendInput $CURSORLOC$

Send ^+u ; Open code view via Ctrl+Shift+U Hotkey
; Copy all source code to clipboard

sleep, 500
Send ^f
sleep, 500
SendInput $CURSORLOC$
Send {Escape}{Backspace}

}


; -------------------------------------------------------------------------------------------------------------------
CleanCode(sHtml){
; Clean table style
sPat = rel="noopener noreferrer.*?"
sHtml := RegExReplace(sHtml,sPat,"")
;sHtml := StrReplace(sHtml,"&nbsp;"," ")
sHtml := StrReplace(sHtml,"&amp;","&")

sHtml := StrReplace(sHtml,"white-space:pre","") ; copy from Visual Studio VS Code

; Remove black font formatting
sPat = Us)<span style="color:rgb\( 0 , 0 , 0 \)">(.*)</span>
sHtml := RegExReplace(sHtml,sPat,"$1")

; Remove id="wikiLink.*"
sPat = U)id="wikiLink.*"
sHtml := RegExReplace(sHtml,sPat,"")

return sHtml
}
; -------------------------------------------------------------------------------------------------------------------
; ----------------------------------------------------------------------
CleanLinks(sHtml){
; Clean ConNext Url from lang tags
sHtml := StrReplace(sHtml,"http://" . PowerTools_ConnectionsRootUrl . "/","https://" . PowerTools_ConnectionsRootUrl . "/")
sHtml  := RegExReplace(sHtml , "[&|\?]lang=[\w-_]+(#!)?" , "")	
; Remove lastMod part from links
sHtml := RegExReplace(sHtml, "&lastMod=\d+" , "")

return sHtml
}
; ----------------------------------------------------------------------
CleanTable(sHtml){
; Clean table style

;sPat = <td style="[^"]*?
;sRep = <td style="text-align:center; 
;sHtml := RegExReplace(sHtml,sPat,sRep)

;<td style="width:288px;border-color:#696969">

sHtml := RegExReplace(sHtml,"<td style=""width:[^;]*","<td style=""")
;sHtml := RegExReplace(sHtml,"<td style=""height:.*?;""","<td ")
sPat = <tr style="width:.*?;"
sHtml := RegExReplace(sHtml,sPat,"<tr ")
sPat = <tr style="height:.*?;"
sHtml := RegExReplace(sHtml,sPat,"<tr ")
;sHtml := RegExReplace(sHtml,"height:.*?;","")
;sHtml := RegExReplace(sHtml,"style=""\s*?""","") ; remove empty style

;Full Width: http://connext.conti.de/blogs/tdalon/entry/nice_html_tables
sPat =  <table (.*?) style="[^"]*
;table border="1" style="border-collapse: collapse; width: 100%;"
sRep = <table $1 style="table-layout: fixed; width: 95_percent;
sRep := StrReplace(sRep,"_percent","%")
sHtml := RegExReplace(sHtml,sPat,sRep)

return sHtml
}
; ----------------------------------------------------------------------


; ----------------------------------------------------------------------
RefreshToc(){
sUrl := GetActiveBrowserUrl()
;If !RegExMatch(sUrl,".*/connext.conti.de/wikis/.*/edit") {
;	MSgBox 0x10, Connections Enhancer: Error, GenTocWiki Function only works for a Wiki page in Edit mode!
;	return
;}

sHtml := CNGetHtmlEditor()
If !sHtml {
	Send {Escape} 
	return
}

If InStr(sUrl, PowerTools_ConnectionsRootUrl . "/wikis/") { ; a wiki page
	sHtml := AutoFixHeadingIds(sHtml) ; Clean Heading Ids
	sPat = sU)<div id="toc-wiki_(.*)_(.*)".*>.*</div>
	If (RegExMatch(sHtml,sPat,sMatch)) { ; refresh TOC SubPages if it exists
		sTocHtml := GenerateWikiTocSubPages(sUrl,sMatch1,sMatch2)
		sHtml := RegExReplace(sHtml, sPat, sTocHtml)
	} Else { ; create new
	; 	sMatch :=
	; 	sTocHtml := GenerateWikiTocSubPages(sUrl,sMatch,sMatch)
	; 	sHtml = %sTocHtml%%sHtml% ; Position Toc on the top
		sPat = <div [^>]* name="intInfo" contenteditable="false" dir="ltr">
		If !(RegExMatch(sHtml,sPat))
			sHTml := InsertWikiToc(sHtml) ; Insert Toc at the top of the page
	}
	
} Else { ; not a wiki e.g. blog, forum
	sPat = sU)<div id="toc_(.*)_(.*)".*>.*</div>
	If RegExMatch(sHtml,sPat,sMatch)
		sHtml := GenerateToc(sHtml, sMatch1, sMatch2)
	Else ; create new
		sHtml := GenerateToc(sHtml,sMatch,sMatch)
}

ClipPasteText(sHtml)

; Select and Press the OK button by Tab
Send {Tab} 
Send {Tab} 
Send {Enter}
}



; ----------------------------------------------------------------------
DeleteToc(){
sUrl := GetActiveBrowserUrl()
;If !RegExMatch(sUrl,".*/connext.conti.de/wikis/.*/edit") {
;	MSgBox 0x10, Connections Enhancer: Error, GenTocWiki Function only works for a Wiki page in Edit mode!
;	return
;}

sHtml := CNGetHtmlEditor()
If !sHtml {
	Send {Escape} 
	return
}

If InStr(sUrl, PowerTools_ConnectionsRootUrl . "/wikis/") { ; a wiki page
	sPat = sU)<div id="toc-wiki_(.*)_(.*)".*>.*</div>
	If (RegExMatch(sHtml,sPat,sMatch))  
		sHtml := RegExReplace(sHtml, sPat, "")
	
	sPat = sU)<div style="border-radius:6px;margin:8px;padding:4px;display:block;width:auto;background-color:#ffc" name="intInfo" contenteditable="false" dir="ltr">Table of Contents:.*</div>
	If (RegExMatch(sHtml,sPat,sMatch)) 
		sHtml := RegExReplace(sHtml, sPat, "")	
	
	
} Else { ; not a wiki e.g. blog, forum
	sPat = sU)<div id="toc_(.*)_(.*)".*>.*</div>
	If RegExMatch(sHtml,sPat,sMatch)
		sHtml := RegExReplace(sHtml, sPat, "")
}

ClipPasteText(sHtml)

; Select and Press the OK button by Tab
Send {Tab} 
Send {Tab} 
Send {Enter}
}
; ----------------------------------------------------------------------
CNWikiGoTo(sUrl,target){

If !RegExMatch(sUrl,"/wikis/.*/wiki/(.*?)/page/([^/]*)",sMatch) {
	MsgBox 0x10, Connections Enhancer: Error, You shall have a ConNext Wiki page opened!
	return
}
sWikiLabel := sMatch1
sPageLabel := sMatch2
sApiUrl = https://%PowerTools_ConnectionsRootUrl%/wikis/basic/anonymous/api/wiki/%sWikiLabel%/navigation/%sPageLabel%/entry?includeSiblings=true&format=json
sResponse := CNGet(sApiUrl)

sPat = {"parent":(.*?),"previousSibling":(.*?),"id":"(?:.*?)","label":"(?:.*?)","nextSibling":(.*?),"mediaId":"(?:.*?)","title":"(?:.*?)"}
If !RegExMatch(sResponse,sPat,sMatch){
	MsgBox 0x10, Connections Enhancer: Error, Could not parse Response!
	return
}

If (target = "parent")
	sNextId := sMatch1
Else If (target = "prev")
	sNextId := sMatch2
Else If (target = "next")
	sNextId := sMatch3

If (sNextId = "null"){
	MsgBox 0x10, Connections Enhancer: Error, No page to navigate to!
	return
}
Else { ; remove surrounding quotes
	sNextId := SubStr(sNextId,2,-1)
}

sNextUrl = https://%PowerTools_ConnectionsRootUrl%/wikis/home/wiki/%sWikiLabel%/page/%sNextId%
Send ^w ; close current window
ClipboardSaved := ClipboardAll

	Clipboard := sNextUrl
	Send ^t ; Open new tab
	Send ^v
	Send {Enter}
; restore Clipboard
Clipboard := ClipboardSaved
} ; eofun

; ----------------------------------------------------------------------
CNCreateNew(){
ReConnectionsRootUrl := StrReplace(PowerTools_ConnectionsRootUrl,".","\.")
sUrl := GetActiveBrowserUrl()
If RegExMatch(sUrl,"/wikis/.*/wiki/(.*?)/page/([^/\?]*)",sMatch) { ; exclude also section
	sWikiLabel := sMatch1
	sPageLabel := sMatch2
	
	OnMessage(0x44, "OnWikiNewPageMsgBox")
	MsgBox 0x23, Wiki New Page, Select where the new page shall be located:
	OnMessage(0x44, "")
	sNewPage = child ; Default
	IfMsgBox, No
		sNewPage = peer	
	IfMsgBox, Cancel
        sNewPage = root
	
	; Root
	sUrl = https://%PowerTools_ConnectionsRootUrl%/wikis/home/wiki/%sWikiLabel%/pages/create
	If (sNewPage = "root") {
		Run, %sUrl%
		return
	}
	
	sApiUrl = https://%PowerTools_ConnectionsRootUrl%/wikis/basic/anonymous/api/wiki/%sWikiLabel%/navigation/%sPageLabel%/entry?includeSiblings=true&format=json
	sResponse := CNGet(sApiUrl) 
	sPat = {"parent":(.*?),"previousSibling":(?:.*?),"id":"(.*?)","label":"(?:.*?)","nextSibling":(?:.*?),"mediaId":"(?:.*?)","title":"(?:.*?)"}
	If !RegExMatch(sResponse,sPat,sMatch){
		MsgBox 0x10, Connections Enhancer: Error, Could not parse Response!
		return
	}
	
	If (sNewPage = "peer")  ; sibling	
		sWikiPageId := sMatch1

	Else If (sNewPage = "child") {
		sWikiPageId := sMatch2
	}

	sUrl = %sUrl%?parentId=%sWikiPageId%
	Run, %sUrl%

} Else If InStr(sUrl,"://" . PowerTools_ConnectionsRootUrl . "/forums/") {
	sForumId := CNGetForumId(sUrl)
	If !sForumId 
		return
	sUrl := "https://" . PowerTools_ConnectionsRootUrl . "/forums/html/forum?id=" . sForumId . "&showForm=true"
	Run, %sUrl%
} Else If RegExMatch(sUrl,"//" . ReConnectionsRootUrl . "/blogs/([^/\?&]*)",sBlogId) {
	sUrl := "https://" . PowerTools_ConnectionsRootUrl . "/blogs/roller-ui/authoring/weblog.do?method=create&weblog=" . sBlogId1
	Run, %sUrl%	
} Else {
	MsgBox 0x10, Connections Enhancer: Error, You shall have a ConNext Forum, Blog or Wiki page opened!
	return
}

} ; eofun

; ----------------------------------------------------------------------
WikiGetPageId(sWikiLabel,sPageLabel){
; Get Wiki Page Id
	sApiUrl = https://%PowerTools_ConnectionsRootUrl%/wikis/basic/anonymous/api/wiki/%sWikiLabel%/navigation/%sPageLabel%/entry?includeSiblings=true&format=json
	sResponse := CNGet(sApiUrl)

	sPat = {"parent":(?:.*?),"previousSibling":(?:.*?),"id":"(.*?)","label":"(?:.*?)","nextSibling":(?:.*?),"mediaId":"(?:.*?)","title":"(?:.*?)"}
	If !RegExMatch(sResponse,sPat,sMatch){
		MsgBox 0x10, Connections Enhancer: Error, Could not parse Response!
		return
	}
	return sMatch1
} ; eofun

; ----------------------------------------------------------------------
WikiGetPageIdOld(sWikiLabel,sPageLabel){
	; Will flash Get command in browser if without VPN
	sApiUrl = https://%PowerTools_ConnectionsRootUrl%/wikis/basic/anonymous/api/wiki/%sWikiLabel%/navigation/%sPageLabel%/entry

	IsVPN := Not (A_IPAddress2 = 0.0.0.0) ; HttpGet UnAuthorized via VPN
	If IsVPN  { ; VPN  
		;MSgBox 0x10, Connections Enhancer: Error, GenTocWiki Function does not work with VPN!
		; return
		sResponse := BrowserGetPage(sApiUrl)
	}
	Else {
		sResponse := CNGet(sApiUrl)
		If (sResponse ~= "Error.*") 
            return
	}

	sPat = <td:uuid>(.*?)</td:uuid> 
	If !RegExMatch(sResponse,sPat,sMatch) {
		MSgBox 0x10, Connections Enhancer: Error,No Page Id found!
		return
	}
	return sMatch1
}
; ----------------------------------------------------------------------
CopyEmails(sHtml){
sPat = U)Office email:\r\n<strong>\r\n<a href="mailto:(.*)" 
Pos = 1 
While Pos := RegExMatch(sHtml,sPat,sEmail,Pos+StrLen(sEmail)){
    sEmailList = %sEmailList%;%sEmail1%
}
return SubStr(sEmailList,2) ; remove first ;
} ; eof

; ----------------------------------------------------------------------
CenterIframe(sHTMLCode){
; Format Iframe with align-center 
;   sHTMLCode := CenterIframe(sHTMLCode)
; Called by: Menu CenterIframe
sPat = <p([^->]*)><iframe ([^>]*)> ; exclude - in text-align
sRep = <p style="text-align: center;" $1><iframe $2>

sHTMLCode := RegExReplace(sHTMLCode,sPat,sRep)

return sHTMLCode
}

; ----------------------------------------------------------------------
ProfileSearchHtml2Mentions(sHtml){
; Syntax:
; sMentionsHtml := ProfileSearchHtml2Mentions(sHtml)
sPat = Us)<div role="article" aria-label="(.*)">.*<span class="x-lconn-userid" style="display: none;">(.*)</span>.*<a class="fn url person bidiAware lotusBold" href="/profiles/.*">.*</a>.*Office email:.*<strong>\r\n<a href="mailto:(.*)" 

Pos = 1 
While Pos := RegExMatch(sHtml,sPat,sMatch,Pos+StrLen(sMatch)){
	sMention := "<span contenteditable='false' class='vcard'><a href='http://" . PowerTools_ConnectionsRootUrl . "/profiles/html/profileView.do?email=" . sMatch3 . "' class='fn url'>@" . sMatch1 . "</a><span class='x-lconn-userid' style='display: none;'>" . sMatch2 . "</span></span>"
	sMentionsHtml = %sMentionsHtml% %sMention%
}
return sMentionsHtml ; remove first ;
} ; eof

; ----------------------------------------------------------------------


; ----------------------------------------------------------------------
GenerateUserTable(sHtml){
; Generate Table of users with pictures and at-mentions from profile search Html source code
; Syntax:
;   sTableHtml := GenerateUserTable(sHtml)
sPat = Us)<div role="article" aria-label="(.*)">.*<span class="x-lconn-userid" style="display: none;">(.*)</span>.*<a class="fn url person bidiAware lotusBold" href="/profiles/.*">.*</a>.*Office email:.*<strong>\r\n<a href="mailto:(.*)" 

nbCol = 5
j_col = 0
Pos = 1 
sTableHtml := "<table align='center' style='width:95`%;'>"


While Pos := RegExMatch(sHtml,sPat,sMatch,Pos+StrLen(sMatch)){	
	sEmail := sMatch3
	sUid := sMatch2
	sName := sMatch1
	sName := RegExReplace(sName," \(.*\)","") ; Remove (uid) in firstname

	;sMention := "<span contenteditable='false' class='vcard'><a href='http://connext.conti.de/profiles/html/profileView.do?email=" . sEmail . "' class='fn url'>@" . sName . "</a><span class='x-lconn-userid' style='display: none;'>" . sUid . "</span></span>"
  
	sMention := "<a class='vcard' data-userid='" . sUid . "' href='https://" . PowerTools_ConnectionsRootUrl . "/profiles/html/profileView.do?userid=" . sUid . "'>@" . sName . "</a>"

    j_col := j_col + 1
	If (j_col > nbCol) {
		j_col = 1
		sTableHtml := sTableHtml . "</tr><tr>"
	}            
	sPicSrc := "https://" . PowerTools_ConnectionsRootUrl . "/profiles/photo.do?email=" . sEmail
	sProfileUrl := "https://" . PowerTools_ConnectionsRootUrl . "/profiles/html/profileView.do?email=" . sEmail
	sPic := "<a href='" . sProfileUrl . "'><img width='100px' src='" . sPicSrc . "'/></a>"
	sTableHtml := sTableHtml . "<td style='text-align: center;'>" . sPic . "<br>" . sMention . "</td>"

} ; End While

sTdFill := "<td style='text-align: center;'><a href='https://connext.conti.de/wikis/home/wiki/W354104eee9d6_4a63_9c48_32eb87112262/page/CoachNet Specialist'><img width='100px' src='https://connext.conti.de/files/form/anonymous/api/library/df3541d3-1ef2-4799-9212-d269c9d79d35/document/672f484e-7bce-4074-8f2a-c64463ae75e6/media/Wanted_CoachNetter_Example.jpg'/></a></td>"

;Fill missing empty cell <td>
Count := nbCol-j_col
Loop , %Count%
    sTableHtml := sTableHtml . sTdFill
    
sTableHtml := sTableHtml . "</tr></table>"
return sTableHtml ; remove first ;
} ; eof

; ----------------------------------------------------------------------
Likers2Mentions(sUrl){
; Only for Blog entries
; Syntax:
;    Likers2Mentions(sUrl)
; Calls DownloadToString
sSource := DownloadToString(sUrl)
; Parse response for string between <title> </title> for atom <title type="html">
sPat = <input type="hidden" name="entryId" value="([^/]*)"/>
If RegExMatch(sSource, sPat, sEntryId)
	sEntryId := sEntryId1

; Get BlogId
Array := StrSplit(sUrl,"/")
sBlogId := Array[5]
sApiUrl = https://%PowerTools_ConnectionsRootUrl%/blogs/%sBlogId%/api/recommend/entries/%sEntryId%

sSource := DownloadToString(sApiUrl)

Pos = 1 
sPat = U)<contributor><name>(.*)</name><email>(.*)</email><snx:userid.*>(.*)</snx:userid>.*</contributor>
While Pos := RegExMatch(sSource,sPat,sMatch,Pos+StrLen(sMatch)) {
	sMention := "<a class='vcard' data-userid='" . sMatch3 . "' href='https://" . PowerTools_ConnectionsRootUrl . "/profiles/html/profileView.do?userid=" . sMatch3 . "'>@" . sMatch1 . "</a>"
	sHtml = %sHtml% %sMention%
}
; Copy to clipboard
WinClip.SetHtml(sHtml) 
TrayTipAutoHide("Copy Mentions", "Mentions were copied to the clipboard in RTF!")
}

; ----------------------------------------------------------------------
Html2Emails(sHtml){
; Extract emails from mailto links
; Syntax:
;    sEmailList :=Html2Emails(sHtml)
; Emails are separated by ;
sPat = <a href="mailto:(.*?)" 
Pos = 1 
While Pos := RegExMatch(sHtml,sPat,sEmail,Pos+StrLen(sEmail)){
    sEmailList = %sEmailList%;%sEmail1%
}
return SubStr(sEmailList,2) ; remove first ;
} ; eof

; ----------------------------------------------------------------------
GetInsLoc(){
OnMessage(0x44, "OnInsLocMsgBox")
MsgBox 0x23, Insert Toc, Select where the TOC shall be located:
OnMessage(0x44, "")
sInsLoc = top ; Default
IfMsgBox, No
	sInsLoc = cursor	
IfMsgBox, Cancel
	sInsLoc = bottom
return sInsLoc
}
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

ClipPasteText(Text){
global PT_wc
PT_wc.iSetText(Text)
PT_wc.iPaste()
}
; ----------------------------------------------------------------------

ClipPasteHtml(Html){
global PT_wc
PT_wc.iSetHtml(Html)
PT_wc.iPaste()
;MsgBox %Html% ; DBG
}



; ##################################################################################
; # This #Include file was generated by Image2Include.ahk, you must not change it! #
; ##################################################################################
Create_ConnextEnhancer_ico(NewHandle := False) {
Static hBitmap := 0
If (NewHandle)
   hBitmap := 0
If (hBitmap)
   Return hBitmap
VarSetCapacity(B64, 88132 << !!A_IsUnicode)
B64 := "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"
B64 .= "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"
B64 .= "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"
B64 .= "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"
B64 .= "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"
B64 .= "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"
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return False
VarSetCapacity(Dec, DecLen, 0)
If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
   Return False
; Bitmap creation adopted from "How to convert Image data (JPEG/PNG/GIF) to hBITMAP?" by SKAN
; -> http://www.autohotkey.com/board/topic/21213-how-to-convert-image-data-jpegpnggif-to-hbitmap/?p=139257
hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
Return hBitmap
}
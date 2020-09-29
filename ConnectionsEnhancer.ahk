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
LastCompiled = 20200806155132
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

Menu, ConNextEnhancerReadMenu, add, &Contact, :SubMenuContact
Menu, ConnectionsEnhancerMenu,Add ; Separator
Menu, ConnectionsEnhancerMenu, add, &Contact, :SubMenuContact
Menu, ConnectionsEnhancerMenu, add, Personalize Mentions (Win+1), PersonalizeMentions

Menu, ConNextEnhancerReadMenu, add, (&Blog) Likers to, :SubMenuContactLikers
Menu, ConNextEnhancerReadMenu, add, &Event to, :SubMenuEvent

Menu, ConNextEnhancerReadMenu, add, (Profile Search) Copy &Mentions, MenuItemProfileSearchToMentions
Menu, ConNextEnhancerReadMenu, add, (Profile Search) Copy &Table of Mentions with Profile pictures, MenuItemCopyUserTable

Menu, ConnectionsEnhancerMenu, Add ; Separator
Menu, ConNextEnhancerReadMenu, add, Create New... (Win+N), CNCreateNew
Menu, ConNextEnhancerReadMenu, add, Download Html, Connections_DownloadHtml
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
		Menu, ConNextEnhancerReadMenu, Enable, (&Blog) Likers to
	Else
		Menu, ConNextEnhancerReadMenu, Disable, (&Blog) Likers to

	If InStr(sUrl,"&eventInstUuid=") 
		Menu, ConNextEnhancerReadMenu, Enable, &Event to
	Else
		Menu, ConNextEnhancerReadMenu, Disable, &Event to

	If InStr(sUrl, PowerTools_ConnectionsRootUrl . "/profiles/html/") {
		Menu, ConNextEnhancerReadMenu, Enable, (Profile Search) Copy &Mentions 
		Menu, ConNextEnhancerReadMenu, Enable, (Profile Search) Copy &Table of Mentions with Profile pictures 
	}
	Else {
		Menu, ConNextEnhancerReadMenu, Disable, (Profile Search) Copy &Mentions 
		Menu, ConNextEnhancerReadMenu, Disable, (Profile Search) Copy &Table of Mentions with Profile pictures
	}

	Menu, ConNextEnhancerReadMenu, Show
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
sHtmlMentions := CNEmails2Mentions(sEmailList)
WinClip.SetHtml(sHtmlMentions)
TrayTipAutoHide("People Connector","Mentions were copied to clipboard in RTF!")   
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
MenuItemCopyUserTable:
sUrl := GetActiveBrowserUrl()
sHtml := CNGet(sUrl) ; TBC

sHtml := GenerateUserTable(sHtml)
WinClip.SetHtml(sHtml)
TrayTipAutoHide("Copy User Table", "Html Table was copied to the clipboard!")
return

; ----------------------------------------------------------------------
MenuItemUpdateUserTable:
If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_enhancer_generate_nice_user_table"
	return
}
sUrl := Clipboard
If !RegExMatch(sUrl,"^https?://" . PowerTools_ConnectionsRootUrl  "/profiles/html/") {
	MsgBox 0x1031,Connections Enhancer, Copy the search url to your clipboard, go back to the page to edit and press ok!
	IfMsgBox Cancel
		return
}
sUrl := Clipboard
If !RegExMatch(sUrl,"^https?://" . PowerTools_ConnectionsRootUrl  "/profiles/html/") {
	;MsgBox 0x10,Connections Enhancer, Clipboard does not match url with 'connext.conti.de/profiles/html/'!
	TrayTip, Update User Table, Clipboard does not match url with 'connext.conti.de/profiles/html/'!,,0x3
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
If InStr(sUrl,PowerTools_ConnectionsRootUrl . "/wikis/")
	return RegExMatch(sURL, PowerTools_ConnectionsRootUrl . "/wikis/.*/edit") || RegExMatch(sURL, PowerTools_ConnectionsRootUrl . "/wikis/.*/create")
Else If InStr(sUrl, PowerTools_ConnectionsRootUrl . "/blogs/")
	return InStr(sUrl, PowerTools_ConnectionsRootUrl . "/blogs/roller-ui/authoring/weblog.do") ; method edit or create
Else
	return InStr(sUrl, PowerTools_ConnectionsRootUrl . "forums/html/topic?") or InStr(sUrl, PowerTools_ConnectionsRootUrl . "/forums/html/threadTopic?") or  RegExMatch(sURL, PowerTools_ConnectionsRootUrl . "/forums/html/forum\?id=.*showForm=true")
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
} Else If RegExMatch(sUrl,"//" . PowerTools_ConnectionsRootUrl . "/blogs/([^/\?&]*)",sBlogId) {
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
	sApiUrl = https://connext.conti.de/wikis/basic/anonymous/api/wiki/%sWikiLabel%/navigation/%sPageLabel%/entry?includeSiblings=true&format=json
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
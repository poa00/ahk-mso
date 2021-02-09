; Hotstrings collection
; Support Rich Text Format e.g. Links
; Author: Thierry Dalon
; See user documentation here: https://connext.conti.de/blogs/tdalon/entry/TextExpander_ahk
; Code Project Documentation is available on ContiSource GitHub here: http://github.conti.de/ContiSource/ahk

; Source : http://github.conti.de/ContiSource/ahk/blob/master/TextExpander.ahk
; Calls: Lib/Clip

#SingleInstance force ; for running from editor
#Include <Clip>
#Include <Login>

; ##### Hotstrings ####

#Hotstring c1 ; Do not conform to typed case

::nwsac::{#}nws_adventcalendar2020

::mfg:: ; MfG
(
Mit freundlichen Gruessen
Thierry Dalon
)

::kr:: ; Kind Regards
(
Kind Regards
Thierry
)

::br:: ; Best Regards
(
Regards Regards
Thierry Dalon
)

::vg:: ; Viele Gruesse
(
Viele Gruesse
Thierry
)

::lg:: ; Liebe Gruesse
(
Liebe Gruesse
Thierry
)


; ##### Rich Hotstrings ####
::teamsme:: ; Teams Me Chat Link (rtf)
sText = Chat with me in Teams
sLink :=  "https://teams.microsoft.com/l/chat/0/0?users=Thierry.Dalon@continental-corporation.com"
Clip_PasteHtml(sLink,sText,True)	
return

::tagging_rules:: ; ConNext Tagging Rules (RTF Link)
sText = ConNext Tagging Rules
sLink := "https://connext.conti.de/wikis/home/wiki/W10f67125ddc8_42e1_a6da_0a8e6a1cd541/page/ConNext%20Basics?section=tagging"
Clip_PasteHtml(sLink,sText,True)
return


::nws_s:: ; NWS Search (RTF Link)
sText = NWS Search
sLink = http://links.conti.de/nws_search
Clip_PasteHtml(sLink,sText,True)	
return

::nws_ss:: ; NWS Social Support (RTF Link)
sText = NWS Social Support
sLink = http://links.conti.de/nws_socialsupport
Clip_PasteHtml(sLink,sText,True)
return

::nws_sh:: ; NWS Search Help Link (rtf)
sText = NWS Search
sLink := "https://connext.conti.de/wikis/home/wiki/W10f67125ddc8_42e1_a6da_0a8e6a1cd541/page/Help%20on%20Start%20Page%20with%20Search"
Clip_PasteHtml(sLink,sText,True)
return

::mo_ty:: ; MO Thank you image
sHtml := ConNext_Kudos2Html("thank_you")
sHtml = <p style="text-align: center;">%sHtml%<p>
Clip_PasteHtml(sHtml)
return

; ################### PowerTools RTF Hotstrings ##########################
::Teams.ahk:: ; Teams.ahk (RTF Link)
sText = ahk/Lib/Teams.ahk -> (function)
sLink := "https://github.com/tdalon/ahk/blob/master/Lib/Teams.ahk"
Clip_PasteHtml(sLink,sText,True)
return

::PT_TS:: ; TeamsShortcuts (RTF Link)
sText = Teams Shortcuts
sLink := "https://tdalon.github.io/ahk/Teams-Shortcuts"
Clip_PasteHtml(sLink,sText,True)
return

::PT_T:: ; Teamsy (RTF Link)
sText = Teamsy
sLink := "https://tdalon.github.io/ahk/Teamsy"
Clip_PasteHtml(sLink,sText,True)
return

::PT_CE:: ; Connections Enhancer (RTF Link)
sText = Connections Enhancer
sLink := "https://tdalon.github.io/ahk/Connections-Enhancer"
Clip_PasteHtml(sLink,sText,True)
return

::PT_N:: ; NWS PowerTool (RTF Link)
sText = NWS PowerTool
sLink := "https://tdalon.github.io/ahk/NWS-PowerTool"
Clip_PasteHtml(sLink,sText,True)
return

::PT_OS:: ; OutlookShortcuts (RTF Link)
sText = Outlook Shortcuts
sLink := "https://tdalon.github.io/ahk/Outlook-Shortcuts"
Clip_PasteHtml(sLink,sText,True)
return




; ##### Hotkeys ####
; Win + q
#q:: ; My Work Email
SendInput Thierry.Dalon@continental-corporation.com
return

; alt + q
!q:: ; My private Email 
SendInput thierry.dalon@gmail.com
return

; Alt + u
!u:: ; uid@contiwan.com
;OfficeUid := People_GetMyOUid()
SendInput %A_UserName%@contiwan.com
return

!p:: ; <--- Enter Password
sPassword := Login_GetPassword()
SendInput %sPassword%
return

; Ctrl+,
^,:: ; CurrentDate
FormatTime, CurrentDateTime,, yyyy-MM-dd
SendInput %CurrentDateTime%
return


; Ctrl+.
^.:: ; DatePicker
DatePicker(sDate)
If !sDate ; empty= cancel
    return

FormatTime, sDate, %sDate%, yyyy-MM-dd
SendInput %sDate%

return


; ----------------------- SUBFUNCTIONS -------------------------------
DatePicker(ByRef DatePicker){
	
Gui, +LastFound 
gui_hwnd := WinExist()
Gui, Add, MonthCal, vDatePicker, 4
Gui, Add, Button, Default , &OK
Gui Add, Button, x+0, Cancel
Gui, Show , , Date Picker Calendar
WinWaitClose, AHK_ID %gui_hwnd%
return

ButtonOK:
Gui, submit ;, nohide
Gui, Destroy
;Gui, Hide

return

GuiEscape:
ButtonCancel:
GuiClose:
Gui, Destroy
return
}
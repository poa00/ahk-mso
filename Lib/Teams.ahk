#Include <People>
#Include <UriDecode>

Emails2TeamsChatDeepLink(sEmailList, askOpen:= true){
sLink := "msteams://teams.microsoft.com/l/chat/0/0?users=" . StrReplace(sEmailList, ";",",")

If InStr(sEmailList,";") { ; Group Chat
    InputBox, sTopicName, Enter Group Chat Name,,,,100
    if (ErrorLevel=1)
	    sLinkDisplayText = Group Chat Link
} Else {
    InputBox, sTopicName, Chat Link Display Name,,,,100
    if (ErrorLevel=1)
	    sLinkDisplayText = Chat Link
}
sLink := sLink . "&topicName=" . StrReplace(sTopicName, ":", "")
sLinkDisplayText = %sTopicName% 
If InStr(sEmailList,";")
    sLinkDisplayText = %sLinkDisplayText% (Group Chat Link)
sHtml = <a href="%sLink%">%sLinkDisplayText%</a>
WinClip.SetHtml(sHtml) 
WinClip.SetText(sLink) 

If askOpen {
	MsgBox 0x1034,People Connector, Teams link was copied to the clipboard. Do you want to open the Group Chat?
	IfMsgBox Yes
		Run, %sLink% 
}

} ; eofun

; -------------------------------------------------------------------------------------------------------------------


; Called by TeamsShortcuts, PeopleConnector

Link2TeamsFav(sUrl:="",FavsDir:="",sFileName :="") {
If (sUrl="") or !(sUrl ~= "https://teams.microsoft.com/*") {
	
    If RegExMatch(sLink,"https://teams.microsoft.com/l/channel/[^/]*/([^/]*)\?.*",sChannelName) 
	    linktext = %sChannelName1% (Channel)
    Else
        linktext = Team Name (Team)
    
    InputBox, sUrl , Teams Fav Target, Paste Teams Link:, , 640, 125,,,,, %linktext%
	If ErrorLevel
		return
}

sUrl:= StrReplace(sUrl,"https:","msteams:")
; folder does not end with filesep

If !(FavsDir) {
    sKeyName := "TeamsFavsDir"
    RegRead, StartingFolder, HKEY_CURRENT_USER\Software\PowerTools, %sKeyName%
    FileSelectFolder, FavsDir , *%StartingFolder%, ,Select folder to store your Teams Shortcut:
    If ErrorLevel
        return
}

If !(sFileName) {
    InputBox, sFileName , Teams Fav File name, Enter the File name:, , 300, 125
    If ErrorLevel
        return
}
sFile := FavsDir . "\" . sFileName . ".url"

;FileDelete %sFile%
IniWrite, %sUrl%, %sFile%, InternetShortcut, URL

; Add icon file:
TeamsExe = C:\Users\%A_UserName%\AppData\Local\Microsoft\Teams\current\Teams.exe
IniWrite, %TeamsExe%, %sFile%, InternetShortcut, IconFile
IniWrite, 0, %sFile%, InternetShortcut, IconIndex

SplitPath, sFile, sFileName, FavsDir
RegWrite, REG_SZ, HKEY_CURRENT_USER\Software\PowerTools, %sKeyName%, %FavsDir%

;Run %FavsDir%

} 

; -------------------------------------------------------------------------------------------------------------------



TeamsFavsSetDir(){
sKeyName := "TeamsFavsDir"
RegRead, StartingFolder, HKEY_CURRENT_USER\Software\PowerTools, %sKeyName%
FileSelectFolder, sKeyValue , StartingFolder, Options, Select folder for your Teams Favorites:
If ErrorLevel
    return
RegWrite, REG_SZ, HKEY_CURRENT_USER\Software\PowerTools, %sKeyName%, %sKeyValue%
return sKeyValue
} ; eofun


; ----------------------------------------------------------------------
Emails2TeamsFavs(sEmailList){
RegRead, FavsDir, HKEY_CURRENT_USER\Software\PowerTools, TeamsFavsDir
If ErrorLevel {
    FavDirs := TeamsFavsSetDir()
    If FavDirs = ""
        return    
}

FileSelectFolder, FavsDir , *%FavsDir%, ,Select folder to store your Teams Contact Shortcuts:
If ErrorLevel
    return
Loop, parse, sEmailList, ";"
{
    Email2TeamsFavs(A_LoopField,FavsDir)
}
Run %FavsDir%	
} ; eofun
; ----------------------------------------------------------------------


Email2TeamsFavs(sEmail,FavsDir){
; Get Firstname
sName := RegExReplace(sEmail,"\..*" ,"")
StringUpper, sName, sName , T

sUrl = https://teams.microsoft.com/l/chat/0/0?users=%sEmail% 
Link2TeamsFav(sUrl,FavsDir,"Chat " + sName)
    
; create shortcut
sLnk := RegExReplace(sFile,"\..*",".lnk")
FileCreateShortcut, %sFile%, %sLnk%,,,, %userprofile%\AppData\Local\Microsoft\Teams\current\Teams.exe 


; folder does not end with filesep
sFile := FavsDir . "\Call " . sName . ".vbs"
; write code

EnvGet userprofile, userprofile
sText = CreateObject("Wscript.Shell").Run "%userprofile%\AppData\Local\Microsoft\Teams\current\Teams.exe callto:%sEmail%"

; Create empty File
FileDelete %sFile%
FileAppend , %sText%, %sFile%

; create shortcut
sLnk := RegExReplace(sFile,"\..*",".lnk")
FileCreateShortcut, %sFile%, %sLnk%,,,, %userprofile%\AppData\Local\Microsoft\Teams\current\Teams.exe

} ; eofun

; NOT USED ########################
Emails2TeamsFavGroupChat(sEmailList){
RegRead, FavsDir, HKEY_CURRENT_USER\Software\PowerTools, TeamsFavsDir
If ErrorLevel {
    FavDirs := TeamsFavsSetDir()
    If FavDirs = ""
        return    
}

FileSelectFolder, FavsDir , *%StartingFolder%, ,Select folder to store your Teams Shortcut:
If ErrorLevel
    return

sUrl := "https://teams.microsoft.com/l/chat/0/0?users=" . StrReplace(sEmailList, ";",",")

InputBox, sName, Enter Chat Group name,,,,100
if ErrorLevel
    return

sName := StrReplace(sName, ":", "")
sLink := sLink . "&topicName=" . sName
 
Link2TeamsFav(sUrl,FavsDir,"Group Chat -" . sName)
} ; eofun
; ----------------------------------------------------------------------

Emails2TeamMembers(EmailList,TeamLink:=""){
; Syntax: 
;     Emails2TeamMembers(EmailList,TeamLink*)
; EmailList: String of email adressess separated with a ;
; TeamLink: (String) optional. If not passed, user will be asked via inputbox
;  e.g. https://teams.microsoft.com/l/team/19%3a12d90de31c6e44759ba622f50e3782fe%40thread.skype/conversations?groupId=640b2f00-7b35-41b2-9e32-5ce9f5fcbd01&tenantId=8d4b558f-7b2e-40ba-ad1f-e04d79e6265a
If (!TeamLink) {
    InputBox, TeamLink , Team Link, Enter Team Link:,,640,125 
    if ErrorLevel
	    return
}
; 
sPat = \?groupId=(.*)&tenantId=(.*)
If !(RegExMatch(TeamLink,sPat,sId)) {
	MsgBox 0x10, Error, Provided Url does not match a Team Link!
	return
}
sGroupId := sId1
sTenantId := sId2

; Create csv file with list of emails
CsvFile = %A_Temp%\email_list.csv
If FileExist(CsvFile)
    FileDelete, %CsvFile%
sText := StrReplace(EmailList,";","`n")
FileAppend, %sText%,%CsvFile% 
; Create .ps1 file
PsFile = %A_Temp%\Teams_AddUsers.ps1
; Fill the file with commands
If FileExist(PsFile)
    FileDelete, %PsFile%

RegRead, OfficeUid, HKEY_CURRENT_USER\Software\PowerTools, OfficeUid

If !OfficeUid {
    OfficeUid := AD_GetUserField("sAMAccountName=" . A_UserName, "mailNickname") ; mailNickname - office uid 
    RegWrite, REG_SZ, HKEY_CURRENT_USER\Software\PowerTools, OfficeUid, %OfficeUid%
}
sText = Connect-MicrosoftTeams -TenantId %sTenantId% -AccountId %OfficeUid%@contiwan.com
MsgBox %sText% ; DBG

; Connect-MicrosoftTeams -TenantId 8d4b558f-7b2e-40ba-ad1f-e04d79e6265a -AccountId $env:UserName@contiwan.com
sText = %sText%`nImport-Csv -header email -Path "%CsvFile%" | foreach{Add-TeamUser -GroupId "%sGroupId%" -user $_.email}
MsgBox %sText% ; DBG
FileAppend, %sText%,%PsFile%

; Run it
RunWait, PowerShell.exe -ExecutionPolicy Bypass -Command %PsFile%,, Hide

}

; ----------------------------------------------------------------------
TeamsLink2TeamName(TeamLink){
If (!TeamLink) {
    InputBox, TeamLink , Team Link, Enter Team Link:,,640,125 
    if ErrorLevel
	    return
}

RegRead, TeamsPowerShell, HKEY_CURRENT_USER\Software\PowerTools, TeamsPowerShell
If !TeamsPowerShell {
    InputBox, TeamName , Team Link, Enter Team Link:,,640,125 
    if ErrorLevel
	    return
    return TeamName
}
; 
sPat = \?groupId=(.*)&tenantId=(.*)
If !(RegExMatch(TeamLink,sPat,sId)) {
	MsgBox 0x10, Error, Provided Url does not match a Team Link!
	return
}
sGroupId := sId1
sTenantId := sId2

; Create .ps1 file
PsFile = %A_Temp%\Teams_AddUsers.ps1
; Fill the file with commands
If FileExist(PsFile)
    FileDelete, %PsFile%

RegRead, OfficeUid, HKEY_CURRENT_USER\Software\PowerTools, OfficeUid
sText = Connect-MicrosoftTeams -TenantId %sTenantId% -AccountId %OfficeUid%@contiwan.com
; Connect-MicrosoftTeams -TenantId 8d4b558f-7b2e-40ba-ad1f-e04d79e6265a -AccountId $env:UserName@contiwan.com
sText = %sText%`nGet-Team -GroupId %sGroupId%
FileAppend, %sText%,%PsFile%

; Get a temporary file path
tempFile := A_Temp "\PowerShell_Output.txt"                           ; "

; Run the console program hidden, redirecting its output to
; the temp. file (with a program other than powershell.exe or cmd.exe,
; prepend %ComSpec% /c; use 2> to redirect error output), and wait for it to exit.
RunWait, powershell.exe -ExecutionPolicy Bypass -Command %PsFile% > %tempFile%,, Hide

; Read the temp file into a variable and then delete it.
FileRead, sOutput, %tempFile%
FileDelete, %tempFile%

; Display the result.
If RegExMatch(sOutput,"DisplayName.*:(.*)",sMatch)
    sName := sMatch1
return sName
}

; ----------------------------------------------------------------------
TeamLink2Users(TeamLink){
If (!TeamLink) {
    InputBox, TeamLink , Team Link, Enter Team Link:,,640,125 
    if ErrorLevel
	    return
}

sPat = \?groupId=(.*)&tenantId=(.*)
If !(RegExMatch(TeamLink,sPat,sId)) {
	MsgBox 0x10, Error, Provided Url does not match a Team Link!
	return
}
sGroupId := sId1
sTenantId := sId2

; Create .ps1 file
PsFile = %A_Temp%\Teams_AddUsers.ps1
; Fill the file with commands
If FileExist(PsFile)
    FileDelete, %PsFile%
RegRead, OfficeUid, HKEY_CURRENT_USER\Software\PowerTools, OfficeUid
If !OfficeUid {
    OfficeUid := AD_GetUserField("sAMAccountName=" . A_UserName, "mailNickname") ; mailNickname - office uid 
    RegWrite, REG_SZ, HKEY_CURRENT_USER\Software\PowerTools, OfficeUid, %OfficeUid%
}
sText = Connect-MicrosoftTeams -TenantId %sTenantId% -AccountId %OfficeUid%@contiwan.com
; Connect-MicrosoftTeams -TenantId 8d4b558f-7b2e-40ba-ad1f-e04d79e6265a -AccountId $env:UserName@contiwan.com
sText = %sText%`nGet-TeamUser -GroupId %sGroupId%
FileAppend, %sText%,%PsFile%

; Get a temporary file path
tempFile := A_Temp "\PowerShell_Output.txt"                           ; "

; Run the console program hidden, redirecting its output to
; the temp. file (with a program other than powershell.exe or cmd.exe,
; prepend %ComSpec% /c; use 2> to redirect error output), and wait for it to exit.
RunWait, powershell.exe -ExecutionPolicy Bypass -Command %PsFile% > %tempFile%,, Hide

; Read the temp file into a variable and then delete it.
FileRead, sOutput, %tempFile%
;FileDelete, %tempFile%

Run %tempFile% ; DBG
}

; -------------------------------------------------------------------------------------------------------------------

TeamsSmartReply(sHtml := "",CopyLink:= True){
GetKeyState, KeyState, Ctrl
If (KeyState = "D") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/teams_smart_reply"
	return
}
If !sHtml
    sHtml := GetSelection("html")
;MsgBox %sHtml% ; DBG
; Extract Teams Link

; Full selection parent topic
sPat = U)<div>&lt;https://teams.microsoft.com/(.*;createdTime=.*)&gt;</div>
If (RegExMatch(sHtml,sPat,sDiv)) {
    sPat = U)(<span>.*</span><span>.*</span>).*<div data-tid="messageBodyContainer">
	If (RegExMatch(sHtml,sPat,sTitle)) {
        sTeamLink = https://teams.microsoft.com/%sDiv1%
        sNewTitle = <a href="%sTeamLink%">%sTitle1%</a>
		sHtml := StrReplace(sHtml,sTitle1,sNewTitle)
	}
} Else {
    ; Full selection reply
    sPat = U)<div><a href="https://teams.microsoft.com/(.*;createdTime=.*)">.*</a>
    RegExMatch(sHtml,sPat,sDiv)
}

If (sDiv) { ; whole thread was selected
	
	;sPat = U)<!--StartFragment--><div>.*<div>(.*)</div></div>
	
    sHtml := RegExReplace(sHtml,"<dl><dd><div>&lt;.*&gt;</div></dd></dl>","")
    sHtml := RegExReplace(sHtml,"<div>&lt;.*&gt;</div>","")

    sHtml := StrReplace(sHtml,"<body>","<body><blockquote>")
    sHtml := StrReplace(sHtml,"</body>","</blockquote></body>")

    sPat = Us)<body>(.*)</body>	
    If RegExMatch(sHtml,sPat,sHtmlBody)
        sHtml := sHtmlBody1
    WinClip.SetHTML(sHtml)

    Send {r}
    Sleep 500
    Send ^v
    Sleep 500
    ;Send {Backspace 2} ; remove strange <> / put focus at the bottom ready to type ahead
    Send +{Enter}
    Send +{Enter}


} Else { ; Partial selection
    If (CopyLink) {
        rc := TeamsCopyLink()
        If (rc) {
            sHtmlTitle := WinClip.GetHTML()
            sPat = Us)<body>(.*)</body>
            If RegExMatch(sHtmlTitle,sPat,sHtmlTitle)
                sHtmlTitle := sHtmlTitle1
        }
    }
    MsgBox %sHtmlTitle%
    Send {r}
    Sleep 1000

    If (sHtml){
        sPat = Us)<body>(.*)</body>	
        If RegExMatch(sHtml,sPat,sHtmlBody)
            sHtml := sHtmlBody1
        sHtml = (...)%sHtml%(...)
        sHtml = <blockquote>%sHtmlTitle%%sHtml%</blockquote>
        WinClip.SetHTML(sHtml)
        WinClip.Paste()
        Sleep 500
        ;PasteText(sText)
    }
    ; escape quote block
    Send +{Enter}
    Send +{Enter}
}
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
TeamsCopyLink(){
Send {Enter}
Send {Tab}
Send {Enter}
;Sleep 500
;Send {Down 2}

MsgBox 0x1001, Wait..., Wait for Copy Link. Press OK to continue. 
IfMsgBox Cancel
    return False
return true
} ; eofun


; -------------------------------------------------------------------------------------------------------------------
TeamsLink2Text(sLink){
sPat = https://teams.microsoft.com/[^>"]*
RegExMatch(sLink,sPat,sLink)
sLink := uriDecode(sLink)
; Link to Teams Channel
; example: https://teams.microsoft.com/l/channel/19%3a16ff462071114e31bd696aa3a4e34500%40thread.skype/DOORS%2520Attributes%2520List?groupId=cd211b48-2e8b-4b60-b5b0-e584a0cf30c0&tenantId=8d4b558f-7b2e-40ba-ad1f-e04d79e6265a
If (RegExMatch(sLink,"U)https://teams.microsoft.com/l/channel/[^/]*/([^/]*)\??groupId=(.*)&",sMatch)) {
    sChannelName := sMatch1
    sTeamName := TeamsLink2TeamName(sLink)
	linktext = %sTeamName% Team - %sChannelName% (Channel)
	;linktext := StrReplace(linktext,"%2520"," ")		
; Link to Teams
; example: https://teams.microsoft.com/l/team/19%3a12d90de31c6e44759ba622f50e3782fe%40thread.skype/conversations?groupId=640b2f00-7b35-41b2-9e32-5ce9f5fcbd01&tenantId=8d4b558f-7b2e-40ba-ad1f-e04d79e6265a
} Else If (RegExMatch(sLink,"https://teams.microsoft.com/team/.*groupId=(.*)&",sMatch)) {
; https://teams.microsoft.com/l/team/19:c1471a18bae04cf692b8da7e9738df3e@thread.skype/conversations?groupId=56bc81d8-db27-487c-8e4f-8d5ea9058663&tenantId=8d4b558f-7b2e-40ba-ad1f-e04d79e6265a    
    sTeamName := TeamsLink2TeamName(sLink)
    InputBox, linktext , Display Link Text, Enter Team name:,,640,125,,,,, Link to Teams Team
	If ErrorLevel ; Cancel
		return
}
return linktext  
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
TeamsMessageLink2Html(sLink){
sPat = https://teams.microsoft.com/[^>"]*
RegExMatch(sLink,sPat,sLink)
sLink := uriDecode(sLink)
RegExMatch(sLink,"https://teams.microsoft.com/l/message/(.*)\?tenantId=(.*)&groupId=(.*)&teamName=(.*)&channelName=(.*)&",sMatch) 
; https://teams.microsoft.com/l/message/19:fbfd482e3b544af387cbe6db65846796@thread.skype/1582627348656?tenantId=8d4b558f-7b2e-40ba-ad1f-e04d79e6265a&amp;groupId=56bc81d8-db27-487c-8e4f-8d5ea9058663&amp;parentMessageId=1582627348656&amp;teamName=GUIDEs&amp;channelName=Technical Questions&amp;createdTime=1582627348656  

; Prompt for Type of link: Team, Channel or Conversation
Choice := ButtonBox("Teams Link:Setting:TocStyle","Do you want a link to the:","Message|Channel|Team")
If ( Choice = "ButtonBox_Cancel") or ( Choice = "Timeout")
    return
Switch Choice
{
Case "Team":
    linktext = sMatch4 (Team)
    sLink = https://teams.microsoft.com/l/team/%sMatch1%/conversations?groupId=%sMatch3%&tenantId=%sMatch2%
Case "Channel":
    linktext = sMatch4 - sMatch5 (Channel)
    sLink = https://teams.microsoft.com/l/channel/%sMatch1%/%sMatch5%?groupId=%sMatch3%&tenantId=%sMatch2%
    ; https://teams.microsoft.com/l/channel/19%3ab4371b8d10234ac9b4e0095ace0aae8e%40thread.skype/Best%2520Work%2520Hacks?groupId=56bc81d8-db27-487c-8e4f-8d5ea9058663&tenantId=8d4b558f-7b2e-40ba-ad1f-e04d79e6265a
Case "Message":
    linktext = sMatch4 - sMatch5 - Message
}
sHtml = <a href="sLink">%linktext%</a>
return sHtml
}

; -------------------------------------------------------------------------------------------------------------------

MenuCb_ToggleSettingTeamsPowerShell(ItemName, ItemPos, MenuName){
GetKeyState, KeyState, Ctrl
If (KeyState = "D"){
    sUrl := "https://connext.conti.de/wikis/home/wiki/Wc4f94c47297c_42c8_878f_525fd907cb68/page/Teams%20PowerShell%20Setup"
    Run, "%sUrl%"
	return
}
RegRead, TeamsPowerShell, HKEY_CURRENT_USER\Software\PowerTools, TeamsPowerShell
TeamsPowerShell := !TeamsPowerShell 
If (TeamsPowerShell) {
 	Menu,%MenuName%,Check, %ItemName%	 
    OfficeUid := AD_GetUserField("sAMAccountName=" . A_UserName, "mailNickname") ; mailNickname - office uid 
    RegWrite, REG_SZ, HKEY_CURRENT_USER\Software\PowerTools, OfficeUid, %OfficeUid%
    ;sWinUid := AD_GetUserField("mail=" . sEmail, "sAMAccountName")  ;- login uid
} Else
    Menu,%MenuName%,UnCheck, %ItemName%	 

RegWrite, REG_SZ, HKEY_CURRENT_USER\Software\PowerTools, TeamsPowerShell, %TeamsPowerShell%

}
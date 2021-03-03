; Calls JiraGet to get Jira issue title from Jira issue link
#Include <Jira>
; Calls Teams_Link2Text
#Include <Teams>
#Include <uriDecode>
#Include <SharePoint>
; Calls: CNGetTitle, Connections_Link2Text, CNFormatImg
#Include <Connections>
#Include <Confluence>
#Include <Clip>


Link2Text(sLink) {
; Syntax:
; 	linktext := Link2Text(sLink)
; Called by: IntelliHtml
; Calls: Connections_Link2Text, Teams_Link2Text


If InStr(sLink,"https://teams.microsoft.com/") {
	return Teams_Link2Text(sLink)
}
If Connections_IsUrl(sLink) {
    return Connections_Link2Text(sLink)
}

If InStr(sLink,"/gist/") {
	InputBox, linktext, Enter Gist Description (Link Text),,,,100
	if ErrorLevel
		return
} Else If RegExMatch(sLink,".*confluence.*&title=(.*)",sTitle) {
	linktext = %sTitle1%
	linktext := StrReplace(linktext,"+"," ")
	; Confluence Page
	; example: http://confluence.conti.de:8090/pages/viewpage.action?spaceKey=projectCFTI&title=Roadmap+-+Process+and+Method+Improvement
}

; beautify IMS/MKS links
; http://%MKSSI_HOST%:%MKSSI_PORT%/si/viewrevision?projectName=%PROJECT%&selection=%MEMBER%
Else If RegExMatch(sLink,"/si/viewrevision\?projectName=(?P<Project>.*)&revision=(?P<Revision>.*)&selection=(?P<Member>.*)",OutputVar) 
	linktext = %OutputVarMember% (Rev %OutputVarRevision%)
Else If RegExMatch(sLink,"/si/viewrevision\?projectName=(?P<Project>.*)&selection=(?P<Member>.*)",OutputVar) 
	linktext = %OutputVarMember%
Else If RegExMatch(sLink,"/si/viewrevision\?projectName=(?P<Project>.*)",OutputVar) 	
	linktext = %OutputVarProject%
Else If RegExMatch(sLink,"/im/issues\?selection=(.*)",OutputVar) 
	linktext = %OutputVar1%
Else If RegExMatch(sLink,"/im/viewissue\?selection=(.*)",OutputVar) 
	linktext = %OutputVar1%

; Jira Issue link
Else If RegExMatch(sLink,"/browse/(?P<IssueKey>.*)$",OutputVar) {
	sLink :=StrReplace(sLink,"/browse/","/rest/api/latest/issue/")
	sLink := sLink . "?fields=summary"
	sResponse := Jira_Get(sLink)
	sPat = "summary":"(.*?)"
	If RegExMatch(sResponse,sPat,sSummary)
		linktext = %OutputVarIssueKey%: %sSummary1%
	Else
		linktext = %OutputVarIssueKey%
}
Else If RegExMatch(sLink,"/browse/(?P<ProjectKey>[A-Z]*)-(?P<IssueNb>\d*)$",OutputVar) {
	linktext = %OutputVarProjectKey%-%OutputVarIssueNb%
; Jira ServiceDesk
; https://jira-it.zone2.agileci.conti.de/plugins/servlet/desk/portal/1/PMTJSD-5745
} Else If RegExMatch(sLink,"/desk/portal/1/(?P<IssueKey>.*)$",OutputVar) {
	sLink := RegExReplace(sLink,"https?://","")
	JiraRoot := RegExReplace(sLink,"/.*","")
	RestUrl = https://%JiraRoot%/rest/api/2/issue/%OutputVarIssueKey%?fields=summary
	sResponse := Jira_Get(RestUrl)
	sPat = "summary":"(.*?)"
	If RegExMatch(sResponse,sPat,sSummary)
		linktext = %OutputVarIssueKey%: %sSummary1%
	Else
		linktext = %OutputVarIssueKey%
}

; LL/RAI Issue Link
; https://aws1.conti.de/sites/LessonsLearned/Lists/LL/DispForm.aspx?ID=11629
Else If RegExMatch(sLink,"https://aws1.conti.de/sites/LessonsLearned/Lists/LL/DispForm.aspx\?ID=([\d]{1,})",sId) {
	linktext = LL-%sId1%
}
Else If RegExMatch(sWord,"https://aws1.conti.de/sites/LessonsLearned/Lists/RAI/DispForm.aspx\?id=([\d]{1,})",sId) {
	linktext = RAI-%sId1%
}
Else If RegExMatch(sLink,"https://web.microsoftstream.com/browse\?q=(.*)",sQuery) {
	linktext = Stream videos matching: %sQuery1% 
	linktext := StrReplace(linktext,"%23","#")		
} Else If RegExMatch(sLink,"/teams/team_[^/]*/[^/]*/(.*)",sMatch) {
; #TODO replace by SharePoint_IntelliHtml
	sMatch1 := uriDecode(sMatch1)
	linktext := StrReplace(sMatch1,"/"," > ") ; Breadcrumb navigation for Teams link to folder
	 
	TeamName := Teams_GetTeamName(sLink)
	If (TeamName != "") ; not empty
		linktext := TeamName . " > " . linktext
	
	If InStr(sLink,".") { ; file with extension
		FileName := RegExReplace(sMatch1,".*/","")
		Result := ListBox("IntelliPaste: File Link","Choose text to display",linktext . "|" . FileName,1)
		If Not (Result ="")
			linktext := Result
	}
	return linktext

} Else {
	
	; remove ending filesep 
	If (SubStr(sLink,0) == "\") or (SubStr(sLink,0) == "/") ; file or url
		sLink := SubStr(sLink,1,StrLen(sLink)-1)	
	
	
	; Truncate link to last /
	FoundPos := InStr(sLink, "/" , , 0) ; reverse search
	If FoundPos
    	linktext :=  SubStr(sLink, FoundPos+1)

	; Blogpost https://tdalon.blogspot.com/2020/08/executor-my-preferred-app-launcher.html
	If InStr(sLink,".blogspot.") {
		linktext := StrReplace(linktext,".html","")
	}
	; Removing ending ? e.g. https://continental.sharepoint.com/:p:/r/teams/team_10000778/Shared%20Documents/Explore/New%20Work%20Style%20%E2%80%93%20O365%20-%20Why%20using%20Teams.pptx?d=we1a512b97ed844fc92dd5a1d028ef827&csf=1&e=crdehv 
	linktext := RegExReplace(linktext, "\?.*$","")
	linktext := StrReplace(linktext,"-"," ")
		;linktext := StrReplace(linktext,"_"," ")
	linktext := StrReplace(linktext,"+"," ") ; for Confluence pages 
	; convert to title case only if all lower
	If RegExMatch(var,"^[a-z\-\s]*$")
		StringUpper, linktext, linktext , T ; upper case first letter Title case
	
	; Add specific prefix to linktext: UserVoice, CoachNet
	If InStr(sLink,".uservoice.")
		linktext = UserVoice: %linktext% 	
	
} ; End If

linktext := uriDecode(linktext) 

return linktext
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

; Intelli-Html feature
; Convert Input to nice formatted Html code to be pasted as rich-text/Html format
; Syntax:
;    sHtml := IntelliHtml(sInput,useico := True)
; No need to switch to HTML Source / Paste in Rich-Text Format

; Called by NWS.ahk - Triggered by Ins
; Calls CNFormatImg, Link2Text, Link2Ico
IntelliHtml(sInput,useico:=True){

If (RegExMatch(sInput,"[A-Z]:\\")) or (RegExMatch(sInput,"^\\\\"))  { ; local files or old SharePoint e.g. \\aws3.conti.de opened in File explorer 
	sInput := GetFileLink(sInput)
}


If SharePoint_IsSPUrl(sInput) {
	sInput := SharePoint_CleanUrl(sInput) ; keep name used as link for ico
	linktext := SharePoint_Link2Text(sInput)
	If InStr(linktext,">") ; use breadcrumbs
		sHtml := SharePoint_IntelliHtml(sInput)
	Else
		sHtml =	<a href="%sInput%">%linktext%</a>

	If RegExMatch(sInput,"/teams/(team_[^/]*)/",sMatch) {
		CsvFile = %A_ScriptDir%\Teams_list.csv
		If !FileExist(CsvFile) {
			CsvFile := Teams_ExportTeams() ; requires PowerShell
			If (CsvFile = "") 
				Goto, Cont1
		}
		MailNickName := sMatch1
		TeamName := ReadCsv(CsvFile,"MailNickName",MailNickName,"DisplayName")
		If Not (TeamName = "") {
			GroupId := ReadCsv(CsvFile,"MailNickName",MailNickName,"GroupId")
			TeamLink = https://teams.microsoft.com/l/team/?groupId=%GroupId%
			sHtml = <a href="%TeamLink%">%TeamName% Team</a>: %sHtml%
		}
	}
	Cont1:

} Else If InStr(sInput,"https://teams.microsoft.com/l/message/") {
	sHtml := Teams_MessageLink2Html(sInput)

} Else If RegExMatch(sInput,"/desk/portal/1/(?P<IssueKey>.*)$",OutputVar) {
	sInput := RegExReplace(sInput,"https?://","")
	JiraRoot := RegExReplace(sInput,"/.*","")
	RestUrl = https://%JiraRoot%/rest/api/2/issue/%OutputVarIssueKey%?fields=summary
	sResponse := Jira_Get(RestUrl)
	sPat = "summary":"(.*?)"
	If RegExMatch(sResponse,sPat,sSummary)
		linktext = %OutputVarIssueKey%: %sSummary1%
	Else
		linktext = %OutputVarIssueKey%
	
	; Convert Ticket Link to Issue link
	MsgBox, 0x24,IntelliPaste: Question, Do you want to convert Ticket link into an Issue link?	
	IfMsgBox Yes 
		sInput = https://%JiraRoot%/browse/%OutputVarIssueKey%
	sHtml = <a href="%sInput%">%linktext%</a>

} Else If RegExMatch(sInput,"(SD[\d]{1,})",sId) { ; SD HPSM issue
		sInput = https://hpsm-web.cw01.contiwan.com/sm/ess.do?ctx=docEngine&file=incidents&query=incident.id_percent3D_percent22%sId1%_percent22
		sInput := StrReplace(sInput,"_percent","%")
		linktext = %sId1%
		;https://hpsm-web.cw01.contiwan.com/sm/ess.do?ctx=docEngine&file=incidents&query=incident.id%3D%22SD422825518%22
} Else If RegExMatch(sInput,"\.png$") or RegExMatch(sInput,"\.gif$") or RegExMatch(sInput,"\.jpg$") or RegExMatch(sInput,"^https?://connext.conti.de/forums/ajax/download\?nodeId=") { ; Image files TODO: change to ConnectionsRootUrl
		sHtml = <img src="%sInput%"/></img>
		sHtml := CNFormatImg(sHtml)
} Else If InStr(sInput,".mp4") or InStr(sInput,".wmv") {
	sHtml = <p dir="ltr" style="text-align: center;"><video controls="controls"  width="95_percent" tabindex="-1" class="HTML5Video"><source src="%sInput%"/></video></p>
	sHtml := StrReplace(sHtml,"_percent","%")
} Else If InStr(sInput,"_player.html") { ; Camtasia player file
	If WinActive("Blog")  
		sHtml = <p dir="ltr" style="text-align: center;"><iframe  allowfullscreen="true"  frameborder="0" src="%sInput%" width="700" height="394"></iframe></p>
	Else
		sHtml = <p dir="ltr" style="text-align: center;"><iframe  allowfullscreen="true"  frameborder="0" src="%sInput%" width="800" height="450"></iframe></p>
} ; SharePoint List => embed in iFrame
Else If RegExMatch(sInput,"/Lists/[^.]*.aspx") & Not InStr(sInput,"DispForm.aspx?ID=") { ; exclude link to item e.g. LL/RAI '
	sHtml = <p><iframe height="500" src=" %sInput%?isFrame=1" width="95_percent"></iframe></p>
	sHtml := StrReplace(sHtml,"_percent","%")



; Embed java script e.g. Gist <script src="https://gist.github.com/tdalon/0924af7a21cc2571a2d25040905fe39d.js"></script>
} Else If RegExMatch(sInput,"^<script .*</script>$") {
	sHtml := sInput
} Else If InStr(sInput,"</script>") { ; TODO Overwritten by previous clause
	sHtml = <html><body>%sInput%</body></html>
	sHtml := uriEncode(sHtml)
	sHtml = <iframe src="data:text/html;charset=utf-8,%sHtml%" frameborder="0" width="100{percent}"/>
	sHtml := StrReplace(sHtml, "{percent}" ,"%")	
}


; Embed iframes e.g. YouTube embed code - center
Else If RegExMatch(sInput,"^<iframe [^>]*</iframe>") {
	sHtml = <p style="text-align: center;">%sInput%</p>
}	
Else If RegExMatch(sInput,"<iframe [^>]* src=""https://www\.youtube\.com/embed/") { ; TODO overwritten by previous clause
	sHtml = <p style="text-align: center;">%sInput%</p>
}
Else If RegExMatch(sInput,"https://www\.youtube\.com/watch\?.*v=([^\?&]*)",sVideoId) { ; bug fix links https://www.youtube.com/watch?time_continue=1&v=y6qT502Ao58 {
	sHtml =	<iframe width="560" height="315" src="https://www.youtube.com/embed/%sVideoId1%" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe>
	sHtml = <p style="text-align: center;">%sHtml%</p>
}
Else If RegExMatch(sInput,"https://youtu\.be/([^\?&]*)",sVideoId) ; https://youtu.be/ngLfEVU46x0?t=1007
{ 
	sHtml =	<iframe width="560" height="315" src="https://www.youtube.com/embed/%sVideoId1%" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe>
	sHtml = <p style="text-align: center;">%sHtml%</p>
}

Else If RegExMatch(sInput,"https://www\.youtube\.com/playlist\?list=([^\?&]*)",sVideoId) ; https://www.youtube.com/playlist?list=PLUSZfg60tAwLIIs8TpcOJIG9ghbQd5nHj
{ 
	sHtml =	<iframe width="560" height="315" src="https://www.youtube.com/embed/videoseries?list=%sVideoId1%" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe>
	sHtml = <p style="text-align: center;">%sHtml%<br><a href="%sInput%">Direct Link to YouTube playlist</a></p>
}

; Stream embed code - center+ add title with direct link	
Else If RegExMatch(sInput,"<iframe [^>]* src=""https://web\.microsoftstream\.com/embed/") {
	sPat = src="([^"]*)"
	RegExMatch(sInput, sPat, sSrc)
	sInput := RegExReplace(sSrc1, "\?.*" , "")
	sInput := StrReplace(sInput,"/embed/","/")	
	sHtml = <p style="text-align: center;">%sInput%<br><a href="%sInput%">Direct Link to Stream video</a></p>
}

; Stream video with start time
Else If RegExMatch(sInput,"https://web\.microsoftstream\.com/video/.*\?st=(.*)",ss)  { 
	linktext := FormatSeconds(ss1)
	sHtml = <a href="%sInput%">%linktext%</a>
} Else If RegExMatch(sInput,"https://web\.microsoftstream\.com/video/([^/]*)",sVideoId) {
	sEmbedCode = <iframe width="640" height="360" src="https://web.microsoftstream.com/embed/video/%sVideoId1%?autoplay=false&amp;showinfo=true" style="border:none;" allowfullscreen ></iframe>
	sHtml = <p style="text-align: center;">%sEmbedCode%<br><a href="%sInput%">Direct Link to Stream video</a></p>
} Else If RegExMatch(sInput,"^https://web\.microsoftstream\.com/channel/([^/]*)",sChannel){
	;MsgBox % sChannel1 ; DBG
	sEmbedCode = <iframe width="960" height="540" src="https://web.microsoftstream.com/embed/channel/%sChannel1%?sort=trending" style="border:none;" allowfullscreen ></iframe>	
	sHtml = <p style="text-align: center;">%sEmbedCode%<br><a href="%sInput%">Direct Link to Stream Channel</a></p>
}


	
If !sHtml {	; empty
	if !linktext ; empty 
		linktext := Link2Text(sInput) 
	if !linktext ; empty => cancel
		return
	sHtml =	<a href="%sInput%">%linktext%</a>
} ; End If sHtml empty

If useico { ; not empty; File Link with FileType icon
	imgsrc := Link2Ico(sInput)
	if imgsrc ; not empty
		sHtml = <a href="%sInput%"><img src="%imgsrc%" width="32"/></a>%sHtml%
}

return sHtml
}


; folder: http://connext.conti.de/files/form/anonymous/api/library/ffd279a0-2764-46bb-a8ec-b1aa3c713072/document/6bf73af1-2152-4797-b383-60339c90f977/media/folder_32.png?logDownload=true&downloadType=view&versionNum=1

; -------------------------------------------------------------------------------------------------------------------
FormatSeconds(NumberOfSeconds)  ; Convert the specified number of seconds to hh:mm:ss format.
; From https://www.autohotkey.com/docs/commands/FormatTime.htm
; https://autohotkey.com/board/topic/44704-converting-time-given-only-seconds/
{
	ElapsedHours := SubStr(0 Floor(NumberOfSeconds / 3600), -1)
  	ElapsedMinutes := SubStr(0 Floor((NumberOfSeconds - ElapsedHours * 3600) / 60), -1)
	ElapsedSeconds := SubStr(0 Floor(NumberOfSeconds - ElapsedHours * 3600 - ElapsedMinutes * 60), -1)	
	If (ElapsedHours = "00") {
		ss = %ElapsedMinutes%:%ElapsedSeconds%
		return ss
	}
	Else
	{
		ss = %ElapsedHours%:%ElapsedMinutes%:%ElapsedSeconds%
		return ss
	}
}

; FUNCTION Link2Ico
Link2Ico(sLink){
; imgsrc := Link2Ico(sLink)
; Calls/include IsSharepointUrl

sLink := RegExReplace(sLink, "\?.*$","") ; remove e.g. ?d=
SplitPath, sLink, sFileName, , sFileExt ; Extension without .
IniFile = PowerTools.ini
If sFileExt {
	IniRead, imgsrc, %IniFile%, ConnectionsIcons, %sFileExt%FileIcon
	If (imgsrc != "ERROR")
		return imgsrc
	If (sFileExt = "pdf")
		IniRead, imgsrc, %IniFile%, ConnectionsIcons, pdfFileIcon
	Else If (sFileExt = "doc") || (sFileExt = "docx") || (sFileExt = "docm")
		IniRead, imgsrc, %IniFile%, ConnectionsIcons, docFileIcon
	Else If (sFileExt = "xlsm") || (sFileExt = "xlsx") || (sFileExt = "xls")
		IniRead, imgsrc, %IniFile%, ConnectionsIcons, xlsFileIcon
	Else If (sFileExt = "pptx") || (sFileExt = "pptm") || (sFileExt = "ppt")
		IniRead, imgsrc, %IniFile%, ConnectionsIcons, pptFileIcon		

	If (imgsrc != "ERROR")
		return imgsrc
}

If (SharePoint_IsSPUrl(sLink))
	IniRead, imgsrc, %IniFile%, ConnectionsIcons, SharepointIcon
Else If InStr(sLink,"communityUuid=") {
	sPat=.*?\?communityUuid=([^ ]*)
	sRep = https://%PowerTools_ConnectionsRootUrl%/communities/service/html/image?communityUuid=$1
	imgsrc:= RegExReplace(sLink,sPat,sRep)
} Else If Connections_IsUrl(sLink,"wiki")
	IniRead, imgsrc, %IniFile%, ConnectionsIcons, WikiIcon
Else If Connections_IsUrl(sLink,"blog")
	IniRead, imgsrc, %IniFile%, ConnectionsIcons, BlogIcon
Else If Connections_IsUrl(sLink,"forum")
	IniRead, imgsrc, %IniFile%, ConnectionsIcons, ForumIcon
Else If InStr(sLink,"/gist/")
	IniRead, imgsrc, %IniFile%, ConnectionsIcons, GistIcon
Else If (InStr(sLink,"github.") and Not (InStr(sLink,"github.io/")))
	IniRead, imgsrc, %IniFile%, ConnectionsIcons, GithubIcon
Else If Jira_IsUrl(sLink)
	IniRead, imgsrc, %IniFile%, ConnectionsIcons, JiraIcon
Else If InStr(sLink,"confluence")
	IniRead, imgsrc, %IniFile%, ConnectionsIcons, ConfluenceIcon
Else If InStr(sLink,"https://teams.microsoft.com/l/channel/") || InStr(sLink,"https://teams.microsoft.com/l/team/") 
	IniRead, imgsrc, %IniFile%, ConnectionsIcons, TeamsIcon
Else If InStr(sLink,"https://web.microsoftstream.com/browse?q=")
	IniRead, imgsrc, %IniFile%, ConnectionsIcons, StreamIcon

If (imgsrc = "ERROR")
	return
Else
	return imgsrc

} ; end function

; -------------------------------------------------------------------------------------------------------------------
CleanUrl(url,decode:=false){
; Clean connext url from lang tag
; Clean-up ugly SharePoint folder url
; Convert Teams file links to SharePoint links
;
; Calls: Lib/GetSharepointUrl, GetFileLink, UriDecode
; Call: SubFunctions: GetGoogleUrl, IsGoogleUrl

url := Trim(url)

; Teams Beautifier
url := Teams_FileLinkBeautifier(url) ; will decode url
If Connections_IsUrl(url) {
	ReConnectionsRootUrl := StrReplace(PowerTools_ConnectionsRootUrl,".","\.")
	
	; Switch to https
	url := StrReplace(url,"http://","https://")
	
	; Link to Blog entries: http://connext.conti.de/blogs/tdalon?tags=chrome&lang=en
	; Wiki page: http://connext.conti.de/wikis/home?lang=en#!/wiki/Wa8a86fe4ac2b_4e9d_8e98_17d4671c70f8 
	; => remove ?lang=en#!
	; Comment permalink http://connext.conti.de/blogs/tdalon/entry/connext_link_format?lang=en#threadid=356630b7-2c83-4d94-b033-d8ca24f456d7
	; => http://connext.conti.de/blogs/tdalon/entry/connext_link_format#threadid=356630b7-2c83-4d94-b033-d8ca24f456d7
	
	; Language might be en-us => add - to the word		; de_de
	
	
	; Wiki section
	;		http://connext.conti.de/wikis/home?lang=en#!/wiki/W10f67125ddc8_42e1_a6da_0a8e6a1cd541/page/Help%20on%20NWS%20Search%20Tool&section=overview
	url := RegExReplace(url, "[&|\?]lang=[\w-_]+(#!)?" , "")
	url := StrReplace(url,"&section=","?section=")
		
	; wiki pages filtered by tag example http://connext.conti.de/wikis/home/wiki/W354104eee9d6_4a63_9c48_32eb87112262/index?lang=en&tag=nws_workflow
	url := StrReplace(url,"index&tag=","index?tag=")
	
	; Remove ?logDownload=true&downloadType=view&versionNum=1 (when copying image from Files)
	; Ex. http://connext.conti.de/files/form/anonymous/api/library/9d7cc0b6-434a-4f44-9091-6f13fbac8a4e/document/f091ded1-da5d-45fb-8910-927439348a90/media/idea_orange_trans.png?versionNum=1
	; Ex. http://connext.conti.de/files/form/anonymous/api/library/ffd279a0-2764-46bb-a8ec-b1aa3c713072/document/87249fa7-87f8-4977-b511-6c49a1597e31/media/wikis_32.jpg?logDownload=true&downloadType=view&versionNum=1

	;url := RegExReplace(url, "\?versionNum=\d+" , "")
	url := RegExReplace(url, "://" . ReConnectionsRootUrl . "/files/form/anonymous/api/library/([^?]*)\?.*", "://" . PowerTools_ConnectionsRootUrl . "/files/form/anonymous/api/library/$1")
	
	; Remove lastMod (when copying profile picture)
	; Ex. http://connext.conti.de/profiles/photo.do?key=7df0fd93-6999-426d-869c-d36d434d11fa&lastMod=1500531017000
	url := RegExReplace(url, "&lastMod=\d+" , "")
	
	; Remove ?preventCache=1500037805880 if copied from wiki/blogs
	url := RegExReplace(url, "\?preventCache=\d+" , "")
	
	; Remove ?logDownload=true&downloadType=view from copy video address from context menu
	url := StrReplace(url, "?logDownload=true&downloadType=view","")
	
	; For FireFox/Chrome support replace https: to http: to make images visible if accessed via http
	;SplitPath, url, sFileName,,sFileExt 
	;If (sFileExt = "png") || (sFileExt = "jpg") || (sFileExt = "gif") 
	;	url := StrReplace(url,"https://","http://")
	
	url := StrReplace(url,"'","%27")

} Else If (SharePoint_IsSPUrl(url)) 
	url := SharePoint_CleanUrl(url)	
Else If InStr(url,".blogspot.") {
	;url := RegExReplace(url,"/\d{4}/\d{2}","") - such shortened links takes longer to load -> remove
	url := StrReplace(url,".html","")
} Else If (IsGoogleUrl(url))
	url := GetGoogleUrl(url)
 
			
If (decode) {
	url := uriDecode(url)	

	; fix beautified url for ConNext wikis
	If (Connections_IsUrl(url,"wiki"))	{
		;url := StrReplace(url,"(","%28")
		;url := StrReplace(url,")","%29")
		url := StrReplace(url,"/#","/%23") ; links to page starting with # won't open
	}	
} Else {
	url := StrReplace(url,"(","%28")
	url := StrReplace(url,")","%29")
	url := StrReplace(url," ","%20")
	url := StrReplace(url,",","%2C")
	
	; Teams conversation link - make it clickable e.g. Worflowy
	If InStr(url,"https://teams.microsoft.com/l/message/")
	{
		url := StrReplace(url,"https://teams.microsoft.com/l/message/","")		
		url := StrReplace(url,":","%3A")
		url := StrReplace(url,"@","%40")
		url = https://teams.microsoft.com/l/message/%url%
	}
	
	;url := StrReplace(url,"""","%22") 
	;url := StrReplace(url,"[","%5B")
	;url := StrReplace(url,"]","%5D")
}	

return url
}

; -------------------------------------------------------------------------------------------------------------------
GetGoogleUrl(url){
; Calls: IsGoogleUrl
	url := uriDecode(url)	
	RegExMatch(url,"&url=(.*)&usg=",newurl) ; exclude & escape ?
	; 1st Matching pattern is stored in RootFolder1
	; MsgBox %RootFolder1%
	; decode url
		
return newurl1

}
; -------------------------------------------------------------------------------------------------------------------

IsGoogleUrl(url){
If RegExMatch(url,"https://www\.google\.[a-z]{2,3}/url\?.*") = 0 
	return false
Else 
	return true
}

; -------------------------------------------------------------------------------------------------------------------
OnIntelliPasteMultiStyleMsgBox() {
    DetectHiddenWindows, On
    Process, Exist
    If (WinExist("ahk_class #32770 ahk_pid " . ErrorLevel)) {
		ControlSetText Button1, lines-break
		ControlSetText Button2, single-line
		ControlSetText Button3, bullet-list
    }
}

; -------------------------------------------------------------------------------------------------------------------
IntelliPaste(){
sClipboard := Clipboard  


If InStr(sClipboard,"`n") { ; MultiLine 
	;MsgBox 0x21, MultiLine, You are pasting Multiple lines!
	sStyle = ?
	If (sStyle = "?") {
		OnMessage(0x44, "OnIntelliPasteMultiStyleMsgBox")
		MsgBox 0x23, MultiLinks Style, Select your style for multi-links:
		OnMessage(0x44, "")
		sStyle = Abort
		IfMsgBox, Yes
			sStyle = lines-break ; Default
		IfMsgBox, No
			sStyle = single-line
		IfMsgBox, Cancel
			sStyle = bullet-list
		If (sStyle = "Abort")
			return
	}
} Else 
	sStyle = single-line ; Default

SetTitleMatchMode, 2 ; partial match

If (Connections_IsWinActive())
	useico := True
Else
	useico := False ; Deactivate till it is refactored. #TODO
; Do not ask for ico for specific applications where pasting icon doesn't work
;If WinActive("ahk_group MSOffice") or  WinActive("Confluence") or WinActive("Microsoft Teams")  or WinActive("Microsoft Stream") or WinActive("Microsoft Whiteboard") or  WinActive("ahk_exe notepad.exe") or WinActive("ahk_exe notepad++.exe") or WinActive("ahk_exe WorkFlowy.exe") or WinActive("ahk_exe atom.exe") ; plain text editor
;	useico := False
; If only one link, set useico to False in case of: video/ image files...
/*
	If !InStr(sClipboard,"`n") { ; single entry
		If RegExMatch(sClipboard,"https://web.microsoftstream.com/browse\?")
			useico := True
		Else If InStr(sClipboard,".mp4") or InStr(sClipboard,"_player.html") or RegExMatch(sClipboard,"\.png$") or RegExMatch(sClipboard,"\.gif$") or RegExMatch(sClipboard,"\.jpg$") or  RegExMatch(sClipboard,"(SD[\d]{1,})") or RegExMatch(sClipboard,"https://www.youtube.com/") or RegExMatch(sClipboard,"https://web.microsoftstream.com/")
			useico := False
	}
*/

If !InStr(sClipboard,"`n") and InStr(sClipboard,"https://github.com/") {
	sFile := StrReplace(sClipboard,"/blob/","/raw/")
	sFullText = <script src="http://gist-it.appspot.com/%sFile%"></script>
	Clip_Paste(sFullText)
	return
}

; Check if one ico is necessary (slower but less user annoying: do not ask if not needed)
If (useico) {
	useico := False
	Loop, parse, sClipboard, `n, `r
	{
		sLink := CleanUrl(A_LoopField) ; TODO call again below
		imgsrc := Link2Ico(sLink)
		;MsgBox %sLink% : %imgsrc% ; DBG
		If not (imgsrc="") { ; icon not empty
			useico := True
			break ; return true
		}
	}
}

WinGet WinId, ID, A
If (useico) {
	MsgBox, 3,IntelliPaste: Question, Would you like to insert icons before links?	
	IfMsgBox Yes
		useico := True
	IfMsgBox No
		useico := False
	IfMsgBox Cancel
		return
}

WinActivate, ahk_id %WinId% ; Bug with Chrome: window looses focus (cursor deactivated)

If  WinActive("ahk_exe notepad.exe") or WinActive("ahk_exe notepad++.exe") { ;-> Plain-text
	Fmt = Text
} Else {
	Fmt = HTML
}

sFullHtml =
sFullText =
; IntelliPaste Links

If !InStr(sClipboard,"`n") { ; single input
	sLink := sClipboard
	If RegExMatch(sClipboard,"^http") ; only for links - exclude script, html code	
		sLink := CleanUrl(sLink)
	
	sHtml := IntelliHtml(sLink, useico)
	If (sLink = sClipboard) { ; no action on clean -> transform to html
		Clip_PasteHtml(sHtml,sHtml)
	} Else {
		Clip_PasteHtml(sHtml,sLink)
	}
	return
}

Loop, parse, sClipboard, `n, `r
{
	sLink_i := CleanUrl(A_LoopField)	; calls also GetSharepointUrl

	Switch Fmt 
	{
	Case "Text":
		If (sStyle = "single-line")
			sFullText := sFullText . sLink_i . " "
		Else If (sStyle = "lines-break")
			sFullText := sFullText . sLink_i . "`n"
		Else If (sStyle = "bullet-list")
			sFullText := sFullText . "- " . sLink_i . "`n"
	Case "HTML":
		sHtml := IntelliHtml(sLink_i, useico)
		if !sHtml ; empty => cancel
			return
		If (sStyle = "single-line")
			sFullHtml := sFullHtml . sHtml . " "
		Else If (sStyle = "lines-break")
			sFullHtml := sFullHtml . sHtml . "<br>"
		Else If (sStyle = "bullet-list")
			sFullHtml := sFullHtml . "<li>" . sHtml . "</li>"
		sFullText := sFullText . sLink_i . " "
	}
}	; End Loop Parse Clipboard

If (Fmt = "Text") {
	Clip_Paste(sFullText)
} Else If (Fmt = "HTML") {
	If (sStyle = "bullet-list")
		sFullHtml :=  "<ul>" . sFullHtml . "</ul>"
	Clip_PasteHtml(sFullHtml,sFullText)
}


} ; eofun IntelliPaste()

; -------------------------------------------------------------------------------------------------------------------

IntelliPaste_HotkeySet(){
; https://autohotkey.com/board/topic/47439-user-defined-dynamic-hotkeys/
If GetKeyState("Ctrl")  { ; exclude ctrl if use in the hotkey
	IntelliPaste_Help()
	return
}

RegRead, IntelliPasteHotkey, HKEY_CURRENT_USER\Software\PowerTools, IntelliPasteHotkey
HK := HotkeyGUI(,IntelliPasteHotkey,,,"IntelliPaste - Set Hotkey")
If ErrorLevel ; Cancelled
  return
If (HK = IntelliPasteHotkey) ; no change
  return
PowerTools_RegWrite("IntelliPasteHotkey",HK)

; Turn off the old Hotkey
If Not (IntelliPasteHotkey == "")
	Hotkey, %IntelliPasteHotkey%, IntelliPaste, Off

;Turn on the new hotkey.
If (HK == "Insert") {
	Hotkey, IfWinNotActive, ahk_group NoIntelliPasteIns
	Hotkey, %HK%, IntelliPaste, On 
	Hotkey, IfWinNotActive,
} Else
	Hotkey, %HK%, IntelliPaste, On 

TrayTip, Set IntelliPaste Hotkey,% HK " Hotkey on"

} ; eofun

; -------------------------------------------------------------------------------------------------------------------

IntelliPaste_Help(){
PowerTools_OpenDoc("intelli_paste")
} ; eofun

; -------------------------------------------------------------------------------------------------------------------

IntelliPaste_Refresh(){
; Refresh SPSync.ini 
SharePoint_UpdateSyncIniFile()
; Refresh Teams Export List
Teams_ExportTeams()
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
GetFileLink(sFile) {
; Function GetFileLink
; Calls: Subfunction: DriveMap_Get, SharePoint_Sync2Url
; Called by: IntelliPaste
	If InStr(sFile,"http") {
		Return sFile
	}
	
	sFile := StrReplace(sFile,"%20"," ") ; replace %20 by spaces etc (uri decode does not work on file path)
   	   
	sUrl := SharePoint_Sync2Url(sFile)
	If !(sUrl = "")
		return sUrl

	StringLeft sDrive, sFile, 2
	sShare := DriveMap_Get(sDrive)
	If sShare <> 
	{
		sFile := StrReplace(sFile, sDrive ,sShare)
	}

	sFile2 := RegExReplace(sFile,"i)@SSL\\DavWWWRoot","") ; case-insensitive
	If !(sFile2 = sFile)
	{
		sFile := StrReplace(sFile2,"\","/") 
		sFile = https:%sFile%
	}
	Else {
		sFile2 := StrReplace(sFile,"\","/") 
		sFile2 = https:%sFile2%
		If SharePoint_IsSPUrl(sFile2)
			return sFile2
	}
	;Else
	;	sFile = file://%sFile%
	return sFile
}


; -------------------------------------------------------------------------------------------------------------------
; Function:       DriveMap_Get      -  retrieves the name of the network share associated with the local drive.
; Parameter:      Drive    -  drive letter to get the network name for followed by a colon (e.g. "Z:").
; Return Values:  The name of the share on success, otherwise an empty string.
;                 ErrorLevel contains the system error code, if any.
; MSDN:           WNetGetConnection() -> http://msdn.microsoft.com/en-us/library/aa385453%28v=vs.85%29.aspx
; -------------------------------------------------------------------------------------------------------------------
; https://autohotkey.com/boards/viewtopic.php?f=6&t=501
DriveMap_Get(Drive) {
	Static Length := 512
	Static MinDL := "E"  ; minimum drive letter
	Static MaxDL := "Z"  ; maximum drive letter
	ErrorLevel := 0
	If !RegExMatch(Drive, "i)^[" . MinDL . "-" . MaxDL . "]:$") ; invalid drive
		 Return ""
	VarSetCapacity(Share, Length * 2, 0)
	If (Result := DllCall("Mpr.dll\WNetGetConnection", "Str", Drive, "Str", Share, "UIntP", Length, "UInt")) {
		 ErrorLevel := Result
		 MsgBox Error DriveMap_Get "%Result%"
		 Return ""
	}
	Return Share
}
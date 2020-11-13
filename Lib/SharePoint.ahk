; Library for SharePoint Utilities

; -------------------------------------------------------------------------------------------------------------------

; Sharepoint_CleanUrl - Get clean sharepoint document library Url from browser url
; Syntax:
;	newurl := Sharepoint_CleanUrl(url)
; Example:
; 	https://continental.sharepoint.com/teams/team_10000035/Guides%20-%20Documents/Forms/AllItems.aspx?cid=e18e743c%2D86a5%2D4eb2%2D9e66%2Dded0fa83986f&RootFolder=%2Fteams%2Fteam%5F10000035%2FGuides%20%2D%20Documents%2FGeneral&FolderCTID=0x012000A46ECE04B4C0CD4D963D7DF0C1F91CBA
; in IE https://continental.sharepoint.com/teams/team_10000035/Guides%20-%20Documents/Forms/AllItems.aspx?cid=e18e743c%2D86a5%2D4eb2%2D9e66%2Dded0fa83986f&id=%2Fteams%2Fteam%5F10000035%2FGuides%20%2D%20Documents%2FGeneral
; =>
; https://continental.sharepoint.com/teams/team_10000035/Guides - Documents/General
; https://continental.sharepoint.com/teams/team_10000035/Guides%20-%20Documents/General
;
#Include <UriDecode>
#Include <People>
; for People_GetMyOUid Personal OneDrive url

#Include <Teams>
; for Teams Name in IntelliPaste
; Calls: uriDecode
; Called by: CleanUrl
SharePoint_CleanUrl(url){
	If InStr(url,"_vti_history") ; speciasl case hardlink for old sharepoints
	{
		url := uriDecode(url)
		RegExMatch(url,"\?url=(.*)",newurl)
		return newurl1
	}

	; For new spo links
	If RegExMatch(url ,":[a-z]:/r/") { ; New SPO links
		url := RegExReplace(url ,":[a-z]:/r/","")
		; keep durable link
		If RegExMatch(url,"\?d=")
			url := RegExReplace(url ,"&.*","")
		Else
			url := RegExReplace(url ,"\?.*","")
		return url
	}

	url := StrReplace(url,"?Web=1","") ; old sharepoint link opening in browser


	; For old SP links
	RegExMatch(url,"https://([^/]*)",rooturl) 
	rooturl = https://%rooturl1%

	If !RegExMatch(url,"(?:\?|&)RootFolder=([^&]*)",RootFolder) 
		If !RegExMatch(url,"(?:\?|&)id=([^&]*)",RootFolder) 
			return RegExReplace(url,"/Forms/AllItems.aspx.*$","")
		
	; exclude & 
	; non-capturing group starts with ?: see https://autohotkey.com/docs/misc/RegEx-QuickRef.htm
		
	; decode url			
	RootFolder:= uriDecode(RootFolder1)
	;msgbox %RootFolder%	
	
	newurl := rooturl . RootFolder
	;MsgBox %newurl%
	return newurl	
}


; TEST old SharePoint
;   https://cws7.conti.de/content/11011979/CFST_Documents/Forms/AllItems.aspx?RootFolder=%2fcontent%2f11011979%2fCFST%5fDocuments%2fTest%5fArchitecture%5fand%5fLibrary%2fStart%5fSmart%5fPackage&FolderCTID=0x012000DB1CEDE45CDD754299DFD96C87BD058F
; TEST Teams
; https://continental.sharepoint.com/teams/team_10000778/Shared%20Documents/Forms/AllItems.aspx?FolderCTID=0x012000761015BBA31DEE4AB964DB6D73C115C2&id=%2Fteams%2Fteam%5F10000778%2FShared%20Documents%2FTools
; TEST Teams with Copy Link
; https://continental.sharepoint.com/:f:/r/teams/team_10000035/Shared%20Documents/General/Pictures?csf=1&e=Um9ecD

; -------------------------------------------------------------------------------------------------------------------

; SharePoint_IsSPUrl(url)
SharePoint_IsSPUrl(url){
If RegExMatch(url,"https://[a-z\-]+\.sharepoint\.com/.*") or InStr(url,"https://mspe.conti.de/") or RegExMatch(url,"https://[a-z]+\d\.conti\.de/.*") 
{
	; workspace1.conti.de cws7.conti.de = url with a few letters followed by one number
    ;  or InStr(url,"https://continental.sharepoint.com/") or InStr(url,"https://continental-my.sharepoint.com/") 
	return true
	}
Else {
	return false
	}
}

; -------------------------------------------------------------------------------------------------------------------
; Called by: IntelliPaste -> IntelliHtml
SharePoint_IntelliHtml(sLink){

If RegExMatch(sLink,"/teams/team_[^/]*/[^/]*/(.*)",sMatch) {
	sMatch1 := uriDecode(sMatch1)
	linktext := StrReplace(sMatch1,"/"," > ") ; Breadcrumb navigation for Teams link to folder
	 
	TeamName := Teams_GetTeamName(sLink)
	If (TeamName != "") ; not empty
		linktext := TeamName . " > " . linktext
	
}


} ; eofun
; -------------------------------------------------------------------------------------------------------------------

GetRootUrl(sUrl){
    RegExReplace(sUrl,"https?://[^/]",rootUrl)
    return rootUrl
}

; -------------------------------------------------------------------------------------------------------------------

SharePoint_GetSyncIniFile(){
    EnvGet, sOneDriveDir , onedrive
	sOneDriveDir := StrReplace(sOneDriveDir,"OneDrive - ","")
	sIniFile = %sOneDriveDir%\SPsync.ini
    return sIniFile
}

; -------------------------------------------------------------------------------------------------------------------
SharePoint_Url2Sync(sUrl){
; Path := SharePoint_Url2Sync(sUrl)
; returns empty if not sync'ed
	
}


; -------------------------------------------------------------------------------------------------------------------
SharePoint_Sync2Url(sFile){
; sUrl := SharePoint_Sync2Url(sFile)
; returns empty if not sync'ed

; Get File Link for Personal OneDrive
EnvGet, sOneDriveDir , onedrive

If InStr(sFile,sOneDriveDir . "\") { 
	MyOUid := People_GetMyOUid()
	rootUrl = https://continental-my.sharepoint.com/personal/%MyOUid%_contiwan_com/Documents
	; TODO replace username by oUid
	sFile := StrReplace(sFile, sOneDriveDir,rootUrl)
	sFile := StrReplace(sFile, "\", "/")
	return sFile
}

; Get File Link for SharePoint/OneDrive Synced location
sOneDriveDir := StrReplace(sOneDriveDir,"OneDrive - ","")
needle :=  StrReplace(sOneDriveDir,"\","\\") ; 
needle := needle "\\[^\\]*"
If Not (RegExMatch(sFile,needle,syncDir))
	Return

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

Loop, Read, %sIniFile%
{
	Array := StrSplit(A_LoopReadLine, A_Tab," `t",2)
	If !Array
		continue
	rootDir := StrReplace(Array[1],"??",".*")
	rootDirRe := StrReplace(rootDir,"\","\\") ; escape filesep
	If (RegExMatch(syncDir, rootDirRe)) {
		rootUrl := Array[2]
		If rootUrl = "#TBD"
			break
		sFile := StrReplace(sFile, syncDir,rootUrl)
		sFile := StrReplace(sFile, "\", "/")
		return sFile
	}
}	; End Loop		

If (!rootUrl) { ; empty
	FileAppend, %syncDir%%A_Tab%#TBD`n, %sIniFile%			
}
TrayTip, NWS PowerTool, File SPsync.ini is not properly filled for %syncDir%! Fill it following user documentation.,,0x23
sCmd = Edit "%sIniFile%"
Run %sCmd%
;Run https://connext.conti.de/blogs/tdalon/entry/onedrive_sync_ahk
return
} ; eofun
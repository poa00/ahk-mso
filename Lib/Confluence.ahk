#Include <WinActiveBrowser>

Confluence_IsUrl(sUrl){
If RegExMatch(sUrl,"confluence[\.-]")
	return True
}
; -------------------------------------------------------------------------------------------------------------------

Confluence_IsWinActive(){
If Not WinActiveBrowser()
    return False
sUrl := GetActiveBrowserUrl()
return Confluence_IsUrl(sUrl)
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Confluence_ExpandLinks(sLinks){
; Called by IntelliPaste
	Loop, parse, sLinks, `n, `r
	{
		sLink_i := CleanUrl(A_LoopField)	; calls also GetSharepointUrl
		ExpandLink(sLink_i)
	}
}

; -------------------------------------------------------------------------------------------------------------------
; Confluence Expand Link
ExpandLink(sLink){
	sLinkText := Link2Text(sLink)
	MoveBack := StrLen(sLinkText)
	
	SendInput {Raw} %sLinkText%
	SendInput {Shift down}{Left %MoveBack%}{Shift up}
	SendInput, ^k
	sleep, 2000 
	SendInput {Raw} %sLink% 
	SendInput {Enter}
	return
}

; Confluence Search - Search within current Confluence Space
; Called by: NWS.ahk (Win+F Hotkey)
Confluence_Search(sUrl){
static sConfluenceSearch, sKey
; http://confluence.conti.de:8090/dosearchsite.action?cql=siteSearch+~+"project+status+report"+and+space+=+"projectCFTPT"+and+type+=+"page"

; Extract project key from Url and confluence root url
; http://confluence.conti.de:8090/spaces/viewspace.action?key=projectCFTI
; or http://confluence.conti.de:8090/display/projectCFTI/Newsletters

RegExMatch(sUrl,"https?://[^/]*",sRootUrl)
If RegExMatch(sUrl,sRootUrl . "/display/([^/]*)",sKey)
    sKey := sKey1
Else If  RegExMatch(sUrl,sRootUrl . "/spaces/viewspace.action\?key=([^&\?/]*)",sKey)
    sKey := sKey1
Else
    Return
sOldKey := sKey
If sKey = %sOldKey%
	sDefSearch := sConfluenceSearch
Else
	sDefSearch = 

InputBox, sSearch , Confluence Search, Enter search string:,,640,125,,,,,%sDefSearch% 
if ErrorLevel
	return
sSearch := Trim(sSearch) 
sConfluenceSearch := StrReplace(sSearch," ","+")
sSearchUrl = /dosearchsite.action?cql=siteSearch+~+"%sConfluenceSearch%"+and+space+=+"%sKey%"+and+type+=+"page"
Run, %sRootUrl%%sSearchUrl%

}
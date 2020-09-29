; Jira Lib
; Includes JiraSearch, IsJiraUrl, JiraAuth
; for GetPassword
#Include <Login> 
; for JiraFormatLinks: CleanUrl, Link2Text
#Include <IntelliPaste>

; ----------------------------------------------------------------------
JiraGet(sUrl){
; Syntax: sResponse .= JiraGet(sUrl)
; Calls: b64Encode

sPassword := Login_GetPassword()
If (sPassword="") ; cancel
    return	

WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
WebRequest.Open("GET", sUrl, false) ; Async=false

 ; https://developer.atlassian.com/cloud/jira/platform/basic-auth-for-rest-apis/
sAuth := b64Encode(A_UserName . ":" . sPassword) ; user:password in base64 
WebRequest.setRequestHeader("Authorization", "Basic " . sAuth) 
WebRequest.setRequestHeader("Content-Type", "application/json")

WebRequest.Send()        
sResponse := WebRequest.responseText
return sResponse
}

; ----------------------------------------------------------------------
Jira_IsUrl(sUrl){
return (InStr(sUrl,"http://it-ea-agile.conti.de:7090")) || (InStr(sUrl,"jira"))
}

; ----------------------------------------------------------------------
; Jira Search - Search within current Jira Project
; Called by: NWS.ahk Quick Search (Win+F Hotkey)
JiraSearch(sUrl){
static sJiraSearch, sKey

; http://confluence.conti.de:8090/dosearchsite.action?cql=siteSearch+~+"project+status+report"+and+space+=+"projectCFTPT"+and+type+=+"page"

; Extract project key from Url and confluence root url
; http://confluence.conti.de:8090/spaces/viewspace.action?key=projectCFTI
; or http://confluence.conti.de:8090/display/projectCFTI/Newsletters

If RegExMatch(sUrl,sRootUrl . "/display/([^/]*)",sKey)
    sKey := sKey1
Else If  RegExMatch(sUrl,sRootUrl . "/spaces/viewspace.action?key=([^&?/]*)",sKey)
    sKey := sKey1
Else
    Return
sOldKey := sKey
If sKey = %sOldKey%
	sDefSearch := sConfluenceSearch
Else
	sDefSearch = 

InputBox, sSearch , Search string, Enter search string (use # for tags):,,640,125,,,,,%sDefSearch% 
if ErrorLevel
	return
sSearch := Trim(sSearch) 
sConfluenceSearch := StrReplace(sSearch," ","+")
sSearchUrl = /dosearchsite.action?cql=siteSearch+~+"%sConfluenceSearch%"+and+space+=+"%sKey%"+and+type+=+"page""
Run, %sSearchUrl%

}
; ----------------------------------------------------------------------


b64Encode(string)
; ref: https://github.com/jNizM/AHK_Scripts/blob/master/src/encoding_decoding/base64.ahk
{
    VarSetCapacity(bin, StrPut(string, "UTF-8")) && len := StrPut(string, &bin, "UTF-8") - 1 
    if !(DllCall("crypt32\CryptBinaryToString", "ptr", &bin, "uint", len, "uint", 0x1, "ptr", 0, "uint*", size))
        throw Exception("CryptBinaryToString failed", -1)
    VarSetCapacity(buf, size << 1, 0)
    if !(DllCall("crypt32\CryptBinaryToString", "ptr", &bin, "uint", len, "uint", 0x1, "ptr", &buf, "uint*", size))
        throw Exception("CryptBinaryToString failed", -1)
    return StrGet(&buf)
}



JiraFormatLinks(sLinks,sStyle){
If Not InStr(sLinks,"`n") { ; single line
    sLink := CleanUrl(sLinks)	; calls also GetSharepointUrl
    sLink := StrReplace(sLink,"%20%"," ")
    linktext := Link2Text(sLink)
    sLink = [%linktext%|%sLink%]  
    return sLink
}

Loop, parse, sLinks, `n, `r
{
    sLink := CleanUrl(A_LoopField)	; calls also GetSharepointUrl
    sLink := StrReplace(sLink,"%20%"," ")
    linktext := Link2Text(sLink)
    sLink = [%linktext%|%sLink%]

    If (sStyle = "single-line")
        sFull := sFull . sLink . " "
    Else If (sStyle = "lines-break")
        sFull := sFull . sLink . "`n"
    Else If (sStyle = "bullet-list")
        sFull := sFull . "* " . sLink . "`n"
}
return sFull
}
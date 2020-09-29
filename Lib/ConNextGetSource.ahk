; Get ConNext Html source from Url (not via Edit)
; Syntax:
; 	sSource := ConNextGetSource(sUrl)
; Calls: DonwloadToString for api| ConNextGet (ConNextAuth must be successful)
ConNextGetSource(sUrl){


If InStr(sUrl,"/api/")
    return sSource := DownloadToString(sUrl)

If InStr(sUrl,"/feed/") 
    return CNGet(sUrl)

If InStr(sUrl,"connext.conti.de/wikis/") {
	sSource := MultiLineInputBox("Get Html code from Editor or Inspection. Paste Html Source to be parsed:", "","ConNext Enhancer" )
    return sSource
}
Else If InStr(sUrl,"connext.conti.de/blogs/") {

}
	
Else If InStr(sUrl, "connext.conti.de/forums/html/topic?") or InStr(sUrl,"connext.conti.de/forums/html/threadTopic?") {

}

sSource := CNGet(sUrl) ; Does not work for Wiki/ Forum

return sSource
}
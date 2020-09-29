; Get Blog Entry Id for further API call
; Syntax:
; 	sEntryId := ConNextGetBlogEntryId(sUrl)
; Calls: DownloadToString
ConNextGetBlogEntryId(sUrl){
sSource := DownloadToString(sUrl)
; Parse response for string between <title> </title> for atom <title type="html">
sPat = <input type='hidden' name='entryId' value='([^/]*)'/>
sPat := StrReplace(sPat,"'","""")
If RegExMatch(sSource, sPat, sEntryId)
	return sEntryId1

If !sEntryId
	MsgBox, 16, Error, Can not get blog entryId.
return sEntryId
}


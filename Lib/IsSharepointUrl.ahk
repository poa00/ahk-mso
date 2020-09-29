; -------------------------------------------------------------------------------------------------------------------
; IsSharepointUrl(url)
IsSharepointUrl(url){
If InStr(url,"https://mspe.conti.de/") or RegExMatch(url,"https://[a-z]+\d\.conti\.de/.*")  or InStr(url,"https://continental.sharepoint.com/") or InStr(url,"https://continental-my.sharepoint.com/") 
{
	; workspace1.conti.de cws7.conti.de = url with a few letters followed by one number
	return true
	}
Else {
	return false
	}
}
; Called by IntelliPaste-> Link2Text -> ConNextGetForumTitle, CreateTicket
; sTitle := GetUrlTitle(sUrl, sPassword*)
; If password is passed as argument, authentification is sent with username/password
GetUrlTitle(sUrl,sPassword :="", sUserName:=""){

WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
WebRequest.Open("GET", sUrl, false) ; Async=false
If sPassword { ; not empty
    If !sUserName
        sUserName := A_UserName
	WebRequest.SetCredentials(sUserName, sPassword, 0)
    WebRequest.SetCredentials(sUserName, sPassword, 1)
}
WebRequest.Send()
; Debug
; MsgBox % WebRequest.ResponseText


; Parse response for string between <title> </title> for atom <title type="html">
sPat = <title>([^<]*?)<\/title>
If RegExMatch(WebRequest.ResponseText, sPat, sTitle){
	sTitle1 := StrReplace(sTitle1,"&amp;","&")
	sTitle1 := StrReplace(sTitle1,"#39","'")
	return sTitle1
}
}
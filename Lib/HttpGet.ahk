; Called by IntelliHtml-> Link2Text -> ConNextGetForumTitle, CreateTicket
; sResponseText := HttpGet(sUrl, sPassword*,sUserName*)
; If password is passed as argument, authentification is sent with username/password
; Example: 
; sPassword = Login_GetPassword() - uses password stored under profile Password.txt
; sResponse := Hpptget(sUrl, sPassword) - uses default username environment variable
HttpGet(sUrl,sPassword :="", sUserName:=""){
WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
WebRequest.Open("GET", sUrl, false) ; Async=false
If !(sPassword = "") { ; not empty
    If !(UserName = "")
        sUserName := A_UserName	
    ;MsgBox %sUserName%: %sPassword%
    WebRequest.SetCredentials(sUserName, sPassword, 0)
    WebRequest.SetCredentials(sUserName, sPassword, 1)
}

; White List - no proxy
IsVPN := Not (A_IPAddress2 = "0.0.0.0") ; HttpGet UnAuthorized via VPN
If IsVPN {
    WebRequest.SetProxy(1) ; direct connection
}

WebRequest.Send()
sResponse := WebRequest.ResponseText
; Debug
sStatusText := WebRequest.StatusText
MsgBox %sStatusText% - %sResponse%
If !(sStatusText = "OK"){
    MsgBox 0x10, Error, Error on WinHttpRequest - GET: %sStatusText%
    return
}

sSource := WebRequest.ResponseText
return sSource
}


URLDownloadToVar(URL){
	http:=ComObjCreate("WinHttp.WinHttpRequest.5.1")
	http.Open("GET",URL,1)
	http.SetRequestHeader("Pragma","no-cache")
	http.SetRequestHeader("Cache-Control","no-cache")
	http.Send(),http.WaitForResponse
	return (http.Status=200?http.ResponseText:"Error")
}
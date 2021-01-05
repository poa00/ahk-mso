; Login Lib
#Include <PowerTools> 
; for password from registry

; #Include <Encrypt> 
; For speed reason encryption is not used. Password is saved in HCU registry key - only accessible by current user

; ----------------------------------------------------------------------
Login_GetPassword(){
RegRead, sPassword, HKEY_CURRENT_USER\Software\PowerTools, Password
If (sPassword=""){
    sPassword := Login_SetPassword()
    If (sPassword="") ; cancel
        return
}
; Decrypt(sPassword,"thierry")
return sPassword
}
; ----------------------------------------------------------------------
Login_SetPassword(){
If GetKeyState("Ctrl") {
	sUrl := "https://connext.conti.de/wikis/home/wiki/Wc4f94c47297c_42c8_878f_525fd907cb68/page/Password%20Setting"
	Run, "%sUrl%"
	return
}
InputBox, sPassword, Password, Enter Password for connection, Hide, 200, 125
If ErrorLevel
    return
;sKey := Encrypt(sPassword,"thierry")
PowerTools_RegWrite("Password",sPassword)
; cmdkey /add:windows /user:%A_UserName% /pass:%sPassword%
return sPassword
}

; ----------------------------------------------------------------------
; ----------------------------------------------------------------------
Login_GetPersonalNumber(){
RegRead, sPersNum, HKEY_CURRENT_USER\Software\PowerTools, PersonalNumber
If (sPersNum=""){
    sPersNum := Login_SetPersonalNumber()
    If (sPersNum="") ; cancel
        return
}
return sPersNum
}
; ----------------------------------------------------------------------
Login_SetPersonalNumber(){
If GetKeyState("Ctrl") {
	sUrl := "https://connext.conti.de/wikis/home/wiki/Wc4f94c47297c_42c8_878f_525fd907cb68/page/Password%20Setting"
	Run, "%sUrl%"
	return
}
InputBox, sPersNum, Personal Number, Enter your Personal Number e.g. for KSSE,, 200, 125
If ErrorLevel
    return
PowerTools_RegWrite("PersonalNumber",sPersNum)
return sPersNum
}
; ----------------------------------------------------------------------

Login_IsVPN(){
    return Not (A_IPAddress2 = "0.0.0.0") 
}
; ----------------------------------------------------------------------

VPNConnect(doWait:=False){
SetTitleMatchMode, 1 ; start with
MainTitle = Cisco AnyConnect Secure
LoginTitle = Cisco AnyConnect |
hWnd := WinExist(LoginTitle)
If (hWnd <> 0)  { ; in case it was already opened
    WinActivate, ahk_id %hWnd%  
    GoTo InputVPNPassword
}
hWnd := WinExist(MainTitle)
If (hWnd <> 0) {
    WinActivate, ahk_id %hWnd% 
} Else {
    VPNexePath := "C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe"
    Run, %VPNexePath%
    WinWaitActive, %MainTitle%
}

SetTitleMatchMode, 3 ; exact match for controlclick Connect - else ErrorLevel is 0 with Disconnect button
ControlClick, Connect
If ErrorLevel ; Button not found - Disconnect
    return

SetTitleMatchMode, 1 ; start with
WinWaitActive,Cisco AnyConnect |

InputVPNPassword:
sPassword := Login_GetPassword()
If (sPassword="")
	return

ControlSetText,Edit2,%sPassword%

Send,{Enter}
If (doWait=True)
    WinWaitClose, Cisco AnyConnect Secure
} ; eofun
; ----------------------------------------------------------------------
IsConnectedToInternet()
; source: https://jacksautohotkeyblog.wordpress.com/2018/04/26/checking-your-internet-connection-plus-a-twist-on-a-secret-windows-feature-autohotkey-quick-tips/
{
  IsConnected := DllCall("Wininet.dll\InternetGetConnectedState", "Str", "0x40","Int",0)
  return (IsConnected = 0)
}
; ----------------------------------------------------------------------
Login_IsContiNet(){
sTmpFile = %A_Temp%\ipconfig.txt
If FileExist(sTmpFile)
    FileDelete, %sTmpFile%
RunWait, %ComSpec% /c "ipconfig >"%sTmpFile%"",,Hide
FileRead, Output, %sTmpFile%
If InStr(Output,"conti.de") or InStr(Output,"contiwan.com") ; with VPN shows conti.de / direct connection shows contiwan.com'
    return True
Else
    return False
}

; ########################################################################################################################################
; Obsolete
GetPasswordInFile()
{
    EnvGet, sProfileDir , userprofile
    sFile = %sProfileDir%\password.txt
    
    If Not FileExist(sFile)
		{
			MsgBox, File %sFile% does not exist! Template was copied to "%sProfileDir%". Fill it following user documentation.
			FileCopy, %A_ScriptDir%\password.txt, %sProfileDir%
			Run https://connext.conti.de/blogs/tdalon/entry/Share_nice_ConNext_links
			Run, notepad.exe %sFile%
    		return
		}
    Try {
        FileRead, sPassword, %sFile%
        return sPassword
    } catch e {
        MsgBox, 16, Error, Error in GetPassword.
        return
    }
}
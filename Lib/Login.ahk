; Login Lib
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

IsVPN(){
    return Not (A_IPAddress2 = "0.0.0.0") ; HttpGet UnAuthorized via VPN
}
; ----------------------------------------------------------------------

VPNConnect(){
VPNexePath := "C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe"
Run, %VPNexePath%
If IsVPN()
    return

WinWaitActive,ahk_exe vpnui.exe
Sleep 500 ; workaround some connect issues
ControlClick, Connect
WinWaitActive,Cisco AnyConnect |

sPassword := Login_GetPassword()
If (sPassword="")
	return

ControlSetText,Edit2,%sPassword%
Send,{Enter}

WinWaitClose, ahk_exe vpnui.exe
}
; ----------------------------------------------------------------------


IsConnectedToInternet()
; source: https://jacksautohotkeyblog.wordpress.com/2018/04/26/checking-your-internet-connection-plus-a-twist-on-a-secret-windows-feature-autohotkey-quick-tips/
{
  IsConnected := DllCall("Wininet.dll\InternetGetConnectedState", "Str", "0x40","Int",0)
  return (IsConnected = 0)
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
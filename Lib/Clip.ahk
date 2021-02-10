; -------------------------------------------------------------------------------------------------------------------
Clip_Paste(sText,restore := True) {
; Syntax: Clip_PasteHtml(sHtml,sText,restore := True)
; If sHtml is a link (starts with http), the Html link will be wrapped around it i.e. sHtml = <a href="%sHtml%">%sText%</a>
If (restore)
    ClipBackup:= ClipboardAll
Clipboard := sText
; Paste
SendInput ^v
Sleep, 150
while DllCall("user32\GetOpenClipboardWindow", "Ptr")
    Sleep, 150
If (restore)
    Clipboard:= ClipBackup
} ; eofun
; -------------------------------------------------------------------------------------------------------------------
Clip_Set(sText){
Clipboard := sText
Sleep, 150
while DllCall("user32\GetOpenClipboardWindow", "Ptr")
    Sleep, 150
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Clip_SetHtml(sHtml,sText:="",HtmlHead := ""){
; Syntax: Clip_SetHtml(sHtml,sText:="",HtmlHead := "")
; If sHtml is a link (starts with http), the Html link will be wrapped around it sHtml = <a href="%sHtml%">%sText%</a>
; If no Text display is passed as argument sText := sHtml
If (sText="")
    sText := sHtml
If RegExMatch(sHtml,"^http")
    sHtml = <a href="%sHtml%">%sText%</a>
SetClipboardHTML(sHtml,HtmlHead,sText)
Sleep, 150
while DllCall("user32\GetOpenClipboardWindow", "Ptr")
    Sleep, 150
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Clip_PasteHtml(sHtml,sText:="",restore := True) {
; Syntax: Clip_PasteHtml(sHtml,sText,restore := True)
; If sHtml is a link (starts with http), the Html link will be wrapped around it i.e. sHtml = <a href="%sHtml%">%sText%</a>
If (restore)
    ClipBackup:= ClipboardAll
Clip_SetHtml(sHtml,sText)
; Paste
SendInput ^v
Sleep, 150
while DllCall("user32\GetOpenClipboardWindow", "Ptr")
    Sleep, 150
If (restore)
    Clipboard:= ClipBackup
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Clip_GetSel(type:="text", doTrim := True){
; Syntax:
; 	sSelection := GetSelection("text"*|"html",restore:= True)
; Default "text"
; Output is trimmed
OldClipboard:= ClipboardAll                         ;Save existing clipboard.

Clipboard=
while(Clipboard){
  Sleep,10
}
SendInput,^c    
Sleep, 150
while DllCall("user32\GetOpenClipboardWindow", "Ptr")
    Sleep, 150                                      ;Copy selected text to clipboard

If (type = "text") {
  sSelection := clipboard
} Else If (type ="html") {
   Clip_GetHtml(sSelection)
}
If doTrim
    sSelection := Trim(sSelection,"`n`r`t`s")

Clipboard := OldClipboard 

return sSelection
} 
; -------------------------------------------------------------------------------------------------------------------


Clip_GetSelectionHtml(){
; Get selection in html format
; Syntax:
; 	sSelection := GetSelectionHtml(restore:= True)
sSelection := Clip_GetSel("html",False)
return sSelection
} 
; -------------------------------------------------------------------------------------------------------------------
Clip_GetSelection(){
; Get selection in plain text format
; Syntax:
; 	sSelection := GetSelection(restore:= True)
sSelection := Clip_GetSel("text",False)
return sSelection
} 
; -------------------------------------------------------------------------------------------------------------------

Clip_GetHtml( byref Data ) { ; www.autohotkey.com/forum/viewtopic.php?p=392624#392624
 If CBID := DllCall( "RegisterClipboardFormat", Str,"HTML Format", UInt )
  If DllCall( "IsClipboardFormatAvailable", UInt,CBID ) <> 0
   If DllCall( "OpenClipboard", UInt,0 ) <> 0
    If hData := DllCall( "GetClipboardData", UInt,CBID, UInt )
       DataL := DllCall( "GlobalSize", UInt,hData, UInt )
     , pData := DllCall( "GlobalLock", UInt,hData, UInt )
     , VarSetCapacity( data, dataL * ( A_IsUnicode ? 2 : 1 ) ), StrGet := "StrGet"
     , A_IsUnicode ? Data := %StrGet%( pData, dataL, 0 )
                   : DllCall( "lstrcpyn", Str,Data, UInt,pData, UInt,DataL )
     , DllCall( "GlobalUnlock", UInt,hData )
 DllCall( "CloseClipboard" )
Return dataL ? dataL : 0
}

; -------------------------------------------------------------------------------------------------------------------
; https://www.autohotkey.com/boards/viewtopic.php?f=6&t=80706&sid=af626493fb4d8358c95469ef05c17563
SetClipboardHTML(HtmlBody, HtmlHead:="", AltText:="") {       ; v0.66 by SKAN on D393/D396
Local  F, Html, pMem, Bytes, hMemHTM:=0, hMemTXT:=0, Res1:=1, Res2:=1   ; @ tiny.cc/t80706
Static CF_UNICODETEXT:=13,   CFID:=DllCall("RegisterClipboardFormat", "Str","HTML Format")

  If ! DllCall("OpenClipboard", "Ptr",A_ScriptHwnd)
    Return 0
  Else DllCall("EmptyClipboard")

  If (HtmlBody!="")
  {
      Html     := "Version:0.9`r`nStartHTML:00000000`r`nEndHTML:00000000`r`nStartFragment"
               . ":00000000`r`nEndFragment:00000000`r`n<!DOCTYPE>`r`n<html>`r`n<head>`r`n"
                         . HtmlHead . "`r`n</head>`r`n<body>`r`n<!--StartFragment -->`r`n"
                              . HtmlBody . "`r`n<!--EndFragment -->`r`n</body>`r`n</html>"

      Bytes    := StrPut(Html, "utf-8")
      hMemHTM  := DllCall("GlobalAlloc", "Int",0x42, "Ptr",Bytes+4, "Ptr")
      pMem     := DllCall("GlobalLock", "Ptr",hMemHTM, "Ptr")
      StrPut(Html, pMem, Bytes, "utf-8")

      F := DllCall("Shlwapi.dll\StrStrA", "Ptr",pMem, "AStr","<html>", "Ptr") - pMem
      StrPut(Format("{:08}", F), pMem+23, 8, "utf-8")
      F := DllCall("Shlwapi.dll\StrStrA", "Ptr",pMem, "AStr","</html>", "Ptr") - pMem
      StrPut(Format("{:08}", F), pMem+41, 8, "utf-8")
      F := DllCall("Shlwapi.dll\StrStrA", "Ptr",pMem, "AStr","<!--StartFra", "Ptr") - pMem
      StrPut(Format("{:08}", F), pMem+65, 8, "utf-8")
      F := DllCall("Shlwapi.dll\StrStrA", "Ptr",pMem, "AStr","<!--EndFragm", "Ptr") - pMem
      StrPut(Format("{:08}", F), pMem+87, 8, "utf-8")

      DllCall("GlobalUnlock", "Ptr",hMemHTM)
      Res1  := DllCall("SetClipboardData", "Int",CFID, "Ptr",hMemHTM)
  }

  If (AltText!="")
  {
      Bytes    := StrPut(AltText, "utf-16")
      hMemTXT  := DllCall("GlobalAlloc", "Int",0x42, "Ptr",(Bytes*2)+8, "Ptr")
      pMem     := DllCall("GlobalLock", "Ptr",hMemTXT, "Ptr")
      StrPut(AltText, pMem, Bytes, "utf-16")
      DllCall("GlobalUnlock", "Ptr",hMemHTM)
      Res2  := DllCall("SetClipboardData", "Int",CF_UNICODETEXT, "Ptr",hMemTXT)
  }

  DllCall("CloseClipboard")

  hMemHTM := hMemHTM ? DllCall("GlobalFree", "Ptr",hMemHTM) : 0
  hMemTXT := hMemTXT ? DllCall("GlobalFree", "Ptr",hMemTXT) : 0

Return (Res1 & Res2)
} ; eofun
; -------------------------------------------------------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------

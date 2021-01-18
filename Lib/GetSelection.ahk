GetSelection(type:="text",doRestoreClip:=True){
; Syntax:
; 	sSelection := GetSelection("text"*|"html")
; Default "text"
; Output is trimmed

; Calls: WinClip.GetHTML

If (doRestoreClip = True)
  OldClipboard:= ClipboardAll                         ;Save existing clipboard.

Clipboard:=""
while(Clipboard){
  Sleep,10
}
Send,^c                                          ;Copy selected text to clipboard
ClipWait 1
If ErrorLevel {
  Clipboard:= OldClipboard 
  return
}

If (type = "text") {
  sSelection := clipboard
  ;sSelection := WinClip.GetText()
} Else If (type ="html") {
  sSelection := WinClip.GetHTML()
}

sSelection := Trim(sSelection,"`n`r`t`s")

If (doRestoreClip = True) ; Restore Clipboard
  Clipboard := OldClipboard 

return sSelection
} 
GetSelection(type:="text"){
; Syntax:
; 	sSelection := GetSelection("text"*|"html")
; Default "text"
; Calls: WinClip.GetHTML

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

; Restore Clipboard
Clipboard := OldClipboard 
return sSelection
} 
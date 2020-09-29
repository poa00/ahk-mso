BrowserGetPage(sUrl){
; Syntax:  sHtml := BrowserGetPage(sUrl*)
; Called by ConNext->CNGetAtom and ConNextEnhancer
; If no input Url, copy current page
ClipboardSaved := ClipboardAll

If sUrl {
	Clipboard := sUrl
	Send ^t ; Open new tab
	Send ^v
	Send {Enter}
	Sleep 3000 ; time to load the page
	; TODO: optimize timing
	; https://jacksautohotkeyblog.wordpress.com/2018/05/02/waiting-for-a-web-page-to-load-into-a-browser-autohotkey-tips/
}

;SendInput %sUrl%{Enter}

Clipboard =
Send ^a
Send ^a
Sleep 500
Send ^c

ClipWait, 5 ; Wait max 5 seconds
If ErrorLevel
{
    MsgBox, The attempt to copy text onto the clipboard failed.
    Return
}
   
sHtml := Clipboard
If sUrl {
	Sleep 500
	Send ^w ; close window 
}
Else
	Click ; Deselect

Clipboard := ClipboardSaved
return sHtml

}
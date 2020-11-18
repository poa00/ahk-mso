Outlook_PersonalizeMentions(sName:=""){
 If GetKeyState("Ctrl") {
	Run, "https://tdalon.blogspot.com/2020/11/teams-shortcuts-personalize-mentions.html"
	return
}
If (sName="") {
    SendInput +{Left}
    sLastLetter := GetSelection()
    SendInput {Right}
} Else
    sLastLetter := SubStr(sName,0)

If (sLastLetter = ")") {
    SendInput +{Backspace}+{Backspace}+{Backspace}^{Left}^{Backspace}^{Backspace}^{Right}
} Else {
    SendInput ^{Left}^{Backspace}^{Backspace}^{Right}
}
  
; Remove numbers from mention
SendInput +{Left}
sLastLetter := GetSelection()
If RegExMatch(sLastLetter,"\d")
    SendInput {Delete}
SendInput +{Left}
sLastLetter := GetSelection()
If RegExMatch(sLastLetter,"\d")
    SendInput {Delete}
SendInput {Right}
} ; eofun 
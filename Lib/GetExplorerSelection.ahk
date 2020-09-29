; -------------------------------------------------------------------------------------------------------------------
; Function GetExplorerSelection
; Called by: Main NWS.ahk File Explorer->Ctrl+E, Ctrl+O
; sFile := Explorer_GetSelection
Explorer_GetSelection() { 
 
if WinActive("ahk_exe explorer.exe") {
	sFile := Explorer_GetSelected()
	If (sFile = "")
		sFile := Explorer_GetPath()
	return sFile
} else {
	ClipSaved := ClipboardAll
	Clipboard := ""
	Send ^c
	ClipWait, 0.5
	sFile := Clipboard
	If (sFile = "")  {		
		if WinActive("ahk_exe FreeCommander.exe") {
			Send ^!{Ins}   ; Ctrl+Alt+Ins (Default Shortcut in FC to copy path to clipboard)
			ClipWait, 0.5
			sFile := Clipboard			
		}
    }
	Clipboard := ClipSaved
    return sFile
} ; end else not File Explorer
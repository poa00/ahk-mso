; -------------------------------------------------------------------------------------------------------------------
; https://autohotkey.com/board/topic/17367-url-encoding-and-decoding-of-special-characters/?p=112822
uriDecode(str) {	
    doc := ComObjCreate("HTMLfile")
    doc.write("<body><script>document.write(decodeURIComponent(""" . str . """));</script>")
    sleep 500
	strout := doc.body.innerText

    ; sometimes does not work - timing issue -> backup solution
    If (strout=""), {
        strout := str
    }
        
    ; Example %2520 -> %20%
    Loop
	If RegExMatch(strout, "i)(?<=%)[\da-f]{1,2}", hex)
		StringReplace, strout, strout, `%%hex%, % Chr("0x" . hex), All
	Else Break
    Return, strout
}

; Test 
; https://continental.sharepoint.com/:p:/r/teams/team_10000778/Shared%20Documents/Explore/New%20Work%20Style%20%E2%80%93%20O365%20-%20Why%20using%20Teams.pptx?d=we1a512b97ed844fc92dd5a1d028ef827&csf=1&e=crdehv

; Previous version does not decode UTF-8
;


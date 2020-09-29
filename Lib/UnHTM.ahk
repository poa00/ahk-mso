; Remove HTML formatting / Convert to ordinary text
; Forum Topic: www.autohotkey.com/forum/topic51342.html  Mod: 16-Sep-2010
; Syntax: sHtml := UnHTM(sHtml)
UnHTM( HTM ) { ; Remove HTML formatting / Convert to ordinary text     by SKAN 19-Nov-2009
 Static HT     ; Forum Topic: www.autohotkey.com/forum/topic51342.html
 IfEqual,HT,,   SetEnv,HT, % "&aacuteá&acircâ&acute´&aeligæ&agraveà&amp&aringå&atildeã&au"
 . "mlä&bdquo„&brvbar¦&bull•&ccedilç&cedil¸&cent¢&circˆ&copy©&curren¤&dagger†&dagger‡&deg"
 . "°&divide÷&eacuteé&ecircê&egraveè&ethð&eumlë&euro€&fnofƒ&frac12½&frac14¼&frac34¾&gt>&h"
 . "ellip…&iacuteí&icircî&iexcl¡&igraveì&iquest¿&iumlï&laquo«&ldquo“&lsaquo‹&lsquo‘&lt<&m"
 . "acr¯&mdash—&microµ&middot·&nbsp &ndash–&not¬&ntildeñ&oacuteó&ocircô&oeligœ&ograveò&or"
 . "dfª&ordmº&oslashø&otildeõ&oumlö&para¶&permil‰&plusmn±&pound£&quot""&raquo»&rdquo”&reg"
 . "®&rsaquo›&rsquo’&sbquo‚&scaronš&sect§&shy­&sup1¹&sup2²&sup3³&szligß&thornþ&tilde˜&tim"
 . "es×&trade™&uacuteú&ucircû&ugraveù&uml¨&uumlü&yacuteý&yen¥&yumlÿ"
 TXT := RegExReplace( HTM,"<[^>]+>" )               ; Remove all tags between  "<" and ">"
 Loop, Parse, TXT, &`;                              ; Create a list of special characters
   L := "&" A_LoopField ";", R .= (!(A_Index&1)) ? ( (!InStr(R,L,1)) ? L:"" ) : ""
 StringTrimRight, R, R, 1
 Loop, Parse, R , `;                                ; Parse Special Characters
  If F := InStr( HT, A_LoopField )                  ; Lookup HT Data
    StringReplace, TXT,TXT, %A_LoopField%`;, % SubStr( HT,F+StrLen(A_LoopField), 1 ), All
  Else If ( SubStr( A_LoopField,2,1)="#" )
    StringReplace, TXT, TXT, %A_LoopField%`;, % Chr(SubStr(A_LoopField,3)), All
Return RegExReplace( TXT, "(^\s*|\s*$)")            ; Remove leading/trailing white spaces
}
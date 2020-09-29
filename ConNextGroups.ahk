; ConNext Groups
; Author: Thierry Dalon
; See user documentation here: https://connext.conti.de/blogs/tdalon/entry/connext_quick_mentions
; Code Project Documentation is available on ContiSource GitHub here: https://github.com/tdalon/ahk
;  Source http://github.conti.de/ContiSource/ahk/blob/master/ConNextGroups.ahk
;

; Included in ConNextEnhancer Submenu Mentions


Menu, ConNextGroupsMenu, add, Team A, PasteMentionsTeamA
Menu, ConNextGroupsMenu, add, Team B, PasteMentionsTeamB
;Menu, ConNextMenu, add, Groups, :ConNextGroupSubmenu



PasteMentionsTeamA:
	sHTMLCode = <span contenteditable="false" class="vcard"><a href="https://connext.conti.de/profiles/html/profileView.do?email=Stefanie.Preisinger@continental-corporation.com" class="fn url">@Stefanie Preisinger</a><span class="x-lconn-userid" style="display: none;">D094F2537C49B7F5C12581BE003D51CD</span></span>, <span contenteditable="false" class="vcard"><a href="https://connext.conti.de/profiles/html/profileView.do?email=Juergen.Hagg@continental-corporation.com" class="fn url">@Juergen Hagg</a><span class="x-lconn-userid" style="display: none;">E57AFF9BB7CE969FC12573A6002F4366</span></span>, <span contenteditable="false" class="vcard"><a href="https://connext.conti.de/profiles/html/profileView.do?email=Florian.2.Walzer@continental-corporation.com" class="fn url">@Florian Walzer</a><span class="x-lconn-userid" style="display: none;">88F4EFEB815B9F71C1257F8C003B0D20</span></span>, <span contenteditable="false" class="vcard"><a href="https://connext.conti.de/profiles/html/profileView.do?email=Markus.3.Zacher@continental-corporation.com" class="fn url">@Markus Zacher</a><span class="x-lconn-userid" style="display: none;">625C35380C33231CC125807400551D1C</span></span>, <span contenteditable="false" class="vcard"><a href="https://connext.conti.de/profiles/html/profileView.do?email=Eduard.Kumer@continental-corporation.com" class="fn url">@Eduard Kumer</a><span class="x-lconn-userid" style="display: none;">772CDCE499CD3BB1C12573A600306C6B</span></span>, <span contenteditable="false" class="vcard"><a href="https://connext.conti.de/profiles/html/profileView.do?email=Beate.Bahl@continental-corporation.com" class="fn url">@Beate Bahl</a><span class="x-lconn-userid" style="display: none;">126D032D78A39C5CC12573A6002D8C63</span></span>
	runfclip(sHTMLCode,sHTMLCode)
	Send ^v
return

PasteMentionsTeamB:
	sHTMLCode = <span contenteditable="false" class="vcard"><a href="https://connext.conti.de/profiles/html/profileView.do?email=thierry.dalon@continental-corporation.com" class="fn url">@Thierry Dalon</a><span class="x-lconn-userid" style="display: none;">9C6BF7AEB4195B66C125803600416EFA</span></span>, <span contenteditable="false" class="vcard"><a href="https://connext.conti.de/profiles/html/profileView.do?email=Martina.Schoepf@continental-corporation.com" class="fn url">@Martina Schoepf</a><span class="x-lconn-userid" style="display: none;">32026521B75E7AC1C1257EF900539BB7</span></span>, <span contenteditable="false" class="vcard"><a href="https://connext.conti.de/profiles/html/profileView.do?email=Hanspeter.Zink@continental-corporation.com" class="fn url">@Hanspeter Zink</a><span class="x-lconn-userid" style="display: none;">E84D5844B55FFC0AC12573A60032FB5E</span></span>, <span contenteditable="false" class="vcard"><a href="https://connext.conti.de/profiles/html/profileView.do?email=Nicole.Locke@continental-corporation.com" class="fn url">@Nicole Locke</a><span class="x-lconn-userid" style="display: none;">B565EBA0A49EE88985257B6000551327</span></span>
	runfclip(sHTMLCode,sHTMLCode)
	Send ^v
return
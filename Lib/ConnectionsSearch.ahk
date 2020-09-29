; ConNext Search - Search within current ConNext App

; Called by: MWS.ahk (Win+F Hotkey)

#Include <Connections>

ConnectionsSearch(sUrl, newWindow := True){
If GetKeyState("Ctrl") and !GetKeyState("Shift") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/connext_search_ahk"
	return
}
static sWikiSearch, sWikiLabel
static sForumSearch, sForumUuid
static sBlogSearch, sBlogId

If Connections_IsConnectionsUrl(sUrl,"wiki") {
	sOldWikiLabel := sWikiLabel
	sWikiLabel := GetWikiLabel(sUrl)
	If !sWikiLabel
		return
	sDefSearch =
	If sWikiLabel = %sOldWikiLabel%
		sDefSearch := sWikiSearch

	If (sDefSearch = "") { ; isempty
		ReConnectionsRootUrl := StrReplace(PowerTools_ConnectionsRootUrl,".","\.")
		; https://connext.conti.de/wikis/home?lang=en#!/wiki/W354104eee9d6_4a63_9c48_32eb87112262/index?sort=mostpopular&tag=ms_teams&tag=meeting
		; https://connext.conti.de/wikis/home?lang=en#!/wiki/W965feafa2312_41f6_b6c9_63bbc7c9a59b/index?tag=ms_outlook
		If RegExMatch(sUrl,ReConnectionsRootUrl . "/wikis/.*[&\?]tag=([^&\?]*)",sMatch) { ; tag-based			
			sPat := "[&\?]tag=([^&]*)" 
			Pos=1
			While Pos :=    RegExMatch(sUrl, sPat, tag,Pos+StrLen(tag)) 
				sTags := sTags . " #" . tag1 
			sTags := Trim(sTags) ; remove starting space			
			sDefSearch := sTags
			
		; keyword-based https://connext.conti.de/wikis/home#!/search?query=ms_teams&mode=this&wikiLabel=W354104eee9d6_4a63_9c48_32eb87112262
		} Else If RegExMatch(sUrl,ReConnectionsRootUrl . "/wikis/.*[\?&]query=([^&\?]*)",sMatch) { ; keyword-based
			sDefSearch := StrReplace(sMatch1,"%20"," ") 
		} 

	}	

	InputBox, sSearch , Wiki Search, Enter search string (use # for tags):,,640,125,,,,,%sDefSearch% 
	if ErrorLevel
		return
	sWikiSearch := Trim(sSearch) 
	SearchWiki(sWikiLabel,sWikiSearch)

} Else If Connections_IsConnectionsUrl(sUrl,"forum") {
	sForumUuid := CNGetForumId(sUrl)
	If !sForumUuid
		return
	ReConnectionsRootUrl := StrReplace(PowerTools_ConnectionsRootUrl,".","\.")
	; sUrl = https://connext.conti.de/forums/html/forum?id=755c8bb9-52a5-4db8-9ac9-933777b4322d&tags=uservoice&query=multiple%20window&ps=100
	If RegExMatch(sUrl,ReConnectionsRootUrl . "/forums/html/forum\?id=(?:.*?)&tags=([^&\?]*)&query=([^&\?]*)",sMatch) {
		sDefSearch := "#" . StrReplace(sMatch1,"%20"," #") . " " . StrReplace(sMatch2,"%20"," ")
	}
	Else If RegExMatch(sUrl,ReConnectionsRootUrl . "/forums/html/forum\?id=(?:.*?)&query=([^&\?]*)&tags=([^&\?]*)",sMatch) {
		sDefSearch := "#" . StrReplace(sMatch2,"%20"," #") . " " . StrReplace(sMatch1,"%20"," ")
	}
	Else If RegExMatch(sUrl, ReConnectionsRootUrl . "/forums/html/forum\?id=(?:.*?)&tags=([^&\?]*)",sMatch) {
		sDefSearch := "#" . StrReplace(sMatch1,"%20"," #") 
	}
	Else If RegExMatch(sUrl,ReConnectionsRootUrl . "/forums/html/forum\?id=(?:.*?)&query=([^&\?]*)",sMatch) {
		sDefSearch := StrReplace(sMatch1,"%20"," ")
	} Else 
		sDefSearch = 
	
	If InStr(sUrl,"&filter=answeredQuestions")
		sDefSearch := sDefSearch . " -a"
	If InStr(sUrl,"&filter=openQuestions")
		sDefSearch := sDefSearch . " -o"
		
	InputBox, sSearch , Forum Search, Enter search string (use # for tags):,,640,125,,,,,%sDefSearch% 
	if ErrorLevel
		return
	SearchForum(sForumUuid,Trim(sSearch))

} Else If Connections_IsConnectionsUrl(sUrl,"blog") {
	sOldBlogId := sBlogId
	sPat = /blogs/([^/\?]*)[/\?]
	Pos := RegExMatch(sUrl, sPat, sBlogId)
	sBlogId := sBlogId1
	If sBlogId = %sOldBlogId%
		sDefSearch := sBlogSearch
	Else {
		ReConnectionsRootUrl := StrReplace(PowerTools_ConnectionsRootUrl,".","\.")

		; sUrl = https://connext.conti.de/blogs/tdalon?tags=connext%20ms_sharepoint
		If RegExMatch(sUrl,ReConnectionsRootUrl . "/blogs/.*\?tags=([^&\?]*)",sMatch) { ; tag-based
			sDefSearch := "#" . StrReplace(sMatch1,"%20"," #")
		; keyword-based https://connext.conti.de/blogs/tdalon/search?t=entry&q=connext+sharepoint&lang=en
		; https://connext.conti.de/blogs/45ca8e69-8c52-4154-8119-b85b7fdb62a0/search?t=entry&f=all&q=connext&lang=en
		} Else If RegExMatch(sUrl,ReConnectionsRootUrl . "/blogs/.*[\?&]q=([^&\?]*)",sMatch) { ; keyword-based
			sDefSearch := StrReplace(sMatch1,"+"," ") 
		} Else 
			sDefSearch = 
	}
	InputBox, sSearch , Blog Search, Enter search string (use # for tags):,,640,125,,,,,%sDefSearch% 
	if ErrorLevel
		return
	sBlogSearch := Trim(sSearch) 
	SearchBlog(sBlogId,sBlogSearch)		

} Else If InStr(sUrl,PowerTools_ConnectionsRootUrl . "/profiles/html/advancedSearch.do") {
	; https://connext.conti.de/profiles/html/advancedSearch.do?pageSize=150&profileTags=nws_guide&keyword=regensburg
	sSpace := "%20"
	sPat = profileTags=([^&]*)	
	If RegExMatch(sUrl, sPat, sTags)
		sProfileSearch := "#" .  StrReplace(sTags1,sSpace," #")
	sPat = keyword=([^&]*)	
	If RegExMatch(sUrl, sPat, sKeywords)
		sProfileSearch := sProfileSearch . " " . StrReplace(sKeywords1,sSpace," ")

	sProfileSearch := StrReplace(sProfileSearch,"%26","&")
	;sDefSearch := StrReplace(sDefSearch,"%23","#")

	InputBox, sProfileSearch , Profiles Search, Enter search string (use # for tags):,,640,125,,,,,%sProfileSearch% 
	if ErrorLevel
		return
	sProfileSearch := Trim(sProfileSearch) 
	SearchProfiles(sProfileSearch)	
} Else {
	TrayTipAutoHide("NWS PowerTool","No Connections Search available on this page!")
}

} ; End Main Function

; #############################################################################################
; SubFunctions

SearchForum(sForumUuid,sSearch,newWindow:=False){
sUrl = https://%PowerTools_ConnectionsRootUrl%/forums/html/forum?id=%sForumUuid%&ps=500

; Check for option -o or -a
sPat = \s*-o\s*
If RegExMatch(sSearch, sPat, sMatch) {
	sSearch := StrReplace(sSearch, sMatch,"")
	sUrl := sUrl . "&filter=openQuestions"
}

sPat = \s*-a\s*
If RegExMatch(sSearch, sPat, sMatch) {
	sSearch := StrReplace(sSearch, sMatch,"")
	sUrl := sUrl . "&filter=answeredQuestions"
}

; Extract Tags
sPat = #([^#\s]*)
sSpace := "%20"
Pos=1
While Pos :=    RegExMatch(sSearch, sPat, tag,Pos+StrLen(tag)) 
	sTags := sTags . sSpace . tag1    

; Remove start sSpace
sTags := SubStr(sTags,StrLen(sSpace)+1)
If !(sTags="")  ; not empty
	sTags := "&tags=" . sTags	
sSearch := RegExReplace(sSearch, sPat , "")
sSearch := Trim(sSearch)
sUrl := sUrl . sTags 

If  (sSearch != "")
	sUrl := sUrl . "&query=" sSearch

If !newWindow
	Send ^w ; Close previous search window
Run, %sUrl%
} ; End function SearchForum



GetWikiLabel(sUrl){
sPat = /wiki/([^/]*)/?
; Search view
If RegExMatch(sUrl, sPat, sWikiLabel) 
	return sWikiLabel1
sPat := "&wikiLabel=([^&]*)"
If RegExMatch(sUrl, sPat, sWikiLabel) 
	return sWikiLabel1
	
MsgBox,48,Warning!,WikiLabel not found!
return

}

SearchWiki(sWikiLabel,sSearch,newWindow:= False){

sPat := "#([^#\s]*)" 
sSpace :="%20"
Pos=1
While Pos :=    RegExMatch(sSearch, sPat, tag,Pos+StrLen(tag)) 
	sTags := sTags . "&tag=" . tag1 


sSearch := RegExReplace(sSearch, sPat , "")
sSearch := Trim(sSearch)


If  (sSearch != "") { ; not empty
	sSearch := sSearch . StrReplace(sTags,"&tag="," ")  ; convert tags to query
	sUrl := "https://" . PowerTools_ConnectionsRootUrl . "/wikis/home/search?query=" . sSearch . "&mode=this&wikiLabel=" . sWikiLabel
} Else {
	sUrl := "https://" . PowerTools_ConnectionsRootUrl . "/wikis/home/wiki/" . sWikiLabel . "/index?sort=mostpopular" . sTags
}	
If !newWindow
	Send ^w ; Close previous search window
Run, %sUrl%

} ; End function SearchWiki

SearchBlog(sBlogId,sSearch,newWindow:= False){

sUrl :=  "https://" . PowerTools_ConnectionsRootUrl . "/blogs/" . sBlogId 
sPat := "#([^#\s]*)" 
sSpace :="%20"
Pos=1
sTags := ""
While Pos :=    RegExMatch(sSearch, sPat, tag,Pos+StrLen(tag)) 
	sTags := sTags . sSpace . tag1    

; Remove start sSpace
sTags := SubStr(sTags,StrLen(sSpace)+1)
	
sSearch := RegExReplace(sSearch, sPat , "")
sSearch := Trim(sSearch)

If  (sSearch != ""){
	sSearch := StrReplace(sSearch," ","+")
	If  (sTags != ""){
		sSearch := sSearch . "+" . StrReplace(sTags,sSpace,"+")  ; convert tags to query
	}
	sUrl := sUrl . "/search?t=entry&q=" sSearch
} Else { ; Tag based search
	sUrl := sUrl . "?tags=" . sTags
}	
If !newWindow
	Send ^w ; Close previous search window
Run, %sUrl%
}


SearchProfiles(sSearch,newWindow := False){

sSearch := StrReplace(sSearch,"&","%26")

sUrl =  https://%PowerTools_ConnectionsRootUrl%/profiles/html/advancedSearch.do?pageSize=150
sPat := "#([^#\s]*)" 
sSpace :="%20"
Pos=1
sTags := ""
While Pos :=    RegExMatch(sSearch, sPat, tag,Pos+StrLen(tag)) 
	sTags := sTags . sSpace . tag1    
; Remove start sSpace
sTags := SubStr(sTags,StrLen(sSpace)+1)
	
sSearch := RegExReplace(sSearch, sPat , "")
sSearch := Trim(sSearch)

If  (sSearch != ""){
	sSearch := StrReplace(sSearch," ",sSpace)
	sUrl := sUrl . "&keyword=" . sSearch
} 
If  (sTags != ""){
	sUrl := sUrl . "&profileTags=" . sTags
}	
If !newWindow
	Send ^w ; Close previous search window
Run, %sUrl%
}
; Library File for People Social Functions
; Used by PeopleConnector and ConNextEnhancer

#Include <Teams>
#Include <ConNext>

; ----------------------------------------------------------------------

GetEmailList(sInput){
; Get EmailList from input string
; If no email is found look for html selection
;   sEmailList := GetEmailList(sInput)
; Calls: ExtractEmails

sEmailList := ExtractEmails(sInput)

If (sEmailList = "")  { 
    sInput := GetSelection("html")
    sInput := StrReplace(sInput,"%40","@") ; for connext profile links
    sEmailList := ExtractEmails(sInput)
}
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
return sEmailList
} ; eof

; ----------------------------------------------------------------------

ExtractEmails(sInput){
; Extract Email from String e.g. copied to clipboard Outlook addresses or Html source
; Syntax: sEmailList := ExtractEmails(sInput)
; Called by: ConNextEnhancer, MyScript

sPat = [0-9a-zA-Z\.\-]+@[0-9a-zA-Z\-]*\.[a-z]{2,3}
; TODO bug if Connext mention
Pos = 1 
While Pos := RegExMatch(sInput,sPat,sMatch,Pos+StrLen(sMatch)){
    If InStr(sEmailList,sMatch . ";")
        continue
    sEmailList = %sEmailList%;%sMatch%
}

sPat = https?://connext.conti.de/profiles/html/profileView.do\?userid=([0-9A-Z]*)
sPat := StrReplace(sPat,".","\.")
Pos = 1 
While Pos := RegExMatch(sInput,sPat,sMatch,Pos+StrLen(sMatch)){
    sUid := sMatch1
    If InStr(sUidList,sUid . ";")
        continue
    sEmail := CNUid2Email(sUid)
    If InStr(sEmailList,sEmail . ";")
        continue
    sEmailList = %sEmailList%;%sEmail%
    sUidList = %sUidList%;%sUid%
}
return SubStr(sEmailList,2) ; remove starting ;
} ; eof
; ----------------------------------------------------------------------

Email2Uid(sEmail,FieldName:="mailNickname"){
    sUid := AD_GetUserField("mail=" . sEmail, FieldName) ; mailNickname - office uid 
    ;sWinUid := AD_GetUserField("mail=" . sEmail, "sAMAccountName")  ;- login uid
    ;sOfficeUid := AD_GetUserField("mail=" . sEmail, "mailNickname")  ;- Office uid
    return sUid
}
; ----------------------------------------------------------------------
Emails2Uids(sEmailList,FieldName:="mailNickname"){
;Emails2Uids(sEmailList,FieldName:="mailNickname"|"sAMAccountName")
Loop, parse, sEmailList, ";"
{
    sUid := Email2Uid(A_LoopField,FieldName)
    sUidList = %sUidList%, %sUid%
}	
return SubStr(sUidList,2) ; remove starting ;
}

; ----------------------------------------------------------------------

winUid2Email(sUid){
sEmail := AD_GetUserField("sAMAccountName=" . sUid, "mail")
return sEmail
}
; ----------------------------------------------------------------------

winUids2Emails(sUidList){
Loop, parse, sUidList, `;%A_Tab%`,
{
    sEmail := winUid2Email(Trim(A_LoopField))
    sEmailList = %sEmailList%;%sEmail%
}	
return SubStr(sEmailList,2) ; remove starting ;
}

; ----------------------------------------------------------------------
AD_GetUserField(sFilter, sField){
; dsquery * -gc dc=contiwan,dc=com -filter "(&(objectClass=user)(mail=benjamin*dosch*))" -attr name sAMAccountName mailNickname
; https://connext.conti.de/forums/html/topic?id=d84ff1e7-8d23-45bb-8ee6-0e3b46408288&permalinkReplyUuid=2fa26181-d2a3-48f2-b195-4f5fb0c49bd8
    strADPath := "GC://dc=contiwan,dc=com" 
	; ADODB Connection to AD
	objConnection := ComObjCreate("ADODB.Connection")
	objConnection.Open("Provider=ADsDSOObject")
	
	; Connection
	objCommand := ComObjCreate("ADODB.Command")
	objCommand.ActiveConnection := objConnection
	
	; Search the AD recursively, starting at root of the domain
	objCommand.CommandText := "<" . strADPath . ">" . ";(&(objectCategory=User)(" . sFilter . "));" . sField . ";subtree"

    objRecordSet := objCommand.Execute
	; Get the record set to be returned later
	if (objRecordSet.RecordCount == 0){
        strTxt :=  "No Data"  ; no records returned
    }else{
        strTxt := objRecordSet.Fields(sField).Value  ; return value
    }
	
	; Close connection
	objConnection.Close()
	
	; Cleanup
	ObjRelease(objRecordSet)
	ObjRelease(objCommand)
	ObjRelease(objConnection)
	
	; And return the record set
	return strTxt
}

OL2XL(sSelection){

if (SubStr(sSelection,0,1) != ";") 
    sSelection := sSelection . ";"

oExcel := ComObjCreate("Excel.Application") ;handle
oExcel.Workbooks.Add ;add a new workbook
oSheet := oExcel.ActiveSheet
; First Row Header
oSheet.Range("A1").Value := "LastName"
oSheet.Range("B1").Value := "FirstName"
oSheet.Range("C1").Value := "email"

oExcel.Visible := True ;by default excel sheets are invisible
oExcel.StatusBar := "Copy to Excel..."
sPat = U)(.*) <(.*)>;
Pos = 1 
RowCount = 2
While Pos := RegExMatch(sSelection,sPat,sMatch,Pos+StrLen(sMatch)){
    
    Email := sMatch2
    If (InStr(sMatch1,",")) {
        LastName := RegexReplace(sMatch1,",.*","")
        FirstName := RegexReplace(sMatch1,".*,","")
    FirstName := RegExReplace(FirstName," \(.*\)","") ; Remove (uid) in firstname

    } Else {
        FirstName := RegexReplace(Email,"\..*","")
        StringUpper, FirstName, FirstName , T
        LastName := StrReplace(sMatch1,FirstName,"")
    }
    oSheet.Range("A" . RowCount).Value := LastName
    oSheet.Range("B" . RowCount).Value := FirstName
    oSheet.Range("C" . RowCount).Value := Email
    RowCount +=1
}

; expression.Add (SourceType, Source, LinkSource, XlListObjectHasHeaders, Destination, TableStyleName)
oTable := oSheet.ListObjects.Add(1, oSheet.UsedRange,,1)
oTable.Name := "OutlookExport"

oTable.Range.Columns.AutoFit

oExcel.StatusBar := "READY"
}


; Library File for People Social Functions
; Used by PeopleConnector and ConNextEnhancer

#Include <Teams>
#Include <Connections>
global PowerTools_ConnectionsRootUrl

; ----------------------------------------------------------------------

; ----------------------------------------------------------------------

People_GetEmailList(sInput){
; Get EmailList from input string
; Extract Emails from String e.g. copied to clipboard Outlook addresses or Html source
; List is separated with a ;
; Syntax: sEmailList := People_GetEmailList(sInput)

global PowerTools_ConnectionsRootUrl
sInput := StrReplace(sInput,"%40","@") ; for connext profile links - decode @

sPat = [0-9a-zA-Z\.\-]+@[0-9a-zA-Z\-\.]*\.[a-z]{2,3}
; TODO bug if Connext mention
Pos = 1 
While Pos := RegExMatch(sInput,sPat,sMatch,Pos+StrLen(sMatch)){
    If InStr(sEmailList,sMatch . ";")
        continue
    If InStr(sMatch, "@thread.sky") ; skip from Teams conversation
        continue
    sEmailList := sEmailList . sMatch . ";"
}

sPat = https?://%PowerTools_ConnectionsRootUrl%/profiles/html/profileView.do\?userid=([0-9A-Z]*)
sPat := StrReplace(sPat,".","\.")
Pos = 1 
While Pos := RegExMatch(sInput,sPat,sMatch,Pos+StrLen(sMatch)){
    sUid := sMatch1
    If InStr(sUidList,sUid . ";")
        continue
    sEmail := CNUid2Email(sUid)
    If InStr(sEmailList,sEmail . ";")
        continue
    sEmailList := sEmailList . sEmail . ";"
    sUidList := sUidList . sUid . ";"
}
return SubStr(sEmailList,1,-1) ; remove ending ;
} ; eof
; ----------------------------------------------------------------------

Email2Uid(sEmail,FieldName:="mailNickname"){
    sUid := People_ADGetUserField("mail=" . sEmail, FieldName) ; mailNickname - office uid 
    ;sWinUid := People_ADGetUserField("mail=" . sEmail, "sAMAccountName")  ;- login uid
    ;sOfficeUid := People_ADGetUserField("mail=" . sEmail, "mailNickname")  ;- Office uid
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
return SubStr(sUidList,2) ; remove starting ,
}

; ----------------------------------------------------------------------

winUid2Email(sUid){
sEmail := People_ADGetUserField("sAMAccountName=" . sUid, "mail")
return sEmail
}
; ----------------------------------------------------------------------

winUids2Emails(sUidList){
Loop, parse, sUidList, `;%A_Tab%`,
{
    sEmail := winUid2Email(Trim(A_LoopField))
    sEmailList := sEmailList . ";" . sEmail
}	
return SubStr(sEmailList,2) ; remove starting ;
}

; ----------------------------------------------------------------------
People_oUid2Email(sUid,sDomain :="") {
; sUid = uid41890@contiwan.com
If (sDomain="")
    sDomain := RegExReplace(sUid,".*@","")
mail := People_ADGetUserField("mailNickname=" . sUid, "mail",sDomain) ; mailNickname = office uid 
return mail
}

; ----------------------------------------------------------------------
People_ADGetUserField(sFilter, sField,sDomain :=""){
; dsquery * -gc dc=contiwan,dc=com -filter "(&(objectClass=user)(mail=benjamin*dosch*))" -attr name sAMAccountName mailNickname
; https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2012-r2-and-2012/cc754232(v=ws.11)
; https://connext.conti.de/forums/html/topic?id=d84ff1e7-8d23-45bb-8ee6-0e3b46408288&permalinkReplyUuid=2fa26181-d2a3-48f2-b195-4f5fb0c49bd8
    
    If (sDomain ="") {
        sDomain := People_GetDomain()
        If (sDomain ="") 
            return "ERROR: Domain not provided"           
    }
    ; strADPath := "GC://dc=contiwan,dc=com" 
    strADPath := "GC://dc=" . StrReplace(sDomain,".",",dc=") 
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

; ----------------------------------------------------------------------
People_GetDomain(){
RegRead, Domain, HKEY_CURRENT_USER\Software\PowerTools, Domain
If (Domain=""){
    Domain := People_SetDomain()
}
return Domain
}
; ----------------------------------------------------------------------
People_SetDomain(){
RegRead, Domain, HKEY_CURRENT_USER\Software\PowerTools, Domain
InputBox, Domain, Domain, Enter your Domain, , 200, 125,,,,, %Domain%
If ErrorLevel
    return
PowerTools_RegWrite("Domain",Domain)
return Domain
} ; eofun

; ----------------------------------------------------------------------
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

; ----------------------------------------------------------------------
People_GetName(sSelection) {
; Get Name from selection e.g. Outlook Person or email adress
; sName := GetName(sSelection)

sEmailPat = [^\s@]+@[^\s\.]*\.[a-z]{2,3}
;sPat = [0-9a-zA-Z\.\-]+@[0-9a-zA-Z\-\.]*\.[a-z]{2,3}

; From Outlook field
sPat = U)(.*,.*) \<(%sEmailPat%)\>
If RegExMatch(sSelection,sPat,sMatch){
    sName := SwitchName(sMatch1)
; From email
} Else If RegExMatch(sSelection,sEmailPat,sMatch){
    
    sName := Email2Name(sMatch)
} Else {
    sName := SwitchName(sSelection)
}
return sName
}

; ----------------------------------------------------------------------
Email2Name(Email){
Email := RegexReplace(Email,"@.*","")
FirstName := RegexReplace(Email,"\..*","")
StringUpper, FirstName, FirstName , T
LastName := RegexReplace(Email,".*\.","")
;LastName := StrReplace(Email,FirstName,"")
StringUpper, LastName, LastName , T

sName = %FirstName% %LastName%
; Remove numbers
sName := RegexReplace(sName,"\d*","")
return sName
}

; ----------------------------------------------------------------------

SwitchName(sName){
; keep only first line
If InStr(sName,"`r")
    sName := SubStr(sName,1,InStr(sName,"`r")-1)


If (InStr(sName,",")) {
    LastName := RegexReplace(sName,",.*","")
    FirstName := RegexReplace(sName,".*, ","")
    FirstName := RegExReplace(FirstName," \(.*\)","") ; Remove (uid) in firstname
    sName = %FirstName% %LastName%
}
return sName
}

; ----------------------------------------------------------------------

People_ConnectionsOpenProfile(sSelection){
global PowerTools_ConnectionsRootUrl
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") {
    sName := SwitchName(sSelection)
    Run, https://%PowerTools_ConnectionsRootUrl%/profiles/html/simpleSearch.do?searchBy=name&searchFor=%sName%
} Else {
    Loop, parse, sEmailList, ";"
    {
         Run,  https://%PowerTools_ConnectionsRootUrl%/profiles/html/profileView.do?email=%A_LoopField%
    }	; End Loop Parse Clipboard
}
} ; eofun

; ----------------------------------------------------------------------
People_PeopleView(sSelection){
sEmailList := People_GetEmailList(sSelection)
Loop, parse, sEmailList, ";"
{
        Uid := People_ADGetUserField("mail=" . A_LoopField, "employeeNumber") 
        Run,  https://performancemanager5.successfactors.eu/sf/orgchart?&company=ContiProd&selected_user=%Uid%
}	; End Loop Parse Clipboard
} ; eofun


; ----------------------------------------------------------------------
People_DownloadProfilePicture(sEmail,sFolder){

If GetKeyState("Ctrl") {
	Run, "https://connext.conti.de/blogs/tdalon/entry/emails2profilepic"
	return
}

; Create .ps1 file
PsFile = %A_Temp%\o365_GetProfilePic.ps1
; Fill the file with commands
If FileExist(PsFile)
    FileDelete, %PsFile%

Domain := People_GetDomain()
If (Domain ="") {
    MsgBox 0x10, Teams Shortcuts: Error, No Domain defined!
    return
}

sUserName := People_ADGetUserField("mail=" . sEmail, "mailNickname",Domain)

sText := "$photo=Get-Userphoto -identity %sUserName% -ErrorAction SilentlyContinue`n If($photo.PictureData -ne $null)`n{[io.file]::WriteAllBytes($path,$photo.PictureData)}"

FileAppend, %sText%,%PsFile%

; Run it
RunWait, PowerShell.exe -NoExit -ExecutionPolicy Bypass -Command %PsFile% ;,, Hide
;RunWait, PowerShell.exe -ExecutionPolicy Bypass -Command %PsFile% ,, Hide


} ; eofun
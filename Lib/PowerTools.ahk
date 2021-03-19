; PowerTools Lib

#Include <AHK>
AppList = ConNextEnhancer,MO,NWS,OutlookShortcuts,PeopleConnector,TeamsShortcuts
global Config
Config := PowerTools_GetConfig()

PowerTools_CheckForUptate(ToolName :="") {
If !a_iscompiled {
	Run, https://github.com/tdalon/ahk ; no direct link because of Lib dependencies
    return
} 

; warning if connected via VPN
If (Login_IsVPN()) {
MsgBox, 0x1011, CheckForUpdate with VPN?,It seems you are connected with VPN.`nCheck for update might not work. Consider disconnecting VPN.`nContinue now?
IfMsgBox Cancel
    return
}

If !ToolName    
    ScriptName := A_ScriptName
Else
    ScriptName = %ToolName%.exe
    ; Overwrites by default
sUrl = https://github.com/tdalon/ahk/raw/master/PowerTools/%ScriptName%

ExeFile = %A_ScriptDir%\%ScriptName%
If Not FileExist(ExeFile) {
    UrlDownloadToFile, %sUrl%, %ScriptName%
    return
}

UrlDownloadToFile, %sUrl%, %ScriptName%.github
guExe = %A_ScriptDir%\github_updater.exe
If Not FileExist(guExe)
    UrlDownloadToFile, https://github.com/tdalon/ahk/raw/master/PowerTools/github_updater.exe, %guExe%
    
sCmd = %guExe% %ScriptName%
RunWait, %sCmd%,,Hide
} ; eof

; ---------------------------------------------------------------------- 
PowerTools_Help(ScriptName,doOpen := True){

Switch ScriptName 
{
Case "ConnectionsEnhancer":
    sUrl = https://tdalon.github.io/ahk/Connections-Enhancer
Case "TeamsShortcuts":
    sUrl = https://tdalon.github.io/ahk/Teams-Shortcuts
Case "MO":
    sUrl := "https://connext.conti.de/wikis/home/wiki/Wc4f94c47297c_42c8_878f_525fd907cb68/page/MO%20PowerTool"
Case "PeopleConnector":
    sUrl = https://tdalon.github.io/ahk/People-Connector
Case "OutlookShortcuts":
    sUrl = https://tdalon.github.io/ahk/Outlook-Shortcuts
Case "Teamsy":
    sUrl = https://tdalon.github.io/ahk/Teamsy
Case "TeamsyLauncher":
    sUrl = https://tdalon.github.io/ahk/Teamsy-Launcher
Case "NWS":
    sUrl := "https://tdalon.github.io/ahk/NWS-PowerTool"
Case "Mute":
    sUrl := "https://tdalon.github.io/ahk/Mute-PowerTool"
Case "Bundler":
    sUrl :="https://tdalon.github.io/ahk/PowerTools-Bundler"
Case "Cursor Highlighter":
    sUrl = https://tdalon.github.io/ahk/Cursor-Highlighter
Case "all":
Default:
    sUrl := "https://tdalon.github.io/ahk/PowerTools"	
}

If (Config="Conti"){
    Switch ScriptName 
    {
    Case "ConnectionsEnhancer":
        sUrl = https://connext.conti.de/blogs/tdalon/entry/connext_enhancer
    Case "PeopleConnector":
        sUrl = https://connext.conti.de/blogs/tdalon/entry/people_connector
    Case "OutlookShortcuts":
        sUrl = https://connext.conti.de/blogs/tdalon/entry/outlook_autohotkey_script
    Case "NWS":
        sUrl := "https://connext.conti.de/wikis/home/wiki/Wc4f94c47297c_42c8_878f_525fd907cb68/page/NWS%20PowerTool"
    }  
}

If doOpen
    Run, %sUrl%
return sUrl
} ; eofun

; ---------------------------------------------------------------------- 

PowerTools_Changelog(ScriptName,doOpen := True){
Switch ScriptName 
{
Case "ConnectionsEnhancer":
    sFileName = Connections-Enhancer-Changelog
Case "TeamsShortcuts":
    sFileName = Teams-Shortcuts-Changelog
Case "MO":
    sUrl := "http://github.conti.de/ContiSource/ahk/wiki/MO-(Release-Notes)"
    Run, %sUrl%
    return
Case "PeopleConnector":
    sFileName = People-Connector-Changelog
Case "NWS":
    sFileName = NWS-PowerTool-Changelog
Case "Mute":
    sFileName = Mute-PowerTool-Changelog
Case "Bundler":
    sFileName = PowerTools-Bundler-Changelog
Case "OutlookShortcuts":
    sFileName = Outlook-Shortcuts-Changelog
Case "Teamsy":
    sFileName = Teamsy-Changelog
Case "Cursor Highlighter":
    Run, https://sites.google.com/site/boisvertlab/computer-stuff/online-teaching/cursor-highlighter-changelog
    return
Case "all":
Default:
    sFileName =  PowerTools-Changelogs
}

If Not doOpen {
    sUrl = https://tdalon.github.io/ahk/%sFileName%
    return sUrl
}
    

If !A_IsCompiled {
    sFile = %A_ScriptDir%\docs\_pages\%sFileName%.md
    If FileExist(sFile) {
        ;Run, Open %sFile% ; does not open Atom
        Run notepad++.exe "%sFile%"
        Return
    }
} Else {
    sUrl = https://tdalon.github.io/ahk/%sFileName%
    Run, %sUrl%
}
} ; eofun

; ---------------------------------------------------------------------- 

PowerTools_News(ScriptName){
If (ScriptName ="NWS")
    ScriptName = NWSPowerTool
Else If (ScriptName ="Mute")
    ScriptName = MutePowerTool
sUrl := "https://twitter.com/search?q=(from%3Atdalon)%23" . ScriptName
Switch ScriptName 
{
Case "ConnectionsEnhancer":
    sUrl := sUrl . "%20%23Connections"
Case "TeamsShortcuts":
    sUrl := sUrl . "%20%23MicrosoftTeams"
Case "OutlookShortcuts":
    sUrl := sUrl . "%20%23MicrosoftOutlook"
Case "Teamsy":
    sUrl := sUrl . "%20%23MicrosoftTeams"
}
Run, %sUrl%

} ; eofun

; ---------------------------------------------------------------------- 

PowerTools_RunBundler(){
If a_iscompiled {
  ExeFile = %A_ScriptDir%\PowerToolsBundler.exe
  If Not FileExist(ExeFile) {
    sUrl = https://raw.githubusercontent.com/tdalon/ahk/master/PowerTools/PowerToolsBundler.exe
		UrlDownloadToFile, %sUrl%, PowerToolsBundler.exe
  }
  Run %ExeFile%
} Else
  Run %A_AHKPath% "%A_ScriptDir%\PowerToolsBundler.ahk"

} ; eofun
; ---------------------------------------------------------------------- 

PowerTools_Support(ScriptName){
    Switch ScriptName 
	{
	Case "ConnectionsEnhancer":
        sTeamLink := "https://teams.microsoft.com/l/channel/19%3a6ab774239328402fbe0b6be8bd60b53a%40thread.skype/ConNext%2520Enhancer?groupId=640b2f00-7b35-41b2-9e32-5ce9f5fcbd01&tenantId=8d4b558f-7b2e-40ba-ad1f-e04d79e6265a"
    Case "TeamsShortcuts":
        sTeamLink := "https://teams.microsoft.com/l/channel/19%3a91b56b23eb864738a80a52663387c227%40thread.skype/Teams%2520AutoHotkey?groupId=640b2f00-7b35-41b2-9e32-5ce9f5fcbd01&tenantId=8d4b558f-7b2e-40ba-ad1f-e04d79e6265a"
    Case "MO":
        sTeamLink :="https://teams.microsoft.com/l/channel/19%3a4e24e9f2cc934eb98656cbbd395e765d%40thread.skype/MO%2520PowerTool?groupId=640b2f00-7b35-41b2-9e32-5ce9f5fcbd01&tenantId=8d4b558f-7b2e-40ba-ad1f-e04d79e6265a"
    Case "PeopleConnector":
        sTeamLink := "https://teams.microsoft.com/l/channel/19%3a2b2fb389d3ec405bb985e07a3016ad9a%40thread.skype/People%2520Connector?groupId=640b2f00-7b35-41b2-9e32-5ce9f5fcbd01&tenantId=8d4b558f-7b2e-40ba-ad1f-e04d79e6265a"
    Case "OutlookShortcuts":
        sTeamLink := "https://teams.microsoft.com/l/channel/19%3a82e3119412a7416a9e1764e39134fe9a%40thread.skype/Outlook%2520ahk?groupId=640b2f00-7b35-41b2-9e32-5ce9f5fcbd01&tenantId=8d4b558f-7b2e-40ba-ad1f-e04d79e6265a"
	Case "NWS":
        sTeamLink := "https://teams.microsoft.com/l/channel/19%3a0a0182fb8fa5455f84c1ef45b10cbe52%40thread.skype/AHK%2520Main?groupId=640b2f00-7b35-41b2-9e32-5ce9f5fcbd01&tenantId=8d4b558f-7b2e-40ba-ad1f-e04d79e6265a"
	Default:
        sTeamLink = "https://teams.microsoft.com/l/team/19%3a12d90de31c6e44759ba622f50e3782fe%40thread.skype/conversations?groupId=640b2f00-7b35-41b2-9e32-5ce9f5fcbd01&tenantId=8d4b558f-7b2e-40ba-ad1f-e04d79e6265a"
	}
    Run, %sTeamLink%
} ; eofun

; ---------------------------------------------------------------------- 

PowerTools_OpenWiki(){
sUrl := "https://connext.conti.de/wikis/home/wiki/Wc4f94c47297c_42c8_878f_525fd907cb68/page/GUIDEs%20Power%20Tools"  
Run,  "%sUrl%"
}

; ---------------------------------------------------------------------- 

PowerTools_OpenDoc(key:=""){
RegRead, PT_DocRootUrl, HKEY_CURRENT_USER\Software\PowerTools, DocRootUrl
If (key ="") {
Switch PowerTools_Config
{
Case "Conti":  
    sUrl := "https://connext.conti.de/wikis/home/wiki/Wc4f94c47297c_42c8_878f_525fd907cb68/page/GUIDEs%20Power%20Tools"  
Default:
    sUrl := "https://github.com/tdalon/ahk"
}
} Else {
    If InStr(PT_DocRootUrl,".blogspot.") 
        key := StrReplace(key,"_","-")
    sUrl = %PT_DocRootUrl%/%key% 
}
Run,  "%sUrl%"
} ; eofun



; ----------------------------------------------------------------------
PowerTools_GetSetting(SettingName){
RegRead, Setting, HKEY_CURRENT_USER\Software\PowerTools, %SettingName%
If (Setting=""){
    Setting := PowerTools_SetSetting(SettingName)
}
return Setting
} ; eofun
; ----------------------------------------------------------------------
PowerTools_SetSetting(SettingName){
; for call from Menu with Name Set <Setting>
SettingName := RegExReplace(SettingName,"^Set ","") 
SettingProp := RegExReplace(SettingName," ","") ; Remove spaces 

RegRead, Setting, HKEY_CURRENT_USER\Software\PowerTools, %SettingProp%
InputBox, Setting, PowerTools Setting, Enter %SettingName%,, 250, 125
If ErrorLevel
    return
PowerTools_RegWrite(SettingProp,Setting)
return Setting
} ; eofun

; ----------------------------------------------------------------------
PowerTools_GetConfig(){
RegRead, Config, HKEY_CURRENT_USER\Software\PowerTools, Config
If (Config=""){
    Config := PowerTools_LoadConfig()
}
return Config
}
; ----------------------------------------------------------------------
PowerTools_SetConfig(){
RegRead, Config, HKEY_CURRENT_USER\Software\PowerTools, Config
DefListConfig := "Default|Conti|Vitesco|Ini"
Select := 0
Loop, parse, DefListConfig, | 
{
    If (A_LoopField = Config) {
        Select := A_Index
        break
    }
}
Config := ListBox("PowerTools Config","Select your configuration:",DefListConfig,Select)
If (Config="")
    return
PowerTools_RegWrite("Config",Config)
return Config
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
PowerTools_RegRead(Prop){
RegRead, OutputVar, HKEY_CURRENT_USER\Software\PowerTools, %Prop%
return OutputVar
}

; -------------------------------------------------------------------------------------------------------------------
PowerTools_RegWrite(Prop, Value){
RegWrite, REG_SZ, HKEY_CURRENT_USER\Software\PowerTools, %Prop%, %Value%    
}

; -------------------------------------------------------------------------------------------------------------------
PowerTools_RegGet(Prop,sPrompt :=""){
RegRead, Value, HKEY_CURRENT_USER\Software\PowerTools, %Prop%
If (Value = "") { ; empty 
    If (sPrompt = "")
        sPrompt = Enter value for %Prop%:
    InputBox, Value, %Prop%, %sPrompt%, , 200, 150, , , , , 
    If ErrorLevel
	    return
    PowerTools_RegWrite(Prop,Value)
}
return Value
} ;eofun

; ----------------------------------------------------------------------
PowerTools_LoadConfig(Config :=""){
If (Config=""){
    Config := PowerTools_SetConfig()
}

IniFile = %A_ScriptDir%\PowerTools.ini
Switch Config
{
    Case "Conti":
        PowerTools_RegWrite("Domain","contiwan.com")
        IniWrite, contiwan.com, %IniFile%, Main, Domain
        PowerTools_RegWrite("TenantName","continental")
        IniWrite, continental, %IniFile%, Main, TenantName

        PowerTools_RegWrite("ProxyServer","http://ep.threatpulse.net:80")
        IniWrite, http://ep.threatpulse.net:80, %IniFile%, Main, ProxyServer

        PowerTools_RegWrite("ConnectionsRootUrl","connext.conti.de")
        IniWrite, connext.conti.de, %IniFile%, Connections, ConnectionsRootUrl       
        PowerTools_RegWrite("ConnectionsName","ConNext")
        IniWrite, ConNext, %IniFile%, Connections, ConnectionsName

        PowerTools_RegWrite("TeamsOnly",1)
        IniWrite, 1, %IniFile%, MicrosoftTeams, TeamsOnly

        PowerTools_RegWrite("DocRootUrl","https://connext.conti.de/blogs/tdalon/entry/")
        IniWrite, https://connext.conti.de/blogs/tdalon/entry/, %IniFile%, Main, DocRootUrl


    Case "Vitesco":
        ;IniWrite, vit.com, %IniFile%, Main, Domain
        PowerTools_RegWrite("ConnectionsName","inVite")
        IniWrite, inVite, %IniFile%, Connections, ConnectionsName
        PowerTools_RegWrite("ConnectionsRootUrl","invite.vitesco-technologies.net")
        IniWrite, invite.vitesco-technologies.net, %IniFile%, Connections, ConnectionsRootUrl

    Case "Default":
        sEmpty=
        PowerTools_RegWrite("Domain","")
        IniWrite,%sEmpty% , %IniFile%, Main, Domain

        PowerTools_RegWrite("TenantName","")
        IniWrite, %sEmpty%, %IniFile%, Main, TenantName

        PowerTools_RegWrite("ProxyServer","n/a")
        IniWrite, n/a, %IniFile%, Main, ProxyServer

        PowerTools_RegWrite("ConnectionsRootUrl","")
        IniWrite, %sEmpty%, %IniFile%, Connections, ConnectionsRootUrl

        PowerTools_RegWrite("TeamsOnly",1)
        IniWrite, 1, %IniFile%, MicrosoftTeams, TeamsOnly

        PowerTools_RegWrite("DocRootUrl","https://tdalon.blogspot.com/")
        IniWrite, https://tdalon.blogspot.com/, %IniFile%, Main, DocRootUrl

    Case "Ini":
        If FileExist(IniFile) {
            IniRead, IniVal, PowerTools.ini, Main, Domain
            PowerTools_RegWrite("Domain",IniVal)
            IniRead, IniVal, PowerTools.ini, MicrosoftTeams, TeamsOnly
            PowerTools_RegWrite("TeamsOnly",IniVal)
            IniRead, IniVal, PowerTools.ini, Main, DocRootUrl
            PowerTools_RegWrite("DocRootUrl",IniVal)
            IniRead, IniVal, PowerTools.ini, Connections, ConnectionsRootUrl
            If (IniVal != "ERROR")
                PowerTools_RegWrite("ConnectionsRootUrl",IniVal)
            Else
                PowerTools_RegWrite("ConnectionsRootUrl","")
        } Else {
            MSgBox 0x10, PowerTools: Error, PowerTools.ini can not be found!
            return
        }
} ; end switch

} ; eofun

; ----------------------------------------------------------------------

PowerTools_NWSSearch(){
If GetKeyState("Ctrl") {
    Run, "" ; TODO
    return
}
sSelection := GetSelection()
If (!sSelection) {
	If (A_UserName = "uid41890") { ; only for me
		sUrl = file:///C:/Users/uid41890/Documents/GitHub/tools/connext/nws_search/nws_search.html
	} Else {
		sUrl = http://github.conti.de/pages/uid41890/tools/nws_search.html
	}	
} Else { ; run with query does not work with local file because google escapes ? to %3f
	sUrl = http://github.conti.de/pages/uid41890/tools/nws_search.html
	sUrl = %sUrl%?q=%sSelection%
}
Run, %sUrl%
}

; ----------------------------------------------------------------------
PowerTools_CursorHighlighter(){
CHFile= %A_ScriptDir%\Cursor Highlighter
If a_iscompiled
	CHFile = %CHFile%.exe
Else
	CHFile = %CHFile%.ahk

If !FileExist(CHFile) { ; download if it doesn't exist
    return
}
Run, %CHFile%
}
; -------------------------------------------------------------------------------------------------------------------

PowerTools_MenuTray(){
; SubMenuSettings := PowerTools_MenuTray()
Menu, Tray, NoStandard
Menu, Tray, Add, &Help, MenuCb_PTHelp
Config := PowerTools_GetConfig()
If (Config="Conti")
    Menu,Tray,Add,Support (Teams Channel), MenuCb_PowerTools_Support
Menu, Tray, Add, Tweet for support, MenuCb_PowerTools_Tweet
Menu, Tray, Add, Check for update, MenuCb_PTCheckForUpdate
Menu, Tray, Add, Changelog, MenuCb_PTChangelog
Menu, Tray, Add, News, MenuCb_PTNews


If !a_iscompiled {
	IcoFile := RegExReplace(A_ScriptFullPath,"\..*",".ico")
	If (FileExist(IcoFile)) 
		Menu,Tray,Icon, %IcoFile%
}

If (A_ScriptName = "Teamsy.exe") or (A_ScriptName = "Teamsy.ahk")
    return

; -------------------------------------------------------------------------------------------------------------------
; SETTINGS
Menu, SubMenuSettings, Add, Launch on Startup, MenuCb_ToggleSettingLaunchOnStartup
SettingLaunchOnStartup := ToStartup(A_ScriptFullPath)
If (SettingLaunchOnStartup) 
  Menu,SubMenuSettings,Check, Launch on Startup
Else 
  Menu,SubMenuSettings,UnCheck, Launch on Startup

Menu, Tray, Add, Settings, :SubMenuSettings

Menu,Tray,Add
Menu,Tray,Standard
Menu,Tray,Default,&Help

return SubMenuSettings
}

; ---------------------------------------------------------------------- STARTUP -------------------------------------------------
MenuCb_ToggleSettingLaunchOnStartup(ItemName, ItemPos, MenuName){
SettingLaunchOnStartup := !ToStartup(A_ScriptFullPath)
If (SettingLaunchOnStartup) {
 	Menu,%MenuName%,Check, %ItemName%	 
	ToStartup(A_ScriptFullPath,True)
}
Else {
    Menu,%MenuName%,UnCheck, %ItemName%	 
	ToStartup(A_ScriptFullPath,False)
}
}

MenuCb_PTHelp(ItemName, ItemPos, MenuName){
ScriptName := RegExReplace(A_ScriptName,"\..*","")
PowerTools_Help(ScriptName)
}

MenuCb_PTChangelog(ItemName, ItemPos, MenuName){
ScriptName := RegExReplace(A_ScriptName,"\..*","") 
PowerTools_Changelog(ScriptName)   
}

MenuCb_PTNews(ItemName, ItemPos, MenuName){
ScriptName := RegExReplace(A_ScriptName,"\..*","") 
PowerTools_News(ScriptName)   
}

MenuCb_PowerTools_Support(ItemName, ItemPos, MenuName){
ScriptName := RegExReplace(A_ScriptName,"\..*","")
PowerTools_Support(ScriptName)    
}

MenuCb_PowerTools_Tweet(ItemName, ItemPos, MenuName){
ScriptName := RegExReplace(A_ScriptName,"\..*","")
PowerTools_TweetMe(ScriptName)    
}

MenuCb_PTCheckForUpdate(ItemName, ItemPos, MenuName){
PowerTools_CheckForUptate()    
}

; -------------------------------------------------------------------------------------------------------------------

PowerTools_TweetPush(ScriptName){

sLogUrl := PowerTools_Changelog(ScriptName,False)
;sToolUrl := PowerTools_Help(ScriptName,False)

If (ScriptName ="NWS")
    ScriptName = NWSPowerTool

sText = New version of #%ScriptName%. See changelog %sLogUrl%

sUrl:= uriEncode(sUrl)
sText := uriEncode(sText)
sTweetUrl = https://twitter.com/intent/tweet?text=%sText%  ;&hashtags=%ScriptName%&url=%sToolUrl%
Run, %sTweetUrl%

} ;eofun

; -------------------------------------------------------------------------------------------------------------------

PowerTools_TweetMe(ScriptName){

sLogUrl := PowerTools_Changelog(ScriptName,False)
;sToolUrl := PowerTools_Help(ScriptName,False)

If (ScriptName ="NWS")
    ScriptName = NWSPowerTool

sText = @tdalon

sText := uriEncode(sText)
sTweetUrl = https://twitter.com/intent/tweet?text=%sText%&hashtags=%ScriptName%
Run, %sTweetUrl%

} ;eofun


; -------------------------------------------------------------------------------------------------------------------

PowerTools_GetParam(Param) {
If FileExist("PowerTools.ini") {
	IniRead, ParamVal, PowerTools.ini, Parameters, %Param%
	If !(ParamVal="ERROR")
		return ParamVal
}

Switch Param
{
    Case "PasteDelay":
        return 500
    Case "TeamsMentionDelay":
        return 1300
}

} ;eofun
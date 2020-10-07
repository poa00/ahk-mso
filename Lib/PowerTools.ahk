; PowerTools Lib

#Include <AHK>
AppList = ConNextEnhancer,MO,NWS,OutlookShortcuts,PeopleConnector,TeamsShortcuts

PTCheckForUpdate(ToolName :="") {
If !a_iscompiled {
	Run, https://github.com/tdalon/ahk ; no direct link because of Lib dependencies
    return
} 

If !ToolName    
    ScriptName := A_ScriptName
Else
    ScriptName = %ToolName%.exe
    ; Overwrites by default
sUrl = https://raw.githubusercontent.com/tdalon/ahk/master/%ScriptName%

ExeFile = %A_ScriptDir%\%ScriptName%
If Not FileExist(ExeFile) {
    UrlDownloadToFile, %sUrl%, %ScriptName%
    return
}

UrlDownloadToFile, %sUrl%, %ScriptName%.github
guExe = %A_ScriptDir%\github_updater.exe
If Not FileExist(guExe)
    UrlDownloadToFile, https://raw.githubusercontent.com/tdalon/ahk/master/github_updater.exe, %guExe%
sCmd = %guExe% %ScriptName%
RunWait, %sCmd%,,Hide
} ; eof

; ---------------------------------------------------------------------- 
PowerTools_Help(ScriptName){
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
Case "NWS":
    sUrl := "https://tdalon.github.io/ahk/NWS%20PowerTool"
Case "Bundler":
    sUrl :="https://tdalon.github.io/ahk/PowerTools%20Bundler"
Case "all":
Default:
    sUrl := "https://tdalon.github.io/ahk/PowerTools"	
}
Run, %sUrl%
}

PowerTools_Changelog(ScriptName){
Switch ScriptName 
{
Case "ConnectionsEnhancer":
    sFileName = Connections-Enhancer-(Changelog)
Case "TeamsShortcuts":
    sFileName = Teams-Shortcuts-(Changelog)
Case "MO":
    sUrl := "http://github.conti.de/ContiSource/ahk/wiki/MO-(Release-Notes)"
    Run, %sUrl%
    return
Case "PeopleConnector":
    sFileName = People-Connector-(Changelog)
Case "NWS":
    sFileName = NWS-PowerTool-(Changelog)
Case "Bundler":
    ; TODO
    sFileName = PowerTools-Bundler-(Changelog)
Case "OutlookShortcuts":
    sFileName = Outlook-Shortcuts-(Changelog)
Case "Teamsy":
    sFileName = Teamsy-(Changelog)
Case "all":
Default:
    sFileName =  PowerTools-Release-Notes
}

If !A_IsCompiled {
    sFile = %A_ScriptDir%\docs\%sFileName%.md
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
}


; ----------------------------------------------------------------------
PowerTools_GetConfig(){
RegRead, Config, HKEY_CURRENT_USER\Software\PowerTools, Config
If (Config=""){
    Config := PowerTools_SetConfig()
}
return Config
}
; ----------------------------------------------------------------------
PowerTools_SetConfig(){
RegRead, Config, HKEY_CURRENT_USER\Software\PowerTools, Config
DefListConfig := "Conti|Vitesco|Public"
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
    Config := PowerTools_GetConfig()
}
Switch Config
{
    Case "Conti":
        PowerTools_RegWrite("Domain","contiwan.com")
        PowerTools_RegWrite("ConnectionsRootUrl","connext.conti.de")
        PowerTools_RegWrite("ConnectionsName","ConNext")
        PowerTools_RegWrite("TeamsOnly",1)
        PowerTools_RegWrite("DocRootUrl","https://connext.conti.de/blogs/tdalon/entry/")
    Case "Vitesco":
         PowerTools_RegWrite("ConnectionsName","InVite")
    Case "Public":
        PowerTools_RegWrite("Domain","")
        PowerTools_RegWrite("ConnectionsRootUrl","")
        PowerTools_RegWrite("TeamsOnly",0)
        PowerTools_RegWrite("DocRootUrl","https://tdalon.blogspot.com/")
    Case "Ini":
        If FileExist("PowerTools.ini") {
            IniRead, IniVal, PowerTools.ini, Main, Domain
            PowerTools_RegWrite("Domain",IniVal)
            IniRead, IniVal, PowerTools.ini, Main, TeamsOnly
            PowerTools_RegWrite("TeamsOnly",IniVal)
            IniRead, IniVal, PowerTools.ini, Main, DocRootUrl
            PowerTools_RegWrite("DocRootUrl",IniVal)
            IniRead, IniVal, PowerTools.ini, Connections, ConnectionsRootUrl
            If (IniVal != "ERROR")
                PowerTools_RegWrite("ConnectionsRootUrl",IniVal)
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

; -------------------------------------------------------------------------------------------------------------------

PowerTools_MenuTray(){
; SubMenuSettings := PowerTools_MenuTray()
Menu,Tray,NoStandard
Menu,Tray,Add, &Help, MenuCb_PTHelp
Menu,Tray,Add,Support (Teams Channel), MenuCb_PowerTools_Support
Menu,Tray,Add,Check for update, MenuCb_PTCheckForUpdate
Menu,Tray,Add,Changelog, MenuCb_PTChangelog

If !a_iscompiled {
	IcoFile := RegExReplace(A_ScriptFullPath,"\..*",".ico")
	If (FileExist(IcoFile)) 
		Menu,Tray,Icon, %IcoFile%
}

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

MenuCb_PowerTools_Support(ItemName, ItemPos, MenuName){
ScriptName := RegExReplace(A_ScriptName,"\..*","")
PowerTools_Support(ScriptName)    
}

MenuCb_PTCheckForUpdate(ItemName, ItemPos, MenuName){
PTCheckForUpdate()    
}
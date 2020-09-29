; Calls Lib/ToStartup
; Uses Registry Key HKEY_CURRENT_USER\Software\PowerTools 
;     HKEY_CURRENT_USER\Software\PowerTools, NotificationAtStartup
;     for Hotkey setting
#SingleInstance force
#NoEnv ; Avoids checking empty variables to see if they are environment variables (recommended for all new scripts).

#Include <PowerTools>
#Include <AHK>
; Calls Lib/HotkeyGUI
LastCompiled = 20200806155131

SubMenuSettings := PowerTools_MenuTray()
Menu,Tray,Insert,Settings,Power Tools Bundler, PowerTools_RunBundler

; Tooltip
If a_iscompiled {
  ;Menu,Tray,Add,Exit,AHKExit
  LastMod := LastCompiled
} Else {
  FileGetTime, LastMod , %A_ScriptFullPath%
}
FormatTime LastMod, %LastMod% D1 R
sTooltip = MO PowerTool %LastMod%.`nDouble-Click to Chat with MO.`nRight-Click to access additional functionalities.
Menu, Tray, Tip, %sTooltip%

; -------------------------------------------------------------------------------------------------------------------
; SETTINGS
Menu, SubMenuSettings, Add, MO Hotkey, MoHotkeySet
Menu, SubMenuSettings, Add, Bundler Hotkey, BundlerHotkeySet

; -------------------------------------------------------------------------------------------------------------------
; Links
Menu, SubMenuLinks, Add, NWS Search, PowerTools_NWSSearch
Menu, SubMenuLinks, Add, Power Tools (Wiki), PowerTools_OpenWiki

; -------------------------------------------------------------------------------------------------------------------
; Add Custom Menus to MenuTray
Menu,Tray,NoStandard
Menu, Tray, Add, Links, :SubMenuLinks
Menu,Tray,Add ; Separator
Menu,Tray,Standard


; -------------------------------------------------------------------------------------------------------------------

Menu, Tray, Insert, &Help, Chat with MO, RunMO
Menu,Tray,Default, Chat with MO


RegRead, MoHotkey, HKEY_CURRENT_USER\Software\PowerTools, MoHotkey
If (MoHotkey)
  Hotkey, %MoHotkey%, RunMO, On               ;Turn on the new hotkey.
RegRead, BundlerHotkey, HKEY_CURRENT_USER\Software\PowerTools, BundlerHotkey
If (BundlerHotkey)
  Hotkey, %BundlerHotkey%, PowerTools_RunBundler, On               ;Turn on the new hotkey.

return
; -------------------------------------------------------------------------------------------------------------------


; -------------------------------------------------------------------------------------------------------------------
MoHotkeySet:
; https://autohotkey.com/board/topic/47439-user-defined-dynamic-hotkeys/
RegRead, MoHotkey, HKEY_CURRENT_USER\Software\PowerTools, MoHotkey
HK := HotkeyGUI(,MoHotkey,,,"MO - Set Hotkey")
If ErrorLevel ; Cancelled
  return
If (HK = MoHotkey) 
  return
If (HK) {
  PowerTools_RegWrite("MoHotkey",HK)                     ;Save the hotkey for future reference.
  Hotkey, %HK%, RunMO, On               ;Turn on the new hotkey.
  TrayTip, Set MO Hotkey,% HK " Hotkey on"
} Else
  Hotkey, %MoHotkey%, RunMO, Off
return


; -------------------------------------------------------------------------------------------------------------------
BundlerHotkeySet:
; https://autohotkey.com/board/topic/47439-user-defined-dynamic-hotkeys/
RegRead, BundlerHotkey, HKEY_CURRENT_USER\Software\PowerTools, BundlerHotkey
HK := HotkeyGUI(,BundlerHotkey,,,"PowerTools Bundler - Set Hotkey")
If ErrorLevel ; Cancelled
  return
If (HK = BundlerHotkey) 
  return
If (HK) {
  PowerTools_RegWrite("BundlerHotkey",HK)                   ;Save the hotkey for future reference.
  Hotkey, %HK%, PowerTools_RunBundler, On               ;Turn on the new hotkey.
  TrayTip, Set Bundler Hotkey,% HK " Hotkey on"
} Else
  Hotkey, %BundlerHotkey%, PowerTools_RunBundler, Off

return


;This label may contain any commands for its hotkey to perform.
RunMO:
Run, https://chatbot.conti.de:88/
return

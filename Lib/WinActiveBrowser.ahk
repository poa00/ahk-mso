
GroupAdd, Browser, ahk_exe iexplore.exe
GroupAdd, Browser, ahk_exe chrome.exe
GroupAdd, Browser, ahk_exe firefox.exe
GroupAdd, Browser, ahk_exe palemoon.exe
GroupAdd, Browser, ahk_exe waterfox.exe
GroupAdd, Browser, ahk_exe vivaldi.exe
GroupAdd, Browser, ahk_exe msedge.exe	
GroupAdd, Browser, ahk_exe ApplicationFrameHost.exe ; Edge or IE with Win10

WinActiveBrowser(){
return WinActive("ahk_group Browser")
}
; ------------------------------------------------
; https://www.autohotkey.com/boards/viewtopic.php?p=353075

Menu_Show(hMenu, MenuLoop:=0, X:=0, Y:=0, Flags:=0) {            ; Ver 0.61 by SKAN on D39F/D39G
Local                                                           ;            @ tiny.cc/showmenu
  If (hMenu="WM_ENTERMENULOOP")
    Return True
  Fn := Func("ShowMenu").Bind("WM_ENTERMENULOOP"), n := MenuLoop=0 ? 0 : OnMessage(0x211,Fn,-1)
  DllCall("SetForegroundWindow","Ptr",A_ScriptHwnd)
  R := DllCall("TrackPopupMenu", "Ptr",hMenu, "Int",Flags, "Int",X, "Int",Y, "Int",0
             , "Ptr",A_ScriptHwnd, "Ptr",0, "UInt"),                     OnMessage(0x211,Fn, 0)
  DllCall("PostMessage", "Ptr",A_ScriptHwnd, "Int",0, "Ptr",0, "Ptr",0)
Return R
}

Menu_TrayParams() {      ; Original function is TaskbarEdge() by SKAN @ tiny.cc/taskbaredge
Local    ; This modified version to be passed as parameter to Menu_Show() @ tiny.cc/showmenu
  VarSetCapacity(var,84,0), v:=&var,   DllCall("GetCursorPos","Ptr",v+76)
  X:=NumGet(v+76,"Int"), Y:=NumGet(v+80,"Int"),  NumPut(40,v+0,"Int64")
  hMonitor := DllCall("MonitorFromPoint", "Int64",NumGet(v+76,"Int64"), "Int",0, "Ptr")
  DllCall("GetMonitorInfo", "Ptr",hMonitor, "Ptr",v)
  DllCall("GetWindowRect", "Ptr",WinExist("ahk_class Shell_SecondaryTrayWnd"), "Ptr",v+68)
  DllCall("SubtractRect", "Ptr",v+52, "Ptr",v+4, "Ptr",v+68)
  DllCall("GetWindowRect", "Ptr",WinExist("ahk_class Shell_TrayWnd"), "Ptr",v+36)
  DllCall("SubtractRect", "Ptr",v+20, "Ptr",v+52, "Ptr",v+36)
  Loop % (8, offset:=0)
    v%A_Index% := NumGet(v+0, offset+=4, "Int")
Return ( v3>v7 ? [v7, Y, 0x18] : v4>v8 ? [X, v8, 0x24]
       : v5>v1 ? [v5, Y, 0x10] : v6>v2 ? [X, v6, 0x04] : [0,0,0] )
}

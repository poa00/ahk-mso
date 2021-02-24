---
title: "Microsoft Teams: Raise your Hand "
redirect_from:
  - /teams-raise-hand
excerpt: "How to raise your hand in a Microsoft Teams meeting, the powertool way, using Teamsy and Teams Shortcuts PowerTools."
categories:
  - Microsoft Teams
tags:
  - teamsy
  - teams-shortcuts
header:
  image: /assets/images/teams_raise_hand.png
  og_image: /assets/images/teams_raise_hand.png
  image_description: "Microsoft Teams: Raise your Hand with PowerTools Teamsy and Teams Shortcuts"
---

# Screencast

{% include video id="Ysytg8_lr74" provider="youtube" %}

# AutoHotkey code

```AutoHotkey
Teams_RaiseHand() {
; Toggle Raise Hand on/off 
WinId := Teams_GetMeetingWindow()
If !WinId ; empty
    return
Tooltip("Teams Toggle Raise Hand...")
WinGet, curWinId, ID, A
WinActivate, ahk_id %WinId%
SendInput ^+k ; toggle video Ctl+Shift+k
WinActivate, ahk_id %curWinId%
} ; eofun
```

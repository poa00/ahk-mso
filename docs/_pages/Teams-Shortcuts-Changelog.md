---
permalink: /teams-shortcuts-changelog/
title: "Teams Shortcuts Changelog"
excerpt: "Release notes for Teams Shortcuts PowerTool."
---

[Teams Shortcuts](Teams-Shortcuts) Changelog

* 2021-02-22
	- add [global hotkey](https://tdalon.github.io/ahk/teams-global-hotkeys) for [Raise Your Hand](https://tdalon.blogspot.com/2021/02/teams-raise-hand-hotkey.html)
	- add [global hotkey](https://tdalon.github.io/ahk/teams-global-hotkeys) for [Mute App](https://tdalon.blogspot.com/2021/02/teams-mute-app.html)
* 2021-02-15
	- fix: [Get Meeting Window](https://tdalon.blogspot.com/2020/10/get-teams-window-ahk.html#getmeetingwindow) with , in title. Improved RegEx; only if Name, Firstname () format
* 2021-02-09
	- Bug fix. First start prompt for Connections Url
	- Remove Meeting->VLC Menus
	- Add global hotkey setting for Toggle Video
* 2021-02-08
	- [Fix](https://tdalon.blogspot.com/2021/02/ahk-tray-no-active-window.html). SysTray Actions will restore previous active window on Toggle mute and Video
* [2021-01-27](https://tdalon.blogspot.com/2021/02/teams-shortcuts-new-features-202101.html)
	- Added Setting to configure a Hotkey to toggle Mute
	- Added SysTray Mouse Click: Right Mouse Click -> Toggle video. Middle Mouse Click -> Toggle Mute
* 2021-01-26
	- Added menu in systray [Clear cache](https://tdalon.blogspot.com/2021/01/teams-clear-cache.html)
* 2021-01-18
	- New conversation: adjusted timing for expanding the compose box
* 2021-01-15
	- Extended keyword cal -> calendar
* 2021-01-14
	- Added Custom Background Setting for Library. Now it can be [set in ini file](https://tdalon.blogspot.com/2021/01/teams-custom-backgrounds.html#openlib).
* 2021-01-12
	- Change hotkey for new conversation from Win+N to Alt+N because of collision with OneNote hotkey
* 2020-12-08
	- Fixed: Personalize Mentions with Firstname with -
* 2020-12-04
  - Fixed Teams_GetMainWindow: collision of previous WinId. Added check for Window name as static variable
* 2020-11-24
  - SysTray Menu:
  	- add Tweet for support.
  	- Contact via Teams only for Conti Config.
  - [Smart Reply](https://tdalon.blogspot.com/2020/11/teams-shortcuts-smart-reply.html#getme): Add setting to store personal email/name so that Smart Reply also works without connection to Active Directory.
* 2020-11-23
	- [Smart Reply](https://tdalon.blogspot.com/2020/11/teams-shortcuts-smart-reply.html) improved version with Mention.
* 2020-11-20
	- [Send/Personalize Mentions](https://tdalon.blogspot.com/2020/11/teams-shortcuts-personalize-mentions.html): handling of case mentioned named can not be autocompleted (e.g. user not a member)
* 2020-11-17
  - Teams_GetMainWindow: improved version using Acc (no need to minimize all windows)
* 2020-11-16
	- [Personalize Mention](https://tdalon.blogspot.com/2020/11/teams-shortcuts-personalize-mentions.html) improved: Unified hotkeys for mention personalization (detect if () used by selecting/copying to clipboard)
* 2020-10-30
	- Fix new conversation with change of hotkey
	- Added: Paste Mentions (Extract Emails from Clipboard and convert Email to Name to be typed as mention in Teams
* 2020-10-28
  - Remove hotkey (Win+M; bad choice) for create a meeting. Rather launch from Teamsy.
  - Add SubMenu Meeting -> Cursor Highlighter
* 2020-09-21
    * [Smart Reply](https://tdalon.blogspot.com/2020/11/teams-shortcuts-smart-reply.html): Update to change of hotkey by Microsoft: new hotkey Alt+Shift+r instead of r for reply
* 2020-09-14
    * VLC Integration for Play
* 2020-07-27
    * Export Team List (used for NWS PowerTool IntelliPaste)
* 2020-07-24
    * Export Team Members List to Excel incl. Emails
* 2020-07-20
    * New conversation: will also focus on the subject line (Shift+Tab)
    * Changed win+r hotkey to alt+r (smart reply)
    * New meeting (alt+m)
* 2020-05-26
    * Add Win+P for /pop pop-out chat shortcut
    * Bulk add user will keep the PowerShell Window open and visible
* 2020-05-07
    * Custom Backgrounds: add open GUIDEs Backgrounds folder
* 2020-04-23
    * Add function "Open Second Instance" and "Open Web App"
* 2020-02-26
    * IntelliPaste Get Team Name by PowerShell
    * IntelliPaste: Paste Conversation: choose between Team|Channel|Message link
* 2020-02-21: add to favorites: prefill link text with Channel Name
* 2020-02-20: added [help](https://connext.conti.de/blogs/tdalon/entry/teams_shortcuts_ahk) to Teams Favorites
* 2020-02-13: Improved Smart Reply

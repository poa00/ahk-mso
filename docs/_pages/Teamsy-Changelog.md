---
permalink: /Teamsy-Changelog
title: "Teamsy Changelog"
excerpt: "Release notes for Teamsy and Teamsy Launcher PowerTools."
---

[Teamsy](Teamsy) Changelog

* 2021-03-04
  - Add keyword 'wn' for whatsnew
* 2021-03-01
  - [Teamsy Launcher](Teamsy-Launcher)
  - Added bg|background keyword to open background folder
* 2021-02-24
  - fix: Restart, ClearCache, CleanRestart: replace WinActive by WinExist and Process, Exist
* 2021-02-23
  - fix Teams_GetMeetingWindow and Teams_GetMainWindow if VirtuaWin is used. (Main Window might be hidden, on another virtual desktop)
* 2021-02-22
	- add [Raise Your Hand](https://tdalon.blogspot.com/2021/02/teams-raise-hand.html)
	- add [Mute App](https://tdalon.blogspot.com/2021/02/teams-mute-app.html)
* 2021-02-15
	- fix: [Get Meeting Window](https://tdalon.blogspot.com/2020/10/get-teams-window-ahk.html#getmeetingwindow) with , in title. Improved RegEx; only if Name, Firstname () format
	- t m: activates meeting window. t n m: to create a new meeting
	- improved WinListBox: double click in list box supported
* 2021-02-08
	- Video Toggle: restore last window focus.
* 2021-01-29
	- Teamsy can be run at startup.
	- Added Icon with Help menu.
* 2021-01-27
	- Teams Mute: restore previously active window
* 2021-01-26
	- Added [Clear cache and Clean restart](https://tdalon.blogspot.com/2021/01/teams-clear-cache.html)
* 2021-01-22
	- Revert TeamsFavs: replace https to msteams to open directly in app (better multiple window handling.)
* 2021-01-21
	- Teams_GetMeetingWindow: Exclude Call in progress windows
	- Share: put back the meeting window if on the front and multiple monitors
* 2021-01-18
	- New conversation: adjusted timing for expanding the compose box
	- Added keyword for news
* 2021-01-11
	- Added keyword to open help/ documentation.
* 2021-01-08
	- Bug fix: Update SmartReply (due to change in UI Ctrl+A behavior.)
* 2021-01-05
	- Fix: Teams_GetMeetingWindow, Teams_GetMainWindow regexp escape |
* 2020-12-04
  - Fixed Teams_GetMainWindow: collision of previous WinId. Added check for Window name as static variable
* 2020-11-17
	- Teams_GetMainWindow: improved version using Acc (no need to minimize all windows)
* 2020-11-09
	- allow quick message via @ (previously / was prepended)
* 2020-10-28
  - Fix: Teams restart
* 2020-10-22
  - Improved new conversation: will work if content pane wasn't selected / from navigation list pane (drawback: flashing of search box if in conversation area)
* 2020-10-15
  - Mute hotkey does not require meeting window to be active. Main client window is enough.
* 2020-10-13
  - If no meeting window founds, clean exit
  - Add q:quit and r: restart
  - GetMeetingWindow with WinListBox; Excludes 1-1 chat window = containing ",". Preselect last Meeting window. Support on-hold meetings.
* 2020-10-07
  - Fix: cal keyword. (missing return)
  - Changed f->find (previously set status to free)
* 2020-10-06
  - Fix: New expanded conversation hotkey [broken](https://tdalon.blogspot.com/teamsy-new-conversation)
* 2020-09-30
    * Integrate to PowerTools Bundle. Add SysTray with link to help/ changelog. Add to Bundler.
* 2020-09-29
    * add 'cal' keyword to open calendar

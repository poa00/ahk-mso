---
permalink: /Teams-Meeting-Reactions
title: "Teams Meeting Reactions Shortcuts"
excerpt: "How to setup Teams Meeting Reactions Shortcuts."
---

## Short Description

Teams Shortcuts PowerTool allows to send Teams Meeting reactions from a Launcher or Hotkey.
It is based on AutoHotkey [ImageSearch](https://www.autohotkey.com/docs/commands/ImageSearch.htm) functionality.
It may requires on the user site some fine-tuning like taking screenshots of some part of your Teams client and saving them at a specific location.

## Screencast

<p style="text-align: center;"><iframe width="560" height="315" src="https://www.youtube.com/embed/sPy07IzEGu4" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe></p>


## [Related Blog post](https://tdalon.blogspot.com/2021/03/teams-meeting-reactions-shortcuts.html)

## How to use

Select with your mouse elements that contains email information. It can be a list of addressees in an Outlook email or meeting, the email in the visit card in Teams, an page in the Browser (e.g. HCL Connections) etc.
Then double-tap the shift key to open the People Connector menu.

## How to setup/ troubleshooting

If when you run the feature, you get an error notification that the image file does not exist, you need to download the **img** subdirectory along with the PowerTools directory.

If when you run the feature, you get an error notification that the image search failed, you shall take yourself the screenshot of the corresponding reaction and save it in the **img** subdirectory, overwriting the previous file.


## Code

This feature is implemented in [/ahk/Lib/Teams.ahk](https://github.com/tdalon/ahk/blob/master/Lib/Teams.ahk) -> Teams_React (function)

## Potential improvements

If the image search does not work in a robust way, or you have different screens setup, I might implement a loop to search over multiple files.
In the code I also have hard-coded in the ImageSearch a resolution variation factor of 2 (*2). We could maybe increase this factor or make it as configurable setting. (I have noticed without variation factor, the image search failed on my PC.)

I could also add configurable global hotkeys for such actions but I find it much easier to run from the launcher via a natural command or keyword.

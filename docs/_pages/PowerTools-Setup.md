---
permalink: /PowerTools-Setup
title: "PowerTools Setup"
excerpt: "How to Setup the PowerTools"
---

## Screencast

<p style="text-align: center;"><iframe width="560" height="315" src="https://www.youtube.com/embed/sN6NxN5cPmA" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe><br><a href="https://youtu.be/sN6NxN5cPmA">Direct Link to YouTube video</a></p>

## Deployment and Requirements

The PowerTools are provided as compiled standalone executables so you don't need to have AutoHotkey installed on your PC.

These don't require any Admin rights to run.

When running these .exe you might get some Windows warning that it is unsecure. Normally you can check these warnings off, directly from the warning message or see for example (here)[https://www.windowscentral.com/how-fix-app-has-been-blocked-your-protection-windows-10].)

Alternatively you could run the tools from the AutoHotkey source (simply clone the [repository]((https://github.com/tdalon/ahk/))).
But I don't ensure that the current version of the repository runs properly, because I am human and might have forgotten to push some files. (Contact me)[Contact] if you want to run from the AHK source. (Most of the users I know don't bother about the AutoHotkey source and providing the compiled version allows me _not_ to check 100% consistency between my local version and the repository and quick deployment of a version.)

## One-by-one Setup

N.B.: For feature based on ImageSearch (like [Teams Shortcuts](Teams-Shortcuts) [Meeting Live reactions](Teams-Meeting-Reactions), the *img* subdirectory is also required, therefore I recommend to always [download all](#download-all) and run what you want only.

Still you can download for each PowerTool the .exe file by going here: [repo root](https://github.com/tdalon/ahk/)->[PowerTools](https://github.com/tdalon/ahk/tree/master/PowerTools) and clicking on the .exe file you want and then click the Download button.

It is recommended to install it in a separate fresh directory because some auxiliary files will be downloaded or created on the same level (e.g. PowerBundler, git_updater, ini file...).

From each PowerTool you can access functions like Open help, Add to Startup, Check for update etc. from the tool System tray icon menu.

## Download All

<a href="https://downgit.github.io/#/home?url=https://github.com/tdalon/ahk/tree/master/PowerTools"><button class="btn"><i class="fa fa-download"></i> Download</button></a>

You can download the whole [PowerTools Directory](https://github.com/tdalon/ahk/raw/master/PowerTools/) [here](https://downgit.github.io/#/home?url=https://github.com/tdalon/ahk/tree/master/PowerTools),
including the subdirectory with images used for ImageSearch related features like the [Teams Meeting Reactions Shortcuts](Teams-Meeting-Reactions).

## Manage via PowerTools Bundler

If you want to manage multiple PowerTools together, the best way is to use the [PowerTools Bundler](PowerTools-Bundler).
You can download the standalone version by clicking directly [here](https://github.com/tdalon/ahk/raw/master/PowerTools/PowerToolsBundler.exe).

From there you can donwload selected PowerTools via the Actions-> Check for Update/ Download.

![PowerTools Bundler Check for Update/Download](/ahk/assets/images/powertools_bundler_checkforupdate.png)

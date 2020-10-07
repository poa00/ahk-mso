# NWS PowerTool Changelog

* 2020-10-06
    * IntelliPaste: [embed github file](https://tdalon.blogspot.com/blogger-embed-github-file) 
    * IntelliPaste: do not ask for icon for Microsoft Whiteboard
* 2020-09-18
    * Share to Teams integrated in Win+F1 Browser menu
* 2020-09-14
    * IntelliPaste: fix Teams message link
* 2020-09-09
    * IntelliPaste support new YouTube links starting with https://youtu.be/ (before only www.youtube.com/)
* 2020-08-07
    * Fix IntelliPaste link to Ms Teams Team
    * Change icon from 42 to power
* 2020-08-05
    * PowerTools.ini
    * Bug fix IntelliPaste set Custom Hotkey
    * IntelliPaste: for ConNext blog listbox/choice to display blog name
    * [ConNext Quick Search](https://connext.conti.de/blogs/tdalon/entry/connext_search_ahk): empty tag in forum search url
    * [ConNext Quick Search](https://connext.conti.de/blogs/tdalon/entry/connext_search_ahk): close previous window
* 2020-08-04
    * IntelliPaste: breadcrumb for Teams files (previously only for folders) - [listbox](https://tdalon.blogspot.com/ahk-listbox) option to select with or without breadcrumb
    * Option for domain: portable for Vitesco
* 2020-08-03
    * ConNext Quick Search: wiki. replace %20 by space for Def Search from url query
* 2020-07-23
    * IntelliPaste: add setting to ask for icons for ConNext entries or not
    * IntelliPaste: do not ask for icon if window title contains " | Microsoft Teams" (Teams open in browser)
* 2020-07-22
    * Create Ticket from Social Support also works from the Systray icon menu.
    * Intelli Paste blog post will keep blog title after post name
* 2020-07-21
    * IntelliPaste: support rich-text format for WorkFlowy
* 2020-07-14
    * bug fix: sync if .ini file not filled properly (#TBD still there)->proper error message
    * fixed intellipaste for wiki search link. Get wiki name
* 2020-07-03
    * IntelliPaste: ConNext search links (Forum/Blog/Wiki) display query info in link display text
    * IntelliPaste: SharePoint old links remove ?Web=1 (fix icon for links ending e.g. with .pptx?Web=1)
* 2020-06-25
    * fix: IntelliPaste Teams link to conversation
* 2020-06-18
    * File Explorer-> Open Selection (Ctrl+O) Support for multiple files: loop on selected files
* 2020-06-17
    * File Explorer -> Open in Browser (Ctrl+E) from Sync with File Explorer: bug in some cases due to clipboard empty [because files were already copied] and window title shortened -> replaced by Explorer Lib
* 2020-06-15
    * OneDrive Mapper Integration
* 2020-06-04
    * IntelliPaste Teams Link to folder will display a breadcrumb e.g. General > Folder 1 > Folder 2
* 2020-06-03
    * Fix IntelliPaste: no icon
* 2020-05-28
    * ConNext Quick Search Blog and Wiki: initialize search with query used in the url
* 2020-05-26
    * IntelliPaste: do not paste on Cancel link text
    * IntelliPaste: link to ConNext Blog comment and Wiki comment are supported
* 2020-05-18
    * QuickSearch (Win+F) - ConNext Search Forum: support options -o (openQuestions) -a (answeredQuestions)
    * Added QuickSearch to F1 Menu
* 2020-05-13
    * CapsLock + B: run Bing search engine on selection
    * Settings -> Phone Number (see https://connext.conti.de/blogs/tdalon/entry/connext2ticket_ahk)
* 2020-05-07
    * From SharePoint DocLib in Browser to Sync Location: fix regexp + create IniFile if it does not exist
* 2020-04-28
    * Create an IT Ticket e.g. from Social Support Forum
* 2020-04-27
    * Open File in Explorer: warning for o365 SharePoint (feature does not work)
* 2020-04-14
    * ODB Sync File Explorer: support Teams with name including icons (replaced by ??)
* 2020-04-03
    * FIX: IntelliPaste in ConNext comment: link is not cleaned-up
    * IntelliPaste: won't ask if icon is wanted if no icon are available
* 2020-03-25
    * SP/ODB: Ctrl+O for office files opened in Browser via Internet Explorer (bug Edge Chromium) - else open local file.
* 2020-03-03
    * FIX: IntelliPaste single link in Jira
* 2020-02-26
    * IntelliPaste Get Team Name by PowerShell
* 2020-02-21
    * Disable IntelliPaste Insert key for Freemind
    * Settings-> IntelliPaste Hotkey: hotkey can be configured by the user (default Insert key)
* 2020-02-19
    * FIX: Open from Synced location: in case mapping is not defined, dummy browser window was opened
    * Added CapsLock+Y: for YouTube search
* 2020-02-18: added link to release notes/ change log
* 2020-02-18: Win+F1 will open main menu with help integrated. (in work)
* 2020-02-14 GetActiveUrl() support Vivaldi browser. Used for Browser Ctrl+E, share page etc.

; -------------------------------------------------------------------------------------------------------------------
; Function OpenFile
; Called by: Main Explorer-> Ctrl+O
; Open SharePoint url with Internet Explorer (Chrome will else download the file)
OpenFile(sFile) { 
    If SharePoint_IsSPUrl(sFile){
        Run iexplore.exe "%sFile%"
    } Else {
        Run, Open "%sFile%"
    }
}
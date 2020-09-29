ie(A_Args[1])

ie(sUrl){
    sUrl:=StrReplace(sUrl,"ie:","")
    Run, iexplore %sUrl%
} ; eof
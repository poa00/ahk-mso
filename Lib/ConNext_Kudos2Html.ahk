ConNext_Kudos2Html(sTag)
; Syntax:
; 	sHtml := ConNext_Kudos2Html(sTag)
; Called by: ConNextEnhancer, TextExpander
{
If InStr(sTag,"thank_you")
{
   imgsrc := "https://connext.conti.de/files/form/anonymous/api/library/e4e04f6d-8d49-457b-82d2-2b3ab397aa87/document/d154efec-aab6-437f-bd48-4af3617959c8/media/kudos_thank_you.jpg"
}    
    
Else If InStr(sTag,"awesome")
    imgsrc := "https://connext.conti.de/files/form/anonymous/api/library/e4e04f6d-8d49-457b-82d2-2b3ab397aa87/document/ab6c8faf-511d-4e5e-bd9d-23298781e79a/media/kudos_awesome.jpg"
Else If InStr(sTag,"visionary")
    imgsrc := "https://connext.conti.de/files/form/anonymous/api/library/e4e04f6d-8d49-457b-82d2-2b3ab397aa87/document/ebc2fb39-e0ac-488c-a904-855189942325/media/kudos_visionary.jpg"
Else If InStr(sTag,"team_player")
    imgsrc := "https://connext.conti.de/files/form/anonymous/api/library/e4e04f6d-8d49-457b-82d2-2b3ab397aa87/document/db9b3e77-8425-42b9-a841-df9a468bcaec/media/kudos_team_player.jpg"
Else If InStr(sTag,"made_my_day")
    imgsrc := "https://connext.conti.de/files/form/anonymous/api/library/e4e04f6d-8d49-457b-82d2-2b3ab397aa87/document/9c78019e-5b18-4206-90e2-cd43028d61b7/media/kudos_made_my_day.jpg"
Else If InStr(sTag,"great_idea")
    imgsrc := "https://connext.conti.de/files/form/anonymous/api/library/e4e04f6d-8d49-457b-82d2-2b3ab397aa87/document/4913e1c1-6604-458a-8457-bc21e7fd5eeb/media/kudos_great_idea.jpg"

sLink := "https://connext.conti.de/search/web/search?scope=allconnections&constraint={%22type%22:%22category%22,%22values%22:[%22Tag/" . sTag . "%22]}"

sHtml = <a href="%sLink%"><img src="%imgsrc%" style="margin: 0px auto; display: block;"/></a><a href="%sLink%">#%sTag%</a>
return sHtml
}
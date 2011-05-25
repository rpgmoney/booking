<SCRIPT language="JavaScript1.2" id="Sothink Widgets:main.pgt" type="text/javascript">
<!--
stm_bm(["uueoehr",400,"","",0,"","",0,0,0,0,0,1,0,0,""],this);
stm_bp("p0",[0,4,0,0,3,4,0,0,100,"",-2,"",-2,90,0,0,"#FFFFFF","#063C07","",3,0,0,"#FFFFFF"]);
stm_ai("p0i0",[0,"","","",-1,-1,0,"","iframe","","","","",0,0,0,"","",0,0,0,0,1,"#063C07",0,"#FFFFCC",0,"","",3,3,0,0,"#fffff7","#FFFFFF","#FFFFFF","#990000","8pt 'Tahoma','Arial'","8pt 'Tahoma','Arial'",0,0]);

<%if session("auth")="M" then%>
stm_ai("p0i9",[6,1,"#aca899","",0,0,0]);
stm_aix("p0i4","p0i0",[0,"管理者資料控管","","",-1,-1,0,]);
stm_bp("p1",[1,4,0,0,2,3,6,0,100,"progid:DXImageTransform.Microsoft.Fade(overlap=.5,enabled=0,Duration=0.43)",-2,"",-2,67,2,3,"#999999","#063C07","",3,1,1,"#aca899"]);
stm_aix("p1i5","p0i0",[0,"填報人員資料維護&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"auth.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"表單定義維護&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"define.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"管理者資料審查&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"verify.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"電子郵件催件通知&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"announce.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"退件作廢原因查詢&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"cancel.asp","_self"]);

stm_ep();
<%end if%>

<%if session("auth")<>""  then%>
stm_ai("p0i9",[6,1,"#aca899","",0,0,0]);
stm_aix("p0i4","p0i0",[0,"填表作業","","",-1,-1,0,]);
stm_bp("p1",[1,4,0,0,2,3,6,0,100,"progid:DXImageTransform.Microsoft.Fade(overlap=.5,enabled=0,Duration=0.43)",-2,"",-2,67,2,3,"#999999","#063C07","",3,1,1,"#aca899"]);
stm_aix("p1i5","p0i0",[0,"01 - 合作研究&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form01add.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"02 - 產學媒合活動&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form02add.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"03 - 專業諮詢&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form03add.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"04 - 產學論壇 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form04add.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"05 - 成果展示&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form05add.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"06 - 教育訓練&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form06add.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"07 - 共同實驗室管理&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form07add.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"08 - 技術服務&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form08add.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"09 - 專利成果&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form09add.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"10 - 可交易技術&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form10add.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"11 - 技術移轉成果&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form11add.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"12 - 附加價值&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form12add.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"13 - 衍生投資&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form13add.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"14 - 企業發展新創事業&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form14add.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"15 - 學校發展新創事業&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form15add.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"表 01-15 均無資料&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"formnodata.asp","_self"]);

stm_ep();
<%end if%>

stm_ai("p0i9",[6,1,"#aca899","",0,0,0]);
stm_aix("p0i4","p0i0",[0,"案件查詢","","",-1,-1,0,]);
stm_bp("p1",[1,4,0,0,2,3,6,0,100,"progid:DXImageTransform.Microsoft.Fade(overlap=.5,enabled=0,Duration=0.43)",-2,"",-2,67,2,3,"#999999","#063C07","",3,1,1,"#aca899"]);
stm_aix("p1i5","p0i0",[0,"01 - 合作研究案件查詢&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form01.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"02 - 產學媒合活動案件查詢&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form02.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"03 - 專業諮詢案件查詢&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form03.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"04 - 產學論壇案件查詢&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form04.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"05 - 成果展示案件查詢&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form05.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"06 - 教育訓練案件查詢&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form06.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"07 - 共同實驗室管理案件查詢&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form07.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"08 - 技術服務案件查詢&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form08.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"09 - 專利成果案件查詢&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form09.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"10 - 可交易技術案件查詢&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form10.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"11 - 技術移轉成果案件查詢&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form11.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"12 - 附加價值案件查詢&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form12.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"13 - 衍生投資案件查詢&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form13.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"14 - 企業發展新創事業案件查詢&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form14.asp","_self"]);
stm_aix("p1i5","p0i0",[0,"15 - 學校發展新創事業案件查詢&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form15.asp","_self"]);

stm_ep();

<%if session("sid")<>""  then%>
stm_ai("p0i9",[6,1,"#aca899","",0,0,0]);
stm_aix("p0i12","p0i0",[0,"待校稿專區","","",-1,-1,0,"deal.asp","_self"]);

stm_ep();
<%end if%>

<%'if session("auth")="M" then%>

stm_ai("p0i9",[6,1,"#aca899","",0,0,0]);
stm_aix("p0i4","p0i0",[0,"匯出Excel","","",-1,-1,0,]);
stm_bp("p1",[1,4,0,0,2,3,6,0,100,"progid:DXImageTransform.Microsoft.Fade(overlap=.5,enabled=0,Duration=0.43)",-2,"",-2,67,2,3,"#999999","#063C07","",3,1,1,"#aca899"]);
stm_aix("p1i5","p0i0",[0,"01 - 合作研究案件匯出&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form01export.asp?formclassify=01","_self"]);
stm_aix("p1i5","p0i0",[0,"02 - 產學媒合活動案件匯出&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form01export.asp?formclassify=02","_self"]);
stm_aix("p1i5","p0i0",[0,"03 - 專業諮詢案件匯出&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form01export.asp?formclassify=03","_self"]);
stm_aix("p1i5","p0i0",[0,"04 - 產學論壇案件匯出&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form01export.asp?formclassify=04","_self"]);
stm_aix("p1i5","p0i0",[0,"05 - 成果展示案件匯出&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form01export.asp?formclassify=05","_self"]);
stm_aix("p1i5","p0i0",[0,"06 - 教育訓練案件匯出&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form01export.asp?formclassify=06","_self"]);
stm_aix("p1i5","p0i0",[0,"07 - 共同實驗室管理案件匯出&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form01export.asp?formclassify=07","_self"]);
stm_aix("p1i5","p0i0",[0,"08 - 技術服務案件匯出&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form01export.asp?formclassify=08","_self"]);
stm_aix("p1i5","p0i0",[0,"09 - 專利成果案件匯出&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form01export.asp?formclassify=09","_self"]);
stm_aix("p1i5","p0i0",[0,"10 - 可交易技術案件匯出&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form01export.asp?formclassify=10","_self"]);
stm_aix("p1i5","p0i0",[0,"11 - 技術移轉成果案件匯出&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form01export.asp?formclassify=11","_self"]);
stm_aix("p1i5","p0i0",[0,"12 - 附加價值案件匯出&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form01export.asp?formclassify=12","_self"]);
stm_aix("p1i5","p0i0",[0,"13 - 衍生投資案件匯出&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form01export.asp?formclassify=13","_self"]);
stm_aix("p1i5","p0i0",[0,"14 - 企業發展新創事業案件匯出&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form01export.asp?formclassify=14","_self"]);
stm_aix("p1i5","p0i0",[0,"15 - 學校發展新創事業案件匯出&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","","",-1,-1,0,"form01export.asp?formclassify=15","_self"]);

stm_ep();

<%'end if%>
stm_ai("p0i9",[6,1,"#aca899","",0,0,0]);
stm_aix("p0i12","p0i0",[0,"操作教學","","",-1,-1,0,"doc/「產學合作填報平台」操作教學.ppt","_new"]);

stm_ep();

stm_ai("p0i9",[6,1,"#aca899","",0,0,0]);
stm_aix("p0i12","p0i0",[0,"登出","","",-1,-1,0,"logout.asp","_self"]);
stm_ai("p0i9",[6,1,"#aca899","",0,0,0]);
stm_ep();

stm_em();

//-->
</SCRIPT>


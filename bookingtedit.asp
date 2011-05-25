<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<!-- #INCLUDE FILE="lib/function.asp" -->
<%
validate=trim(request("validate"))
nextrec=trim(request("nextrec"))
id = trim(request("id"))
category=trim(request("category")) '老師或小老師
bdate=trim(request("bdate"))
btime=trim(request("btime"))
item=trim(request("item"))
teachername=trim(request("teachername"))
timeflag=trim(request("timeflag"))
sid=trim(request("sid"))
name=trim(request("name"))
slevel=trim(request("slevel"))
grade=trim(request("grade"))
class1=trim(request("class1"))
department=trim(request("department"))
score=trim(request("score"))
orallevel=trim(request("orallevel"))
oralset=trim(request("oralset"))
topic=trim(request("topic"))
briefing=trim(request("briefing"))
ptime=trim(request("ptime"))
languagecode=trim(request("languagecode"))
yn=trim(request("yn"))
tid=trim(request("tid"))
scid=trim(request("scid"))
yms=trim(request("yms"))
'

today=year(date())-1911 & right("0" & month(date()),2) & right("0" & day(date()),2)

sender=ifnull(trim(request("sender")),"booked.asp?category=" & category)

rulemsg = "詳細資訊請參考預約規則&nbsp;<a href='#' title='英外語診斷輔導中心　預約規則' onclick=""window.open('showrule.asp','','fullscreen=1,scrollbars=1');"" ><img src='images/icon_question.gif' border='0'></a>"

set dic = server.CreateObject("scripting.dictionary")
dic.Add "1","日"
dic.Add "2","一"
dic.Add "3","二"
dic.Add "4","三"
dic.Add "5","四"
dic.Add "6","五"
dic.Add "7","六"

function dateformat(vdate)
	if vdate<>"" then
		if  Cstr(left(vdate,1)) ="9"  then
		dateformat=cStr(cint(left(vdate,2))+1911 ) & "/" & mid(vdate,3,2) & "/" & right(vdate,2)
		else
		dateformat=cStr(cint(left(vdate,3))+1911 ) & "/" & mid(vdate,4,2) & "/" & right(vdate,2)
		end if
	end if
	
end function

set rs = server.CreateObject("adodb.recordset")
set rs2 = server.CreateObject("adodb.recordset")

if validate="edit" then
	sql = "select * from boo_book_T_M where id='"&id&"' "
	'response.write sql
	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
    if not rs.EOF then
		if name<>"" then
			rs("name")=name
		end if
		if slevel<>"" then
			rs("slevel")=slevel
		end if
		if grade<>"" then
			rs("grade")=grade
		end if
		if class1<>"" then
			rs("class1")=class1
		end if
		if department<>"" then
			rs("department")=department
		end if
		if score<>"" then
			rs("score")=score
		end if
		if orallevel<>"" then
            rs("orallevel")=orallevel
        end if
		if oralset<>"" then
            rs("oralset")=oralset
        end if
		if topic<>"" then
            rs("topic")=topic
        end if
		if briefing<>"" then
            rs("briefing")=briefing
        end if
		rs("initdate") = date()
		rs("inituid") = session("sid")

		rs.Update
        if Err.Number=0 then 
			
			'response.redirect "orallevellist.asp"

		else
			showmessage= Err.Description
		end if

	else
		showmessage="找不到該筆資料。"
	end if

	rs.close
elseif validate = "AddS_d" then
	tid_d=trim(request("tid_d"))
	sid_d=trim(request("sid_d"))
	name_d=trim(request("name_d"))
	slevel_d=trim(request("slevel_d"))
	grade_d=trim(request("grade_d"))
	class1_d=trim(request("class1_d"))
	department_d=trim(request("department_d"))
	score_d=trim(request("score_d"))
	flag="true"
	'限定人數
	icnt=trim(request("icnt"))
	'response.write "icnt = " & icnt
	if item="診斷"  then
		'2人
		if icnt >=1 then
			flag="false"
			showmessage1 = "同一時段診斷療程限定二人。" & rulemsg
		else
			'2人時須調整為50分鐘
			if timeflag<>"A" then
				sql = "select * from boo_book_T_M where bdate='"&bdate&"' and btime='"&btime&"' and teachername='"&teachername&"' and id<>'"&id&"' and category='"&category&"' and yn='Y'"
				rs.Open sql,msconn,adOpenStatic,adLockReadonly
				if not rs.EOF then
					flag=false
					showmessage1 = "二人登記同時段時，請選擇50分鐘。因指定預約的老師同一時段的另一節，已被其它人預約，無約邀請其它人參與。" & rulemsg
				else
				sql="update boo_book_T_M set timeflag='A' where id='"&id&"'"
				msconn.Execute sql
				timeflag="A"
				end if
				rs.close
			end if

		end if
	elseif item="諮商" then
		'2人
		if icnt >=1 then
			flag="false"
			showmessage1 = "同一時段諮商療程限定二人"& rulemsg
		else
			'2人時須調整為50分鐘
			if timeflag<>"A" then
				sql = "select * from boo_book_T_M where bdate='"&bdate&"' and btime='"&btime&"' and teachername='"&teachername&"' and id<>'"&id&"' and category='"&category&"' and yn='Y'"
				rs.Open sql,msconn,adOpenStatic,adLockReadonly
				if not rs.EOF then
					flag="false"
					showmessage1 = "二人登記同時段時，請選擇50分鐘。因指定預約的老師同一時段的另一節，已被其它人預約，無約邀請其它人參與。" & rulemsg
				else
					sql="update boo_book_T_M set timeflag='A' where id='"&id&"'"
					msconn.Execute sql
					timeflag="A"
				end if
				rs.close
			end if
		end if
	elseif item="口語" then
		'3人
		if icnt >=3 then
			flag="false"
			showmessage1 = "同一時段口語療程限定三人。" & rulemsg
		end if
	elseif item = "簡報" then
		'4人
		if icnt >=4 then
			flag="false"
			showmessage1 = "同一時段簡報療程限定四人。" & rulemsg
		end if
	elseif item = "寫作" then
		'2人
		if icnt >=2 then
			flag="false"
			showmessage1 = "同一時段寫作療程限定二人。" & rulemsg
		end if
	elseif item = "閱讀" then
		'2人
		if icnt >=2 then
			flag="false"
			showmessage1 = "同一時段閱讀療程限定二人。" & rulemsg
		end if
	elseif item = "詩歌" then
		'4人
		if icnt >=4 then
			flag="false"
			showmessage1 = "同一時段詩歌療程限定二人。" & rulemsg
		end if
	else
		showmessage1 = "資料有誤，請洽資教中心。"
	end if


	'若為口語則要大專英檢成績同一等級
	
	if item="口語" then
		if cdbl(score) > 90 then
			tmpv="A"
		else
			tmpV="B"
		end if
		if cdbl(score_d) > 90 then
			tmpv2="A"
		else
			tmpV2="B"
		end if
		'response.write "score=" & score & "<br>"
		'response.write "score_d=" & score_d & "<br>"
		if tmpv <> tmpv2 then
			flag="false"
			showmessage1 = "無法邀請該學員無法一起參與。"
		end if
	end if
	if (flag = "true" ) then
		'每人一天相同療程限預約一次
		sql = "select * from boo_book_T_M where bdate='"&bdate&"' and sid='"&sid_d&"' and item='"&item&"' and yn='Y'"
			
		rs.Open sql,msconn,adOpenStatic,adLockOptimistic
		if rs.EOF then
			rs.AddNew
			id_d= getguid()
			if id_d<>"" then
				rs("id")=id_d
			end if
			if id<>"" then
				rs("pid")=id
			end if
			if bdate<>"" then
				rs("bdate")=bdate
			end if
			if btime<>"" then
				rs("btime")=btime
			end if
			if ptime<>"" then
				rs("ptime")=ptime 
			end if
			if timeflag<>"" then
				rs("timeflag")=timeflag
			end if
			if ptime<>"" then
				rs("ptime")=ptime
			end if
			if teachername<>"" then
				rs("teachername")=teachername
			end if
			if item<>"" then
				rs("item")=item
			end if
			if sid_d<>"" then
				rs("sid")=sid_d
			end if
			if name_d<>"" then
				rs("name")=name_d
			end if
			if slevel_d<>"" then
				rs("slevel")=slevel_d
			end if
			if grade_d<>"" then
				rs("grade")=grade_d
			end if
			if class1_d<>"" then
				rs("class1")=class1_d
			end if
			if department_d<>"" then
				rs("department")=department_d
			end if
			if score_d<>"" then
				rs("score")=score_d
			end if
			if orallevel<>"" then
				rs("orallevel")=orallevel
			end if
			if oralset<>"" then
				rs("oralset")=oralset
			end if
			if topic<>"" then
				rs("topic")=topic
			end if
			if briefing<>"" then
				rs("briefing")=briefing
			end if
			if category<>"" then
				rs("category")=category
			end if
			if languagecode<>"" then
				rs("languagecode")=languagecode
			end if
			if tid_d<>"" then
				rs("tid")=tid_d
			end if
			if scid<>"" then
				rs("scid")=scid
			end if
			if yms<>"" then
				rs("yms")=yms
			end if
			rs("yn") ="Y"
			rs("initdate") = date()
			rs("inituid") = session("sid")


			rs.Update
			if Err.Number<>0 then 
				showmessage= Err.Description
			else
				'更新預約次數
				a=UpdateItemTime(item,tid_d,"1")
				tid_d=""
				sid_d=""
				name_d=""
				slevel_d=""
				grade_d=""
				class1_d=""
				department_d=""
				score_d=""
			end if

		else
			showmessage1 = "每人一天相同療程限預約一次，請勿重覆預約。" & rulemsg
		end if

		rs.close
	end if 'if flag = true
elseif validate="delete" then
	sqlm="update boo_book_T_M set yn='N',canceldate= Convert(varchar(10),Getdate(),111)  where id='"&id&"'"
	msconn.Execute sqlm
	'更新預約次數
	a=UpdateItemTime(item,tid,"2")

	sqld="update boo_book_T_M set yn='N',canceldate= Convert(varchar(10),Getdate(),111)   where pid='"&id&"'"
	msconn.Execute sqld
	'更新預約次數
	Str_tids = trim(request("Str_tids"))
	'response.write "Str_tids = " & Str_tids
	'更新預約次數
	a=UpdateItemTime(item,Str_tids,"2")

	
	if Err.number=0 then
		yn="N"
    else
        showmessage= Err.Description
    end if
elseif validate="delete_item" then
	delete_item = trim(request("delete_item"))
	delete_item_tid = trim(request("delete_item_tid"))
	sqld="update boo_book_T_M set yn='N',canceldate= Convert(varchar(10),Getdate(),111)   where id='"&delete_item&"'"
	msconn.Execute sqld
	'更新預約次數
	a=UpdateItemTime(item,delete_item_tid,"2")
	if Err.number=0 then
		delete_item=""


    else
        showmessage= Err.Description
    end if


else
	
	sql = "select * from boo_book_T_M   where id='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
    if rs.EOF then
		response.redirect sender
	else
		bdate=trim(rs("bdate"))
		btime=trim(rs("btime"))
		ptime=trim(rs("ptime"))
		item=trim(rs("item"))
		teachername=trim(rs("teachername"))
		timeflag=trim(rs("timeflag"))
		sid=trim(rs("sid"))
		name=trim(rs("name"))
		slevel=trim(rs("slevel"))
		grade=trim(rs("grade"))
		class1=trim(rs("class1"))
		department=trim(rs("department"))
		score=ifnull(trim(rs("score")),0)
		orallevel=trim(rs("orallevel"))
		oralset=trim(rs("oralset"))
		topic=trim(rs("topic"))
		briefing=trim(rs("briefing"))
		yn=trim(rs("yn"))
		category=trim(rs("category"))
		languagecode=trim(rs("languagecode"))
		signin=trim(rs("signin"))
		tid=trim(rs("tid"))
		scid=trim(rs("scid"))
		yms=trim(rs("yms"))
		if scid<>"" then
			sql2 = "select * from boo_schedule   where scid='"&scid&"' "
			rs2.Open sql2,msconn,adOpenStatic,adLockReadonly
			if not rs2.EOF then
				deptgroup = trim(rs2("deptgroup"))
			end if
			rs2.close
			
			if InStr(deptgroup,"ELC") > 0   then
				Response.Write "<Script language=javascript>alert('您預約的項目需至『ELC英語學習中心』執行\n地點：正氣樓１樓E106\n預約當日請直接至該地點報到\n累計預約未到滿2次者，停權2個月。\n');</Script>"

				deptgroupStr ="您預約的項目需至『ELC英語學習中心』執行，地點：正氣樓１樓E106。預約當日請直接至該地點報到。<br>累計預約未到滿2次者，停權2個月。"

			elseif  InStr(deptgroup,"LDCC") > 0 then
				Response.Write "<Script language=javascript>alert('您預約的項目需至『LDCC英外語能力診斷輔導中心』執行\n地點：露德樓3樓G326\n預約當日請直接至該地點報到\n累計預約未到滿2次者，停權2個月。\n');</Script>"
				deptgroupStr = "您預約的項目需至『LDCC英外語能力診斷輔導中心』執行，地點：露德樓3樓G326。預約當日請直接至該地點報到。<br>※累計預約未到滿2次者，停權2個月。"
			end if
		end if
	end if
	rs.close

end if


'口語題目
StrSubject=""
if oralset <> "" then
	set rsLoad = server.CreateObject("adodb.recordset")
	sql ="select * from boo_orallevel where category='"&oralset&"'"
	rsLoad.Open sql,msconn,adOpenStatic,adLockReadonly
	
	while not rsLoad.EOF
		if topic=rsLoad("topic") then
			StrSubject=StrSubject&"<option value="""&rsLoad("topic")&""" selected>"&  rsLoad("topic")&"</option>"
		else
			StrSubject=StrSubject&"<option value="""&rsLoad("topic")&""" >"&  rsLoad("topic")&"</option>"
		end if 
		rsLoad.MoveNext 
	wend
	set rsLoad=nothing
else
	StrSubject="<option value="""" selected>- 無 -</option>"
end if	



if   yn="N" or yn="A" or  signin<>"" then
	btnState="disabled"
else
	btnState=""
end if
'過期資料除非管理者,否則無法變更
'if session("classify")="S" and  cdbl(bdate) =< cdbl(today) then
'' 修改為上課的前一天就不能取消, 2011/05/16, shihchi
if session("classify")="S" and  (cdbl(bdate)-1) =< cdbl(today) then
	btnState="disabled"

end if




%>
<HTML>
<HEAD>
<TITLE> 【LDCC英外語能力診斷輔導中心】 </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/stm31.js" type="text/javascript"></SCRIPT>
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>
<script language="javascript">

function check_input()
{
    var errmsg=""

	if (form1.item.value=="口語")
	{
		if (form1.orallevel.value=="")
			errmsg += "請選擇口語級數\n";
		if (form1.oralset.value=="")
			errmsg += "請選擇口語系列\n";
		if (form1.topic.value=="")
			errmsg += "請選擇口語題目\n";
		
	
	}else if (form1.item.value=="4")
	{
		if (form1.briefing.value=="")
			errmsg += "簡報題目不能為空白\n";
	}
	
	
	
    if (errmsg=="")
        return true;
    else
    {
        alert(errmsg);
        return false;
    }
}

function check_input_d()
{
    var errmsg=""
	
	if (AddStudent_Form.sid_d.value=="")
        errmsg += "一起進行的同學學號不能為空白\n";
    if (AddStudent_Form.name_d.value=="")
        errmsg += "一起進行的同學姓名不能為空白\n";
	
    if (errmsg=="")
        AddStudent_Form.submit();
    else
        alert(errmsg);
}
function ChkStudent_d()
{
	vWinCal2 = window.open("lib/checkstudent_d.asp?sid="+AddStudent_Form.sid_d.value + "&item="+ AddStudent_Form.item.value+ "&catgory="+ AddStudent_Form.category.value  ,"iframe_query", "width=200,height=50,status=no,resizable=no,scrollbars=no");
	vWinCal2.opener = AddStudent_Form;
}
function changesubject()
{
	vWinCal2 = window.open("lib/changesubject.asp?oralset="+form1.oralset.value  ,"iframe_query", "width=200,height=50,status=no,resizable=no,scrollbars=no");
	vWinCal2.opener = form1;
}



function copy_record()
{
    form1.validate.value="";
    form1.onsubmit="";
    form1.action="bookingt.asp";
    form1.submit();
}
function delete_record()
{
    if (confirm("您確定要取消整筆資料嗎?(取消的範圍包含預約者和一起參與之同學)"))
    {
        form1.validate.value="delete";
		form1.Str_tids.value=AddStudent_Form.StrSid_tid.value
        form1.submit();
    }
}

function delete_record_item(vid,vtid)
{
    if (confirm("您確定要取消該學嗎預約資料嗎?"))
    {
        AddStudent_Form.validate.value="delete_item";
		AddStudent_Form.delete_item.value=vid;
		AddStudent_Form.delete_item_tid.value=vtid;
        AddStudent_Form.submit();
    }
}
function ChkStudent()
{
	vWinCal2 = window.open("lib/checkstudent.asp?sid="+form1.sid.value+"&languagecode="+ form1.languagecode.value ,"iframe_query", "width=200,height=50,status=no,resizable=no,scrollbars=no");
	vWinCal2.opener = form1;
}
</script>
</HEAD>
<BODY bottomMargin=0 leftMargin=0 topMargin=0 rightMargin=0 marginheight="0" marginwidth="0"  bgcolor="#f4c60d">
<TABLE cellSpacing=1 cellPadding=0 width="760"  height="100%" align="center" >
<TR><TD>
<TABLE cellSpacing=0 cellPadding=0 width="760"  height="100%" align="center" bgColor=#ffffff border=0>
<TR height="70"><TD><img src="images\top.jpg" border="0"></TD></TR>
<TR height="25" bgcolor="#333333">
	<TD align="center">
	<TABLE cellSpacing=0 cellPadding=0 border=0 width="100%">
	<TR>
	<TD align="left"><!-- #INCLUDE FILE="lib\link.inc" --></TD>
	<TD align="right"><!-- #INCLUDE FILE="lib\promsg.inc" --></TD>
	</TR>
	</TABLE>
	</TD>
</TR>
<TR valign="top">
	<TD>
<!-- ---------------------------------------------------------------------------------------- -->
	<P><BR>
	<TABLE cellSpacing=3 cellPadding=3 border=0 width="100%">
	<TR >
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center"><%if category="T" then response.write "預約教師輔導療程(編輯)" else response.write "預約小老師輔導療程(編輯)" end if%></TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="bookingtedit.asp" onsubmit="return check_input();">
			<input type="hidden" value="edit" name="validate">
			<input type="hidden" value="<%=ptime%>" name="ptime">
			<input type="hidden" value="<%=id%>" name="id">
			<input type="hidden" value="<%=yn%>" name="yn">
			<input type="hidden" value="<%=category%>" name="category">
			<input type="hidden" value="<%=sender%>" name="sender">
			<input type="hidden" value="<%=languagecode%>" name="languagecode">
			<input type="hidden" value="<%=tid%>" name="tid">
			<input type="hidden"  name="Str_tids" id="Str_tids" value="<%=Str_tids%>">
			<input type="hidden"  name="scid" id="scid" value="<%=scid%>">
			<input type="hidden"  name="yms" id="yms" value="<%=yms%>">
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>學號：</TD>
						<TD>姓名：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=sid%>" maxlength="25" size="35" onblur="ChkStudent()"  name="sid" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						<TD>
						<input type="text" value="<%=name%>" maxlength="25" size="35" name="name" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly >
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>學制：</TD>
						<TD>系所：</TD>
						<TD>年級：</TD>
						<TD>班級：</TD>
						<TD>大專英檢成績：</TD>
						<TD></TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=slevel%>" maxlength="10"   name="slevel" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						<TD>
						<input type="text" value="<%=department%>" maxlength="10"  name="department" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						<TD>
						<input type="text" value="<%=grade%>" maxlength="10" size="10"  name="grade" class="inputtext"  style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						
						<TD>
						<input type="text" value="<%=class1%>" maxlength="25"  name="class1" class="inputtext"  style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						<TD>
						<input type="text" value="<%=score%>" maxlength="25"  name="score" class="inputtext"  style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD><%if category = "T" then response.write "老師" else response.write "小老師" end if%>：</TD><TD>&nbsp;語言別：&nbsp;</TD><TD>預約狀態：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=teachername%>" maxlength="25" size="35"  name="teachername" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						<TD align="center"><%=languagecode%></TD>
						<TD>&nbsp;&nbsp;<%=replace(replace(yn,"Y","<font color=""blue"">已預約</font>"),"N","<font color=""red"">取消</font>")%></TD>
					</TR>
				</TABLE>
			</TD></TR>
			
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>預約項目：&nbsp;&nbsp;</TD>
						<TD>預約日期：</TD>
						<TD>星期：</TD>
						<TD>預約時段：</TD>
						<TD></TD>
					</TR>
					<TR>
						<TD><%=item%><input type="hidden" value="<%=item%>" name="item"></TD>
						<TD>
						<input type="text" value="<%=bdate%>" maxlength="25" size="15" name="bdate" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						<TD>
						<%="（&nbsp;"&dic.Item(cstr(cint(weekday(dateformat(bdate)))))&"&nbsp;）"%>
						</TD>
						<TD>
							<input type="hidden" value="<%=btime%>" maxlength="25"  name="btime" class="inputtext" readonly>
							<select name="btime1" class="inputtext" disabled>
							<option value=""> - 請指定 -</option>
							<optgroup label="上午">
							<option value="1010" <%if btime="1010" then response.write "selected" end if%>>10:10～11:00</option>
							<option value="1110" <%if btime="1110" then response.write "selected" end if%>>11:10～12:00</option>
							</optgroup>
							<optgroup label="中午">
							<option value="1210" <%if btime="1210" then response.write "selected" end if%>>12:10～13:00</option>
							</optgroup>
							<optgroup label="下午">
							<option value="1310" <%if btime="1310" then response.write "selected" end if%>>13:10～14:00</option>
							<option value="1410" <%if btime="1410" then response.write "selected" end if%>>14:10～15:00</option>
							<option value="1510" <%if btime="1510" then response.write "selected" end if%>>15:10～16:00</option>
							<option value="1610" <%if btime="1610" then response.write "selected" end if%>>16:10～17:00</option>
							
							</optgroup>
							</select>
						</TD>
						<TD>
							<TABLE cellSpacing=0 cellPadding=0  border=0 >
							<TR>
							<TD>&nbsp;<%=replace(replace(replace(timeflag,"U","上一節(25分)"),"B","下一節(25分)"),"A","上下二節(50分)")%></TD>
							</TR>
							</TABLE>
						</TD>
						
					</TR>
				</TABLE>
			</TD></TR>
			<TR id="area_oral" style="DISPLAY:<%if item="口語"  then response.write "block" else response.write "none" end if %>"><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						
						<TD>口語系列：</TD>
						<TD>口語題目：</TD>
						<TD ></TD>
					</TR>
					<TR>
						

						<TD>
						<select name="oralset" class="inputtext" onchange=changesubject()>
						<option value=""> - 請指定口語系列 -</option>
						<%if cdbl(score) >90 and languagecode="E" then %>
						
						<option value="Issues in English I" <%if oralset="Issues in English I" then response.write "selected" end if%>>Issues in English I</option>
						<option value="Issues in English II" <%if oralset="Issues in English II" then response.write "selected" end if%>>Issues in English II</option>
						<%else%>
						<option value="Conversation Topics" <%if oralset="Conversation Topics" then response.write "selected" end if%>>Conversation Topics</option>
						
						<%end if%>
						<option value="My ET" <%if oralset="My ET" then response.write "selected" end if%>> My ET</option>
						</select>
						</TD>
						<TD>
						<select name="topic" class="inputtext" >
						<option value=""> - 請指定口語題目 -</option>
						<%=StrSubject%>
						</select>
						</TD>
						<TD>
						<%if cdbl(score) > 90 and languagecode="E" then %>
						<select name="orallevel" class="inputtext">
						<option value=""> - 請指定口語級數 -</option>
						<option value="Level 1" <%if orallevel="Level 1" then response.write "selected" end if%>>Level 1</option>
						<option value="Level 2" <%if orallevel="Level 2" then response.write "selected" end if%>>Level 2</option>
						<option value="Level 3" <%if orallevel="Level 3" then response.write "selected" end if%>>Level 3</option>
						<option value="Level 4" <%if orallevel="Level 4" then response.write "selected" end if%>>Level 4</option>
						</select>
						<%else%>
						<input type=hidden value=Level1 name=orallevel >
						<%end if%>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR id="area_briefing" style="DISPLAY:<%if item="簡報"  then response.write "block" else response.write "none" end if %>"><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>簡報題目：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=briefing%>" maxlength="100" size="55" name="briefing" class="inputtext" >
						</TD>
						
					</TR>
				</TABLE>
			</TD></TR>
			<TR>
			<TD>
			<BR>
			<input  type="submit" value="儲存"      class="inputbutton" >
			<input  type="button" value="取消預約"   <%=btnState%>   class="inputbutton" onclick="delete_record();">
			<input  type="button" value="返回" class="inputbutton" onclick="window.location='<%=sender%>'">
			</TD>
			</TR>
			<TR><TD><font color="blue"><%=deptgroupStr%></font></TD></TR>
			</form>
			</TABLE>
		</TD>
	</TR>
	<tr height="20"><TD></TD><td></td></tr>
	<tr height="7"><TD></TD><td background="images/lin04.gif"></td></tr>
	<TR>
		<TD></TD><TD>
		<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0 >
		<TR><TD><font color="blue">※此區塊可加入一起參與的同學</font><BR></TD></TR>

		<TR><TD class="errmsg"><%=showmessage1%></TD></TR>
		<TR>
		<TD width="100%" valign="Top">
			<TABLE class=normal cellSpacing=0 cellPadding=0 width="100%"  border=0>
			<TR><TD>
				<form id="AddStudent_Form" name="AddStudent_Form" method="post" action="bookingtedit.asp" >
				<input type="hidden" value="AddS_d" name="validate">
				<input type="hidden" value="<%=id%>" name="id">
				
				<input type="hidden" value="<%=sid%>" name="sid">
				<input type="hidden" value="<%=name%>" name="name">
				<input type="hidden" value="<%=slevel%>" name="slevel">
				<input type="hidden" value="<%=department%>" name="department">
				<input type="hidden" value="<%=grade%>" name="grade">
				<input type="hidden" value="<%=class1%>" name="class1">
				<input type="hidden" value="<%=score%>" name="score">
				<input type="hidden" value="<%=teachername%>" name="teachername">
				<input type="hidden" value="<%=bdate%>" name="bdate">
				<input type="hidden" value="<%=btime%>" name="btime">
				<input type="hidden" value="<%=ptime%>" name="ptime">
				<input type="hidden" value="<%=item%>" name="item">
				<input type="hidden" value="<%=yn%>" name="yn">
				<input type="hidden" value="<%=category%>" name="category">
				<input type="hidden" value="<%=timeflag%>" name="timeflag">
				<input type="hidden" value="<%=oralset%>" name="oralset">
				<input type="hidden" value="<%=topic%>" name="topic">
				<input type="hidden" value="<%=orallevel%>" name="orallevel">
				<input type="hidden" value="<%=briefing%>" name="briefing">
				<input type="hidden" value="<%=sender%>" name="sender">
				<input type="hidden" value="<%=languagecode%>" name="languagecode">
				<input type="hidden" value="<%=tid%>" name="tid">
				<input type="hidden"  name="scid" id="scid" value="<%=scid%>">
				<input type="hidden"  name="delete_item">
				<input type="hidden"  name="delete_item_tid">
				<input type="hidden"  name="yms" id="yms" value="<%=yms%>">

				<TABLE cellSpacing=0 cellPadding=0  width="80%" border=0 >
				<TR><TD colspan=11 class="inputlabel">請輸入學號：</TD></TR>
				<TR><TD colspan=11>
				<input type="hidden" value="<%=tid_d%>" name="tid_d">
				<input type="hidden"  name="slevel_d" id="slevel_d" value="<%=slevel_d%>">
				<input type="hidden"  name="grade_d" id="grade_d" value="<%=grade_d%>">
				<input type="hidden"  name="class1_d" id="class1_d" value="<%=class1_d%>">
				<input type="hidden"  name="department_d" id="department_d" value="<%=department_d%>">
				<input type="hidden"  name="score_d" id="score_d" value="<%=score_d%>">
				<input type="text" value="<%=sid_d%>" maxlength="25" size="20" onblur="ChkStudent_d()" name="sid_d" class="inputtext" >
				<input type="text" value="<%=name_d%>" maxlength="25" size="35" name="name_d" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly >
				<input  type="button" class="inputbutton"  onclick="check_input_d()" <%=btnState%> value="加入一起進行的同學" >
				</TD></TR>
				<TR><TD class="errmsg" colspan=11><%=showmessage2%></TD></TR>
				<TR><TD height="1" bgcolor="#000000" colspan=11></TD></TR>
				<TR class="inputlabel"><TD>同組人員姓名</TD>
				<TD>預約狀態</TD><TD>加入日期</TD><TD>取消日期</TD><TD></TD>
				</TR>
				<TR><TD height="1" bgcolor="#000000" colspan=11></TD></TR>
				<%
				set rs3 = server.CreateObject("adodb.recordset")
				sql = "select * from boo_book_T_M where pid='"&id&"' order by tid "
				rs3.Open sql,msconn,adOpenStatic,adLockReadonly

				icnt=0
				ccnt=0 ' 預約人數
				StrSid_tid = "" '一起參與的學員tid(處方)
				if rs3.EOF then
					response.write "<TR><TD class=""norecord"" colspan=""11"">沒有一起參與的同學</TD></TR>"
				else
					while not rs3.EOF
					icnt=icnt+1
					if rs3("yn")="Y" then
						ccnt = ccnt +1
						if StrSid_tid<>"" then StrSid_tid=StrSid_tid & "','" end if
						StrSid_tid = StrSid_tid & trim(rs3("tid"))
					end if
					if icnt mod 2 = cint(0) then
						vcolor="#E7E7E7"
					else
						vcolor="#FFFFFF"
					end if
				%>
				<TR bgcolor="<%=vcolor%>">
				<TD><%=rs3("sid")%> - <%=rs3("name")%>(<%=rs3("department")%>，<%=rs3("grade")%>)</TD>
				<TD><%=replace(replace(rs3("yn"),"Y","<font color=""blue"">已預約</font>"),"N","<font color=""red"">取消</font>")%></TD><TD><%=rs3("initdate")%></TD><TD><%=rs3("canceldate")%></TD><TD><input type="button" <%=btnState%> onclick="delete_record_item('<%=rs3("id")%>','<%=rs3("tid")%>');" class="inputbutton" value="取消預約" <%if rs3("yn")="N" then response.write "disabled" end if%> ></TD>
				</TR>
				<TR>
					<TD height="1" background="/include/lib/images/untitled.bmp" colspan="11"></TD>
				</TR>
				<%
					rs3.MoveNext
					wend
				end if
			
				set rs3 = nothing
				%>
				<input type="hidden"  name="icnt" id="icnt" value="<%=ccnt%>">
				<input type="hidden"  name="StrSid_tid" id="StrSid_tid" value="<%=StrSid_tid%>">
				
				</TABLE>
				</form>
			</TD></TR>
			</TABLE>	
		</TD>
		</TR>
		</TABLE>
		</TD>
	</TR>
	</TABLE>
<!-- ---------------------------------------------------------------------------------------- -->
	</TD>
</TR>
<TR bgcolor="#333333" height="30">
	<TD class="T1">
	<!-- #include file="lib\bottom.inc" -->
	
	</TD>
</TR>
</TABLE>

</TD>
</TR>
</TABLE>
<iframe style="display:none"  name="iframe_query" id="iframe_query"></iframe>
</BODY>
</HTML>

<!-- #INCLUDE virtual="/include/lib/conn/syconnclose.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->


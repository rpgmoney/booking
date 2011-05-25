<% SESSION.CODEPAGE="950"%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<!-- #INCLUDE FILE="lib/function.asp" -->
<!-- #INCLUDE file="lib/parameter.inc" -->
<%
validate=trim(request("validate"))
nextrec=trim(request("nextrec"))
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
consult=trim(request("consult"))
languagecode=trim(request("languagecode"))
tid=trim(request("tid"))
scid=trim(request("scid"))
deptgroup=trim(request("deptgroup"))

'response.write "deptgroup=" & deptgroup
'response.end
set rs = server.CreateObject("adodb.recordset")
set rs2 = server.CreateObject("adodb.recordset")

btnstatus=""
if timeflag="B" then
	ptime=cdbl(btime+25)
else
	ptime=btime 
end if
'
'response.write "btime=" & btime
if category = "T" then
	sender=ifnull(trim(request("sender")),"bookteacher.asp")
	languagecode="E"
else
	sender=ifnull(trim(request("sender")),"booksteacher.asp")
	'item="口語"
	timeflag="A"
	
	if languagecode="E" then '預約英文小老師才要檢查

		sql = "select * from boo_parameter where priority='1'"
		rs.Open sql,msconn,adOpenStatic,adLockReadonly
		if not rs.EOF then
			'以駐診老師為第一順位
			ww = weekday(bdate)
			sql1 ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate  from boo_schedule a "
			sql1 = sql1 & "left join boo_book_T_M b on a.teacher=b.teachername  and b.btime='"&btime&"' and  b.bdate='"&bdate&"' and b.yn='Y' "
			sql1 = sql1 & " where b.pid is null and a.category='T' and a.btime='"&btime&"' and a.bweek='"&cint(ww)+1&"'  and a.yn='Y' and b.name is null"
			'response.write sql1
			rs2.Open sql1,msconn,adOpenStatic,adLockReadonly
			if not rs2.EOF then
				'駐診老師有空位
				response.write "<script>alert('必須先以駐診老師為優先選擇，口語小老師為第二順位。');</script>"
				showmessage="必須先以駐診老師為優先選擇，口語小老師為第二順位。<a href=bookteacher.asp?BOOK_DATE="&bdate & "><font color='blue'>&nbsp;&nbsp;<img src='images/u24.gif' border=0>&nbsp;&nbsp;跳至『預約駐診老師』</font></a>"
				btnstatus = "disabled"
				'response.redirect "bookteacher.asp?BOOK_DATE="&bdate
			end if
			rs2.close
		end if
		rs.close
	end if
end if


if validate="add" then
	sql = "select * from boo_book_T_M where bdate='"&bdate&"' and sid='"&sid&"' and item='"&item&"' and yn='Y'"
	rs2.Open sql,msconn,adOpenStatic,adLockReadonly
	if not rs2.EOF then
		showmessage = "每人一天相同療程限預約一次，請勿重覆預約。詳細資訊請參考預約規則&nbsp;<a href='#' title='英外語診斷輔導中心　預約規則' onclick=""window.showModalDialog('showrule.asp','','dialogWidth=650px;dialogHeight=650px;status=no');"" ><img src='images/icon_question.gif' border='0'></a>"	
	else
		if timeflag="A" then
			sql = "select * from boo_book_T_M where bdate='"&bdate&"' and btime='"&btime&"' and teachername='"&teachername&"' and category='"&category&"' and yn='Y'"
		else
			sql = "select * from boo_book_T_M where bdate='"&bdate&"' and btime='"&btime&"' and ptime='"&ptime&"' and teachername='"&teachername&"' and category='"&category&"' and yn='Y'"
		end if
		'response.write sql
		rs.Open sql,msconn,adOpenStatic,adLockOptimistic
		if rs.EOF then
			rs.AddNew
			id= getguid()
			if id<>"" then
				rs("id")=id
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
			if timeflag<>"" then
				rs("ptime")=ptime
			end if
			if teachername<>"" then
				rs("teachername")=teachername
			end if
			if scid<>"" then
				rs("scid")=scid
			end if
			
			if item<>"" then
				rs("item")=item
			end if
			if sid<>"" then
				rs("sid")=sid
			end if
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
			if consult<>"" then
				rs("consult")=consult
			end if
			
			if category<>"" then
				rs("category")=category
			end if
			if languagecode<>"" then
				rs("languagecode")=languagecode
			end if
			if tid<>"" then
				rs("tid")=tid
			end if
			rs("yms") = par_yms
			rs("yn") ="Y"
			rs("initdate") = date()
			rs("inituid") = session("sid")


			rs.Update
			if Err.Number=0 then 
				'更新預約次數
				a=UpdateItemTime(item,tid,"1")
				
				response.redirect "bookingtedit.asp?id=" & id & "&category=" & category
			else
				showmessage= Err.Description
			end if
		else
			showmessage="此時段已有人預約。"
		end if
		rs.close
	end if
	rs2.close
end if

if session("classify")="S" or session("classify")="E" then
	sid = session("sid")
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
	
	if (form1.sid.value=="")
        errmsg += "學號不能為空白\n";
	if (form1.name.value=="")
        errmsg += "姓名不能為空白\n";
	if (form1.item0.checked==true)
        errmsg += "請指定預約項目\n";
    if (form1.timeflag.value=="")
        errmsg += "請指定預約時段\n";
	if (form1.bdate.value=="")
        errmsg += "日期不能為空白\n";
    if (form1.btime.value=="")
        errmsg += "時段不能為空白\n";
	if (form1.timeflag0.checked==true)
		errmsg += "請指定節數\n";
	
	if (errmsg=="")
	{
		if (form1.item3.checked==true)
		{
			if (form1.orallevel.value=="")
				errmsg += "請選擇口語級數\n";
			if (form1.oralset.value=="")
				errmsg += "請選擇口語系列\n";
			if (form1.topic.value=="")
				errmsg += "請選擇口語題目\n";
		
		 }
		 else if (form1.item2.checked==true)
		{
			if (form1.consult.value=="")
				errmsg += "諮商主題不能為空白\n";
		
		
		}
		else if (form1.item4.checked==true)
		{
			if (form1.briefing.value=="")
				errmsg += "簡報題目不能為空白\n";
		}
	
	}
	
	
    if (errmsg=="")
        return true;
    else
    {
        alert(errmsg);
        return false;
    }
}

function ChkStudent()
{
	vWinCal2 = window.open("lib/checkstudent.asp?sid="+form1.sid.value+"&languagecode="+ form1.languagecode.value + "&category=" + form1.category.value  ,"iframe_query", "width=200,height=50,status=no,resizable=no,scrollbars=no");
	//vWinCal2.opener = form1;
}
function changesubject()
{
	vWinCal2 = window.open("lib/changesubject.asp?oralset="+form1.oralset.value  ,"iframe_query", "width=200,height=50,status=no,resizable=no,scrollbars=no");
	//vWinCal2.opener = form1;
}
function clickblock(id)
{
	obj=document.getElementById('area_timeflag');
	obj1=document.getElementById('area_consult');
	obj2=document.getElementById('area_oral');
	obj3=document.getElementById('area_briefing');
	

	
	if (id=="0")
	{
		form1.timeflag0.checked=true;
		obj.style.display="none";
		obj1.style.display="none";
		obj2.style.display="none";
		obj3.style.display="none";
	}else if (id=="1")
	{
		form1.timeflag0.checked=true;
		obj.style.display="block";
		obj1.style.display="none";
		obj2.style.display="none";
		obj3.style.display="none";
	}else if (id==2)
	{
		form1.timeflag0.checked=true;
		obj.style.display="block";
		obj1.style.display="block";
		obj2.style.display="none";
		obj3.style.display="none";
	}else if (id=="3")
	{
		form1.timeflag3.checked=true;
		obj.style.display="none";
		obj1.style.display="none";
		obj2.style.display="block";
		obj3.style.display="none";
		
		
	}else if (id=="4")
	{
		form1.timeflag3.checked=true;
		obj.style.display="none";
		obj1.style.display="none";
		obj2.style.display="none";
		obj3.style.display="block";
	}else if (id=="5")
	{
		form1.timeflag3.checked=true;
		obj.style.display="none";
		obj1.style.display="none";
		obj2.style.display="none";
		obj3.style.display="none";
	}else if (id=="6")
	{
		form1.timeflag3.checked=true;
		obj.style.display="none";
		obj1.style.display="none";
		obj2.style.display="none";
		obj3.style.display="none";
	}
	else if (id=="7")
	{
		form1.timeflag3.checked=true;
		obj.style.display="none";
		obj1.style.display="none";
		obj2.style.display="none";
		obj3.style.display="none";
	}
	
	
}

function ChkScore()
{
	vWinCal2 = window.open("lib/checkscore.asp?sid="+form1.sid.value  ,"iframe_query", "width=200,height=50,status=no,resizable=no,scrollbars=no");
	//vWinCal2.opener = form1;
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center"><%if category="T" then response.write "預約教師輔導療程" else response.write "預約小老師輔導療程" end if%></TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="bookingt.asp" onsubmit="return check_input();">
			<input type="hidden" value="add" name="validate">
			<input type="hidden" id="nextrec" name="nextrec">
			<input type="hidden" value="<%=category%>" id="category"  name="category" >
			<input type="hidden" value="<%=languagecode%>" id="languagecode"  name="languagecode" >
			<input type="hidden" value="<%=tid%>"  id="tid"  name="tid" >
			<input type="hidden" value="<%=scid%>"  id="scid"  name="scid" >
			<input type="hidden" value="<%=deptgroup%>"  id="deptgroup"  name="deptgroup" >
			
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>學號：</TD>
						<TD>姓名：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=sid%>" maxlength="25" size="35" onblur="ChkStudent()" name="sid" id="sid" class="inputtext" <%if session("classify")<>"A" then response.write "readonly" end if%>>
						</TD>
						<TD>
						<input type="text" value="<%=name%>" maxlength="25" size="35" name="name" id="name" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly >
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
						<input type="text" value="<%=slevel%>" maxlength="10"   name="slevel" id="slevel" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						<TD>
						<input type="text" value="<%=department%>" maxlength="10"  name="department" id="department" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						<TD>
						<input type="text" value="<%=grade%>" maxlength="10" size="10"  name="grade" id="grade" class="inputtext"  style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						<TD>
						<input type="text" value="<%=class1%>" maxlength="25"  name="class1" id="class1" class="inputtext"  style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						<TD>
						<input type="text" value="<%=score%>" maxlength="25"  name="score" id="score" class="inputtext"  style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD><%if category = "T" then response.write "老師" else response.write "小老師" end if%>：</TD><TD>&nbsp;語言別：&nbsp;</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=teachername%>" maxlength="25" size="35"  name="teachername" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						<TD align="center"><%=languagecode%></TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR id="area_item" ><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>預約項目：<%if category="T"  then%><font color="blue">(預約項目還包括口語、簡報、詩歌、寫作、閱讀，該項目須根據處方籤方可預約)</font><%end if%></TD>
					</TR>
					<TR>
						<TD>
							<TABLE cellSpacing=0 cellPadding=0  border=0 >
							<TR>
							<TD><input type="radio" name="item" id="item0" onclick="clickblock('0');" class="inputtext" value="" <%if item="" or isnull(item) or isempty(item) then response.write "checked" end if%> ></TD><TD >無</TD>
							<TD style="DISPLAY:<%if category="T"  then response.write "block" else response.write "none" end if %>"><input type="radio" name="item" id="item1" onclick="clickblock('1');" class="inputtext" value="診斷" <%if item="診斷" then response.write "checked" end if%> ></TD>
							<TD  style="DISPLAY:<%if category="T"  then response.write "block" else response.write "none" end if %>">診斷</TD>
							<TD style="DISPLAY:<%if category="T"  then response.write "block" else response.write "none" end if %>"><input type="radio" name="item" id="item2" onclick="clickblock('2');" class="inputtext" value="諮商" <%if item="諮商" then response.write "checked" end if%>></TD>
							<TD  style="DISPLAY:<%if category="T"  then response.write "block" else response.write "none" end if %>">諮商</TD>
							<TD  id="item3_option"  style="DISPLAY:none" ><input type="radio" name="item" id="item3" onclick="clickblock('3');" class="inputtext" value="口語" <%if item="口語" then response.write "checked" end if%>></TD>
							<TD  id = "item3_lab"  style="DISPLAY:none" >口語</TD>
							<TD  id="item4_option"  style="DISPLAY:none" ><input type="radio" name="item" id="item4" onclick="clickblock('4');" class="inputtext" value="簡報" <%if item="簡報" then response.write "checked" end if%>></TD>
							<TD  id = "item4_lab"  style="DISPLAY:none" >簡報</TD>
							<TD  id="item5_option"  style="DISPLAY:none" ><input type="radio" name="item" id="item5" onclick="clickblock('5');" class="inputtext" value="詩歌" <%if item="詩歌" then response.write "checked" end if%>></TD>
							<TD  id = "item5_lab"  style="DISPLAY:none" >詩歌</TD>
							<TD  id="item6_option"  style="DISPLAY:none" ><input type="radio" name="item" id="item6" onclick="clickblock('6');" class="inputtext" value="寫作" <%if item="寫作" then response.write "checked" end if%>></TD>
							<TD  id = "item6_lab"  style="DISPLAY:none" >寫作技巧</TD>
							<TD  id="item7_option"  style="DISPLAY:none" ><input type="radio" name="item" id="item7" onclick="clickblock('7');" class="inputtext" value="閱讀" <%if item="閱讀" then response.write "checked" end if%>></TD>
							<TD  id = "item7_lab"  style="DISPLAY:none" >閱讀技巧</TD>
							</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>預約日期：</TD>
						<TD>預約時段：</TD>
						<TD></TD>
					</TR>
					<TR>
						<TD valign="top">
						<input type="text" value="<%=bdate%>" maxlength="25" size="15"  name="bdate" class="inputtext" readonly>
						</TD>
						<TD valign="top">
							<input type="hidden" value="<%=btime%>" maxlength="6"   name="btime" class="inputtext" readonly>
							<select name="btime1" class="inputtext" disabled>
							<option value=""> - 請指定 -</option>
							<optgroup label="上午">
							<option value="0810" <%if btime="0810" then response.write "selected" end if %>>8:10～9:00</option>
							<option value="0910" <%if btime="0910" then response.write "selected" end if %>>9:10～10:00</option>
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
							<option value="1710" <%if btime="1710" then response.write "selected" end if %>>17:10～18:00</option>
							</optgroup>
							</select>
						</TD>
						<TD id="area_timeflag" style="DISPLAY:<%if item="診斷" or item="諮商"  then response.write "block" else response.write "none" end if %>">
							<TABLE cellSpacing=0 cellPadding=0  border=0 >
							<TR>
							<TD><input type="radio" name="timeflag" id="timeflag0" class="inputtext" value="" <%if timeflag="" or isnull(timeflag) or isempty(timeflag) then response.write "checked" end if%> ></TD><TD >未指定</TD>
							<TD><input type="radio" name="timeflag" id="timeflag1" class="inputtext" value="U" <%if timeflag="U" then response.write "checked" end if%> ></TD><TD>上一節（25分）</TD>
							<TD><input type="radio" name="timeflag" id="timeflag2" class="inputtext" value="B" <%if timeflag="B" then response.write "checked" end if%>></TD><TD>下一節（25分）</TD>
							<TD><input type="radio" name="timeflag" id="timeflag3" class="inputtext" value="A" <%if timeflag="A" then response.write "checked" end if%>></TD><TD>上下二節(50分)</TD>
							</TR>
							<TR><TD></TD><TD></TD><TD colspan="11"><font color='blue'>《若邀請其它同學一起參與，請選50分鐘》</font></TD></TR>
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
						<select name="oralset" class="inputtext">
						<option value=""> - 請指定 -</option>
						</select>
						</TD>
						<TD>
						<select name="topic" class="inputtext">
						<option value=""> - 請指定 -</option>
						</select>
						</TD>
						<TD>
						<select name="orallevel" class="inputtext">
						<option value=""> - 請指定口語級數 -</option>
						<option value="Level 1" <%if orallevel="Level 1" then response.write "selected" end if%>>Level 1</option>
						<option value="Level 2" <%if orallevel="Level 2" then response.write "selected" end if%>>Level 2</option>
						<option value="Level 3" <%if orallevel="Level 3" then response.write "selected" end if%>>Level 3</option>
						<option value="Level 4" <%if orallevel="Level 4" then response.write "selected" end if%>>Level 4</option>
						</select>
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
			<TR id="area_consult" style="DISPLAY:<%if item="諮商"  then response.write "block" else response.write "none" end if %>"><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>諮商主題：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=consult%>" maxlength="100" size="55" name="consult" class="inputtext" >
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR>
			<TD>
			<BR>
			<input  type="submit" value="預約" <%=btnstatus%> class="inputbutton" >
			<input  type="button" value="返回" class="inputbutton" onclick="window.location='<%=sender%>'">
			</TD>
			</TR>
			</form>
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

<%
if 1=1 then
%>
<SCRIPT LANGUAGE=javascript>
<!--
	if (form1.sid.value!="")
		ChkStudent();
//-->
</SCRIPT>
<%
end if
%>

<%

showflag = trim(request("showflag"))
if showflag="1" then
	set rsLoad = server.CreateObject("adodb.recordset")
	sql = "select * from boo_parameter where ID = 'A'"
	rsLoad.Open sql,msconn,adOpenStatic,adLockReadOnly
	if not rsLoad.EOF then
		showhint = trim(rsLoad("showhint"))
		
	end if
	rsLoad.close
	if showhint="Y"  then
	%>
	<script language="javascript">
	function window.onload()
	{
		var ls_parm = 'dialogWidth=650px;'
						+ 'dialogHeight=650px;'
						+ 'center=yes;'
						+ 'border=thin;'
						+ 'help=no;'
						+ 'directories=no;'
						+ 'location=no;'
						+ 'status=no';
		window.open('showrule.asp','訊息公告','fullscreen=1,scrollbars=1');
	}
	</script>
	<%end if%>
<%end if%>
<!-- #INCLUDE virtual="/include/lib/conn/syconnclose.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->
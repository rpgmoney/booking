<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
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




'response.write "id=" & id
sender=ifnull(trim(request("sender")),"bookteacher.asp")

set rs = server.CreateObject("adodb.recordset")
set rs2 = server.CreateObject("adodb.recordset")


sql = "select * from boo_book_T_M where id='"&id&"' "
'response.write sql
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
	canceldate=trim(rs("canceldate"))

end if



'response.write "orallevel=" & orallevel

'口語題目
StrSubject=""
if oralset <> "" then
	set rs2 = server.CreateObject("adodb.recordset")
	sql ="select * from boo_orallevel where category='"&oralset&"'"
	rs2.Open sql,msconn,adOpenStatic,adLockReadonly
	
	while not rs2.EOF
		if topic=rs2("topic") then
			StrSubject=StrSubject&"<option value="""&rs2("topic")&""" selected>"&  rs2("topic")&"</option>"
		else
			StrSubject=StrSubject&"<option value="""&rs2("topic")&""" >"&  rs2("topic")&"</option>"
		end if 
		rs2.MoveNext 
	wend
	set rs=nothing
else
	StrSubject="<option value="""" selected>- 無 -</option>"
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

	if (form1.item.value=="3")
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
	vWinCal2 = window.open("lib/checkstudent_d.asp?sid="+AddStudent_Form.sid_d.value  ,"iframe_query", "width=200,height=50,status=no,resizable=no,scrollbars=no");
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

</script>
</HEAD>
<BODY bottomMargin=0 leftMargin=0 topMargin=0 rightMargin=0 marginheight="0" marginwidth="0"  bgcolor="#f4c60d">
<TABLE cellSpacing=1 cellPadding=0 width="760"  height="100%" align="center" >
<TR><TD>
<TABLE cellSpacing=0 cellPadding=0 width="760"  height="100%" align="center" bgColor=#ffffff border=0>
<TR valign="top">
	<TD>
<!-- ---------------------------------------------------------------------------------------- -->
	<P><BR>
	<TABLE cellSpacing=3 cellPadding=3 border=0 width="100%">
	<TR >
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">預約明細(瀏覽)</TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="bookingtedit.asp" onsubmit="return check_input();">
			<input type="hidden" value="add" name="validate">
			<input type="hidden" value="<%=ptime%>" name="ptime">
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>學號：</TD>
						<TD>姓名：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=sid%>" maxlength="25" size="35"  name="sid" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
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
						<TD></TD>
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
						<input type="hidden" value="<%=score%>" maxlength="25"  name="score" class="inputtext"  style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>老師：</TD><TD>預約狀態：</TD><TD>&nbsp;&nbsp;取消日期：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=teachername%>" maxlength="25" size="35"  name="teachername" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						<TD>&nbsp;&nbsp;<%=replace(replace(yn,"Y","<font color=""blue"">已預約</font>"),"N","<font color=""red"">取消</font>")%></TD>
						<TD>&nbsp;&nbsp;<%=canceldate%></TD>
					</TR>
				</TABLE>
			</TD></TR>
			
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>預約項目：&nbsp;&nbsp;</TD>
						<TD>預約日期：</TD>
						<TD>預約時段：</TD>
						<TD></TD>
					</TR>
					<TR>
						<TD><%=item%><input type="hidden" value="<%=item%>" name="item"></TD>
						<TD>
						<input type="text" value="<%=bdate%>" maxlength="25"  name="bdate" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						
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
						<%if cdbl(score) <=90 then %>
						<option value=""> - 請指定 -</option>
						<option value="Conversation Topics" <%if oralset="Conversation Topics" then response.write "selected" end if%>>Conversation Topics</option>
						<%else%>
						<option value=""> - 請指定 -</option>
						<option value="Issues in English I" <%if oralset="Issues in English I" then response.write "selected" end if%>>Issues in English I</option>
						<option value="Issues in English II" <%if oralset="Issues in English II" then response.write "selected" end if%>>Issues in English II</option>
						<%end if%>
						</select>
						</TD>
						<TD>
						<select name="topic" class="inputtext" >
						<option value=""> - 請指定口語題目 -</option>
						<%=StrSubject%>
						</select>
						</TD>
						<TD>
						<%if cdbl(score) >= 90 then %>
						<select name="orallevel" class="inputtext">
						<option value=""> - 請指定口語級數 -</option>
						<option value="Level 1" <%if orallevel="Level 1" then response.write "selected" end if%>>Level 1</option>
						<option value="Level 2" <%if orallevel="Level 2" then response.write "selected" end if%>>Level 2</option>
						<option value="Level 3" <%if orallevel="Level 3" then response.write "selected" end if%>>Level 3</option>
						<option value="Level 4" <%if orallevel="Level 4" then response.write "selected" end if%>>Level 4</option>
						</select>
						
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
			</TD>
			</TR>
			</form>
			</TABLE>
		</TD>
	</TR>
	<tr height="20"><TD></TD><td></td></tr>
	<tr height="7"><TD></TD><td background="images/lin04.gif"></td></tr>
	<TR>
		<TD></TD><TD>
		<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0 >
		<TR><TD><font color="blue">※此區塊為一起參與的同學</font><BR></TD></TR>

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
				
				<input type="hidden" value="<%=timeflag%>" name="timeflag">
				<input type="hidden" value="<%=oralset%>" name="oralset">
				<input type="hidden" value="<%=topic%>" name="topic">
				<input type="hidden" value="<%=orallevel%>" name="orallevel">
				<input type="hidden" value="<%=briefing%>" name="briefing">
				<input type="hidden" value="<%=sender%>" name="sender">

				<TABLE cellSpacing=0 cellPadding=0  width="80%" border=0 >
				
				<TR><TD class="errmsg" colspan=11><%=showmessage2%></TD></TR>
				<TR><TD height="1" bgcolor="#000000" colspan=11></TD></TR>
				<TR class="inputlabel"><TD>同組人員姓名</TD>
				<TD>預約狀態</TD><TD>加入日期</TD><TD>取消日期</TD><TD></TD>
				</TR>
				<TR><TD height="1" bgcolor="#000000" colspan=11></TD></TR>
				<%
				set rs3 = server.CreateObject("adodb.recordset")
				sql = "select * from boo_book_T_M where pid='"&id&"'"
				rs3.Open sql,msconn,adOpenStatic,adLockReadonly

				icnt=0
				if rs3.EOF then
					response.write "<TR><TD class=""norecord"" colspan=""11"">沒有一起參與的同學</TD></TR>"
				else
					while not rs3.EOF
					icnt=icnt+1
					if icnt mod 2 = cint(0) then
						vcolor="#E7E7E7"
					else
						vcolor="#FFFFFF"
					end if
				%>
				<TR bgcolor="<%=vcolor%>">
				<TD><%=rs3("sid")%> - <%=rs3("name")%>(<%=rs3("department")%>，<%=rs3("grade")%>)</TD>
				<TD><%=replace(replace(rs3("yn"),"Y","<font color=""blue"">正常</font>"),"N","<font color=""red"">取消</font>")%></TD><TD><%=rs3("initdate")%></TD><TD><%=rs3("canceldate")%></TD><TD></TD>
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
				<input type="hidden"  name="icnt" id="icnt" value="<%=icnt%>">
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


<%@Language=VBScript LCID=1033%>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<!-- #INCLUDE virtual="/include/lib/asp/jsCalendar.asp" -->

<%

sid=trim(request("sid"))
enable=trim(request("enable"))
page=trim(request("page"))
initdateS=trim(request("initdateS"))
initdateE=trim(request("initdateE"))

category=trim(request("category"))
sender=ifnull(trim(request("sender")),"studentlist.asp")

slevel=trim(request("slevel"))
department=trim(request("department"))
if session("hos_code")<>"930105" then
	HOS_CODE = session("hos_code")
end if

sender=server.urlencode(replace(request.servervariables("PATH_INFO")&"?page="&page& "&sid=" & sid& "&enable=" & enable& "&slevel=" & slevel& "&department=" & department & "&category="&category ,"%","*"))

'系所
StrDepartment="<option value=''> - 全部 - </option>"
set rsLoad = server.CreateObject("adodb.recordset")
sql ="select * from s90_unit where unt_std='Y' order by unt_sort_seq  "
rsLoad.Open sql,syconn,adOpenStatic,adLockReadonly

while not rsLoad.EOF
	if department=rsLoad("unt_name_abr") then
		StrDepartment=StrDepartment&"<option value="""&rsLoad("unt_name_abr")&""" selected>"&  rsLoad("unt_name_abr")&"</option>"
	else
		StrDepartment=StrDepartment&"<option value="""&rsLoad("unt_name_abr")&""" >"&  rsLoad("unt_name_abr")&"</option>"
	end if 
	rsLoad.MoveNext 
wend
rsLoad.close

'學制
Strslevel="<option value=''> - 全部 - </option>"
sql ="select * from s90_degree "
rsLoad.Open sql,syconn,adOpenStatic,adLockReadonly

while not rsLoad.EOF
	if slevel=rsLoad("dgr_name") then
		Strslevel=Strslevel&"<option value="""&rsLoad("dgr_name")&""" selected>"&  rsLoad("dgr_name")&"</option>"
	else
		Strslevel=Strslevel&"<option value="""&rsLoad("dgr_name")&""" >"&  rsLoad("dgr_name")&"</option>"
	end if 
	rsLoad.MoveNext 
wend
rsLoad.close

'
set rsLoad=nothing

if enable="" or isnull(enable) or isempty(enable) then
	enable="Y"
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
function JumpPage1()
{
	var obj;
	obj= document.getElementById("selectPage");
	var index=obj.value;
	var  frmlistform = document.getElementById("news_form");
	frmlistform.page.value=index;
	frmlistform.submit();
}
function changepage(v)
{
	
	var  frmlistform = document.getElementById("news_form");
	frmlistform.page.value=v;
	frmlistform.submit();
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
<TR>
	<TD align="center"><font color="red"><%=showmessage%></font></TD>
</TR>

<TR valign="top">
	<TD>
<!-- ---------------------------------------------------------------------------------------- -->
	<P><BR>
	<TABLE cellSpacing=3 cellPadding=3 border=0 width="100%">
	<TR >
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%if category="S" then response.write "中心學員基本資料維護" else response.write "學員學習紀錄分析統計" end if%> </TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>
	
	<TR>
		<TD></TD><TD valign="top">
			<form id="news_form" name="news_form" method="post" action="studentlist.asp" >
			<input type="hidden" name="page" value="">
			<input type="hidden" name="category" value="<%=category%>">
			
			<TABLE cellSpacing=0 cellPadding=0 border=0>
			<TR class="inputlabel"><TD width="20"></TD>
				<TD>學號</TD>
				<TD>學制</TD>
				<TD>系所</TD>
				<TD>開啟</TD>
				<TD></TD>
			</TR>
			<TR><TD></TD>
				<td>
					<input type="text" name="sid" value="<%=sid%>"  maxlength="10" size="15" class="inputtext" >
				</td>
				<td>
					<select name="slevel" class="inputtext">
					<%=Strslevel%>
					</select>
				</td>
				<td>
					<select name="department" class="inputtext">
					<%=StrDepartment%>
					</select>
				</td>
				
				<td>
					<select name="enable" class="inputtext">
					<option value="all" <%if enable="all" then response.write "selected" end if%>> - 全部 - </option>
					<option value="Y" <%if enable="Y" then response.write "selected" end if%>> - 是 - </option>
					<option value="N" <%if enable="N" then response.write "selected" end if%>> - 否 - </option>
					</select>
				</td>
				<TD>
				<input  type="submit" value="查詢" class="inputbutton">
				<%if category="S" then%>
				<input  type="button"  onclick="window.location='register.asp?flag=1&sid=<%=sid%>&sender=<%=sender%>&btncontrol=Y'" value="新增" class="inputbutton" <%if sid="" or isempty(sid) or isnull(sid) then response.write "disabled" end if%>>
				<%end if%>
				</TD>
			</TR>
			</TABLE>
			<TABLE cellSpacing=0 cellPadding=0 border=0>
			<TR class="inputlabel"><TD width="20"></TD>
				<TD>註冊日期起迄：</TD>
				<TD></TD>
			</TR>
			<TR class="inputlabel"><TD width="20"></TD>
				<TD>
					<TABLE cellSpacing=0 cellPadding=0 border=0 >
					<TD><input type="text" id="initdateS" name="initdateS" value="<%=initdateS%>"  maxlength="6" class="inputtext"><img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('initdateS')" class="showhand"></TD>
					<TD class="inputlabel">&nbsp;~&nbsp;</TD>
					<TD><input type="text" id="initdateE" name="initdateE" value="<%=initdateE%>"  maxlength="6" class="inputtext"><img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('initdateE')" class="showhand"></TD>
					</table>
				</TD>
				<TD></TD>
			</TR>
			</form>
			</TABLE>
			
		</TD>
	</TR>
	<%
		set rs = server.CreateObject("adodb.recordset")
		sql = "select * from boo_profile a where classify in ('S','E') "
		if sid<>"" then
			sql = sql & " and ( a.sid='"&sid&"' or a.name like '%"&sid&"%' ) "
		end if
		if initdateS<>"" and initdateE="" then
			sql = sql & " and a.initdate >= " & NumberToDateFormat(initdateS)
		end if
		if initdateS="" and initdateE<>"" then
			sql = sql & " and a.initdate<=" & NumberToDateFormat(initdateE)
		end if
		if initdateS<>"" and initdateE<>"" then
			sql = sql & " and a.initdate>='" & NumberToDateFormat(initdateS) & "' and a.initdate<='" & NumberToDateFormat(initdateE) & "' "
		end if
		if slevel<>"" then
			sql = sql & " and  a.slevel='"&slevel&"'  "
		end if
		if department<>"" then
			sql = sql & " and  a.department='"&department&"'  "
		end if
		if enable<>"all" then
			sql = sql & " and  a.enable='"&enable&"'  "
		end if
		
		sql = sql & " order by a.slevel,a.department,a.grade,a.class1,a.sid"
		
			
		'response.write sql
		rs.Open sql,msconn,adOpenStatic,adLockReadonly
		if not rs.EOF then
			rscount=rs.RecordCount
			lcount=30   '設定每頁顯示的筆數
			m_page=request("page")
			if m_page="" then
				m_page=1
			else
				m_page=cint(m_page)   
			end if
			point=(m_page-1)*lcount+1   'Record Point
			if m_page>1 then
			  rs.move point-1
			end if

			'計算共幾頁
			pagecount=int(rscount/lcount)
			if rscount mod lcount >0 then
			  pagecount=pagecount+1
			end if   
			ln=point
		end if
	%>
	
	<TR>
		<TD></TD><TD valign="top">
		<!--上一頁 , 下一頁  -->
			<TABLE cellSpacing=0 cellPadding=0 align="center" border=0 width="95%">
			<TR>
			<%if category="S" then%>
			<TD align="left">※先以學號查詢後再按新增</TD>
			<%end if%>
			<TD>
				<TABLE cellSpacing=0 cellPadding=0 border=0 align="right">
				<TR>
				<%if m_page<=1 then  %>
				<TD><img src="/include/lib/images/arrow_left_no1.gif"></TD>
				<TD>&nbsp;<font color="#CCCCCC">上一頁</font>&nbsp;</TD>
				<%else%>
				<TD><img src="/include/lib/images/arrow_left1.gif"></TD>
				<TD class="showhand" onclick="changepage(<%=m_page-1%>)">&nbsp;上一頁&nbsp;</TD>
				<%end if%>
				<TD>｜</TD>
				<%if m_page>=pagecount then %>
				<TD>&nbsp;<font color="#CCCCCC">下一頁&nbsp;</font></TD>
				<TD><img src="/include/lib/images/arrow_right_no1.gif"></TD>
				<%else%>
				<TD class="showhand" onclick="changepage(<%=m_page+1%>)">&nbsp;下一頁&nbsp;</TD>
				<TD><img src="/include/lib/images/arrow_right1.gif"></TD>
				<%end if%>
				</TR>
				</TABLE>
			</TD></TR>
			</TABLE>
		<!--  -->
		</TD>
	</TR>
	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=0 cellPadding=0 align="center" border=0 width="95%">
			<TR>
				<TD height="1" bgcolor="#000000" colspan="12"></TD>
			</TR>
			<TR class="inputlabel">
				<TD></TD><TD>學號</TD><TD>姓名</TD><TD>性別</TD><TD>學制</TD><TD>系所</TD><TD>年級</TD><TD>班別</TD><TD>註冊時間</TD><TD>策略</TD><TD>風格</TD><TD></TD>
			</TR>
			<TR>
				<TD height="1" bgcolor="#000000" colspan="12"></TD>
			</TR>
			<% 
			icnt=0
			if rs.EOF then
				response.write "<TR><TD class=""norecord"" colspan=""12"">沒有符合條件的資料顯示</TD></TR>"
			else
				do while not rs.eof and ln<=(point+lcount)-1 
				icnt=icnt+1
				if icnt mod 2 = cint(0) then
					if category = "S" then
						vcolor="#F8D6D1"
					else
						vcolor="#E7E7E7"
					end if
				else
					vcolor="#FFFFFF"
				end if
			%>
			<TR bgcolor="<%=vcolor%>">
				<TD>
				<%if category="S" then%>
				<a href="studentedit.asp?sid=<%=rs("sid")%>&btncontrol=Y&sender=<%=sender%>"><img border="0"  src="/include/lib/images/wri.gif"></a>
				<%end if%>
				</TD>
				<TD><%=rs("sid")%></TD><TD><%=rs("name")%></TD><TD><%=rs("sex")%></TD><TD><%=rs("slevel")%></TD><TD><%=rs("department")%></TD>
				<TD><%=rs("grade")%></TD><TD><%=rs("class1")%></TD><TD><%=datetoNumformat(rs("initdate"))%></TD>
				<TD>
				<input type="button" value="問卷分析" <%if rs("strategy_yn")<>"Y" then response.write "disabled" end if%> onclick="window.open('qstrategyreport.asp?sid=<%=rs("sid")%>','_blank','height=600, resizable=0, scrollbars=1, menubar=1, toolbar=1, top=10')" class="inputbutton"></TD>
				<TD>
				<input type="button" value="問卷分析" <%if rs("sytle_yn")<>"Y" then response.write "disabled" end if%> onclick="window.open('qstylereport.asp?sid=<%=rs("sid")%>','_blank','height=600, resizable=0, scrollbars=1, menubar=1, toolbar=1, top=10')" class="inputbutton">
				</TD>
				<TD>
				<%if category="R" then%>
				<a href="recordprofile.asp?sid=<%=rs("sid")%>&sender=<%=sender%>"><img border="0"  src="images/record.gif"></a>&nbsp;&nbsp;
				<a href="recordreport.asp?sid=<%=rs("sid")%>&sender=<%=sender%>"><img border="0"  src="images/icon_01.gif"></a>&nbsp;&nbsp;
				<a href="languagescore.asp?sid=<%=rs("sid")%>&sender=<%=sender%>"><img border="0"  src="images/75.gif"></a>
				<%end if%>
				</TD>
			</TR>
			<TR>
				<TD height="1" background="/include/lib/images/untitled.bmp" colspan="12"></TD>
			</TR>
			<%
				rs.MoveNext
				ln=ln+1
				Loop
			end if
			

			%>
			</TABLE>
		</TD>
	</TR>
	<%if rscount>0 then %>
	<TR valign="top"><TD></TD>
	<TD >
			<table cellSpacing=1 cellPadding=2 border=0 align="right">
			<tr><td>
			<%
				response.write "第" & m_page & "頁/共" &pagecount &"頁</td>"
				Response.Write "<td>&nbsp;第&nbsp;</td><td><select name=selectPage id=selectPage onchange=JumpPage1() class=inputtext style=width:50>"
				for i=1 to pagecount
					if (i<>m_page)  then
						Response.Write "<option value="&i&">"&i&"</option>"
					else
						Response.Write "<option value="&i&" selected>"&i&"</option>"
					end if
				Next
				Response.Write "</select><td>&nbsp;頁</td></td>"
			%>
			<td width="20">&nbsp;</td></tr>
			</table>
	</TD></TR>
	<%end if%>
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

</BODY>
</HTML>
<!-- #INCLUDE virtual="/include/lib/conn/syconnclose.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->
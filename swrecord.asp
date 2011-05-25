<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<!-- #INCLUDE virtual="/include/lib/asp/jsCalendar.asp" -->

<%
sid=trim(request("sid"))
category=trim(request("category"))
if sid="" or isnull(sid) or isempty(sid) then
	sid=session("sid")
end if
if sid="S224955279" then sid="1095101007" end if

sender=trim(request("sender"))


set dic = server.CreateObject("scripting.dictionary")
dic.Add "1","日"
dic.Add "2","一"
dic.Add "3","二"
dic.Add "4","三"
dic.Add "5","四"
dic.Add "6","五"
dic.Add "7","六"


set rs = server.CreateObject("adodb.recordset")
set rs2 = server.CreateObject("adodb.recordset")
sql = "select * from boo_profile   where sid='"& sid &"' and  classify in ('S','E') "
rs.Open sql,msconn,adOpenStatic,adLockReadonly
if  rs.EOF then
	showmessage ="你不是學員喔，此介面只提供學員瀏覽。"
else
	sid=trim(rs("sid"))
	name=trim(rs("name"))
	slevel=trim(rs("slevel"))
	grade=trim(rs("grade"))
	class1=trim(rs("class1"))
	department=trim(rs("department"))


end if

rs.close

%>
<HTML>
<HEAD>
<TITLE> 【LDCC英外語能力診斷輔導中心】 </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/stm31.js" type="text/javascript"></SCRIPT>
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>

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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center"><%if category="T" then response.write "個人補充教材上課紀錄" else if category="S" then response.write "個人自學軟體療程上課紀錄" else if category="F" then response.write "個人自學療程上課紀錄"  end if%></TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>
	<TR >
		<TD></TD>
		<TD>
			<TABLE cellSpacing=1 cellPadding=2  border=0  width="95%"  >
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR >
						<TD class="inputlabel">學號：</TD><TD><%=sid%>&nbsp;</TD>
						<TD class="inputlabel">姓名：</TD><TD><%=name%>&nbsp;</TD>
						<TD class="inputlabel">學制：</TD><TD><%=slevel%>&nbsp;</TD>
						<TD class="inputlabel">系所：</TD><TD><%=department%>&nbsp;</TD>
						<TD class="inputlabel">年級：</TD><TD><%=grade%>&nbsp;</TD>
						<TD class="inputlabel">班級：</TD><TD><%=class1%>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD><TD align="right"><%if sender<>"" then%><input  type="button" value="&nbsp;&nbsp;返回&nbsp;&nbsp;" class="inputbutton" onclick="window.location='<%=sender%>'"><%end if%></TD></TR>
			</TABLE>
		</TD>
	</TR>
	
	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=0 cellPadding=0  border=0 width="95%" >
			<TR>
				<TD height="1" bgcolor="#000000" colspan="11"></TD>
			</TR>
			<TR class="inputlabel">
				<TD></TD><TD>日期</TD><TD>開始時間</TD><TD>結束時間</TD><TD>分鐘</TD><TD>項目</TD><TD></TD><TD></TD>
			</TR>
			<TR class="inputlabel">
				<TD></TD><TD colspan="6"><%if category="S" then response.write "層級/主題" else response.write "使用項目" end if%> </TD><TD></TD><TD></TD>
			</TR>
			<TR>
				<TD height="1" bgcolor="#000000" colspan="11"></TD>
			</TR>
			<%
				sql = "select a.*,b.itemname,c.level as softwarelevel,c.topic,c.usageitem  from boo_book_software a  "
				sql = sql  & "left join ( "
				sql = sql &						"select floor+ '( ' +software  + ' )' as itemname ,id from boo_software where category='S' "
				sql =sql &						" union "
				sql = sql &						" select  software  as itemname ,id  from boo_software where category='T'  "
				sql =sql &						" union "
				sql = sql &						" select  item  as itemname ,id  from boo_self_item where yn='Y'  "
				sql = sql &				  " ) b  on   a.item=b.id   "
				sql = sql & " left join  boo_software_record c on a.id=c.tid "
				sql = sql & "  where   a.signin is not null and a.sid='"&sid&"' and a.category='"&category&"' "
				sql = sql & " order by a.bdate desc"

				'response.write sql
				rs.Open sql,msconn,adOpenStatic,adLockReadonly
				if rs.EOF then
					response.write "<TR><TD  colspan=11><font   class=""norecord"" >沒有上課紀錄顯示</font></TD></TR>"
				else
					icnt=0
					while not rs.EOF
					icnt=icnt+1
					if icnt mod 2 = cint(0) then
						vcolor="#E7E7E7"
					else
						vcolor="#FFFFFF"
					end if
				%>
				<TR  bgcolor="<%=vcolor%>">
				<TD></TD>
				<TD><%=rs("bdate")& "(&nbsp;"&dic.Item(cstr(cint(weekday(NumberToDateFormat(rs("bdate"))))))&"&nbsp;)"%></TD>
				<TD><%=rs("stime")%></TD><TD><%=rs("etime")%></TD>
				<TD><%=rs("summin")%></TD>
				<TD><%=rs("itemname")%></TD>
				<TD></TD><TD></TD>
				</TR>
				<TR bgcolor="<%=vcolor%>">
					<TD></TD><TD colspan="5"><%if category="S" then response.write rs("softwarelevel")  & "&nbsp;/&nbsp;"  & rs("topic") else response.write rs("usageitem") end if%> &nbsp;</TD><TD></TD><TD></TD>
				</TR>
				<TR>
					<TD height="1" background="/include/lib/images/untitled.bmp" colspan="11"></TD>
				</TR>
				<%
						rs.MoveNext
					wend 
				end if

				rs.Close



			%>
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


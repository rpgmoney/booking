<%@Language=VBScript LCID=1033%>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->

<%

sdate=trim(request("sdate"))
edate=trim(request("edate"))

slevel=trim(request("slevel"))
department=trim(request("department"))

set rs=server.CreateObject("adodb.recordset")
set rs2=server.CreateObject("adodb.recordset")
%>
<HTML>
<HEAD>
<TITLE> 【LDCC英外語能力診斷輔導中心】  </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>
</HEAD>
<BODY bottomMargin=0 leftMargin=2 topMargin=0 rightMargin=2 marginheight="0" marginwidth="0"  >


<P align="center" class="inputlabel"><font size="4">全校學生每天使用人次人數分析表</font></P>

<TABLE cellSpacing=1 cellPadding=3 align="center" width="700" border=1 >
<TR class="inputlabel" bgcolor="#E7E7E7">
	<td nowrap width="50">&nbsp;</td>
	<td nowrap width="100">日期</td>
	<td nowrap width="100">人次</td>
	<td nowrap width="100">人數</td>
	
</TR>
<%

sql = "select bdate,count(*) as cnt  from boo_book_T_M where yn='Y' and signin  is not null  "
if sdate<>"" and edate="" then
	sql = sql & " and  Cast(bdate as int) >= '" & sdate& "'  "
end if
if sdate="" and edate<>"" then
	sql = sql & " and  Cast(bdate as int)<= '" & edate& "'  "
end if
if sdate<>"" and edate<>"" then
	sql = sql & " and ( Cast(bdate as int) >= '" & sdate & "'   and Cast(bdate as int) <= '" & edate& "' )"
end if
if slevel<>"" then
	sql = sql & " and slevel='"&slevel&"'"
end if
if department<>"" then
	sql = sql & " and department='"&department&"'"
end if

sql = sql & " group by bdate  order by bdate "
'response.write sql
'response.end
rs.Open sql,msconn,adOpenStatic,adLockReadonly
	if not rs.eof then 
		totalcnt = 0
		while not rs.EOF 
			rc=rc +1
			totalcnt = cdbl(totalcnt) + cdbl(ifnull(rs("cnt"),0))
		%>
		<tr >
			<td nowrap><%=rc%></td>
			<td nowrap><%=rs("bdate")%></td>
			<td nowrap ><%=rs("cnt")%></td>
			<td nowrap ><%=rs("cnt")%></td>
		</tr>
	<%	
			rs.movenext
		wend
	%>
		<tr >
			<td nowrap>&nbsp;</td>
			<td nowrap>合計</td>
			<td nowrap ><%=totalcnt%></td>
			<td nowrap ><%=totalcnt%></td>
		</tr>
<%
	else
%>
		<TR ><TD colspan="6" align="center"><FONT color=gray>沒有符合條件的資料顯示</FONT></TD></TR>
<%
	end if
	rs.close

%>

</table>
<P><BR>
</BODY>
</HTML>
<!-- #INCLUDE virtual="/include/lib/conn/syconnclose.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->

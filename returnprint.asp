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

function getreturndate(vdate,vsid)
	dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	sql = "select bdate from boo_book_T_M where sid='"&vsid&"' and Cast(bdate as int) >='"&datetoNumformat(dateadd("d",-12,NumberToDateFormat(vdate)))&"' and Cast(bdate as int) <='"&datetoNumformat(dateadd("d",12,NumberToDateFormat(vdate)))&"' and item='�E�_' "
	'getreturndate = sql
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
	if not rs.Eof then
		getreturndate=rs("bdate")
	else
		getreturndate="&nbsp;"
	end if

	rs.close
	set rs = nothing
end function

%>
<HTML>
<HEAD>
<TITLE> �iLDCC�^�~�y��O�E�_���ɤ��ߡj  </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>

</HEAD>
<BODY bottomMargin=0 leftMargin=2 topMargin=0 rightMargin=2 marginheight="0" marginwidth="0"  >


<P align="center" class="inputlabel"><font size="4">�C�����^�E�W��</font></P>

<TABLE cellSpacing=1 cellPadding=3 align="center" width="700" border=1 >
<TR class="inputlabel" bgcolor="#E7E7E7">
	<td nowrap width="30">&nbsp;</td>
	<td nowrap width="80">�E�_���</td>
	<td nowrap width="100">��t</td>
	<td nowrap width="50">�~��</td>
	<td nowrap width="100">�m�W</td>
	<td nowrap width="100">�Ѯv</td>
	<td nowrap width="100">�E�_���e</td>
	<td nowrap width="80">���^�E���</td>
	<td nowrap width="80">�^�E���</td>
</TR>
<%

sql = "select a.sid,a.name,a.slevel,a.grade,a.class1,a.department,a.teachername,b.backdate,a.bdate,b.notice,b.tid,b.content  "
sql = sql & "  from boo_book_T_M a inner join boo_diagnosis b on a.id=b.tid where 1=1  "
if sdate<>"" and edate="" then
	sql = sql & " and  Cast(backdate as int) >= '" & sdate& "'  "
end if
if sdate="" and edate<>"" then
	sql = sql & " and  Cast(backdate as int) <= '" & edate& "'  "
end if
if sdate<>"" and edate<>"" then
	sql = sql & " and ( Cast(backdate as int) >= '" & sdate & "'   and Cast(backdate as int) <= '" & edate& "' )"
end if
if slevel<>"" then
	sql = sql & " and slevel='"&slevel&"'"
end if
if department<>"" then
	sql = sql & " and department='"&department&"'"
end if

sql = sql & " order by b.backdate"
'response.write sql
'response.end
rs.Open sql,msconn,adOpenStatic,adLockReadonly
	if not rs.eof then 
		
		while not rs.EOF 
			rc=rc +1
		%>
		<tr >
			<td nowrap><%=rc%></td>
			<td nowrap><%=rs("bdate")%></td>
			<td nowrap ><%=rs("department")%></td>
			<td nowrap ><%=rs("grade")%></td>
			<td nowrap ><%=rs("sid")%> - <%=rs("name")%></td>
			<td nowrap ><%=rs("teachername")%></td>
			<td nowrap ><%=rs("content")%></td>
			<td nowrap ><%=rs("backdate")%></td>
			<td nowrap ><%=getreturndate(rs("backdate"),rs("sid"))%></td>
			
		</tr>
	<%	
			rs.movenext
		wend
	%>

<%
	else
%>
		<TR ><TD colspan="6" align="center"><FONT color=gray>�S���ŦX���󪺸�����</FONT></TD></TR>
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

<%@Language=VBScript LCID=1033%>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->

<%

sdate=trim(request("sdate"))
edate=trim(request("edate"))

scope=replace(trim(request("scope")),", ","','")



set rs=server.CreateObject("adodb.recordset")
set rs2=server.CreateObject("adodb.recordset")

sql = "select a.*,b.subject,b.date1,stime,etime  from boo_book_lecture a  "
sql = sql  & "left join ( "
sql = sql &						"select  id,subject,date1,stime,etime  from boo_lecture  "
sql = sql &				  " ) b  on  a.lid=b.id  "
sql = sql & "  where 1=1  "
if sdate<>"" and edate="" then
	sql = sql & " and  Cast(date1  as int) >= '" & sdate& "'  "
end if
if sdate="" and edate<>"" then
	sql = sql & " and  Cast(date1 as int) <= '" & edate& "'  "
end if
if sdate<>"" and edate<>"" then
	sql = sql & " and ( Cast(date1 as int) >= '" & sdate & "'   and Cast(date1 as int) <= '" & edate& "' )"
end if

if scope<>"" then
	sql = sql & " and category in ( '" & scope &"')  "
end if
sql = sql & " order by date1"
'response.write sql
'response.end
rs.Open sql,msconn,adOpenStatic,adLockReadonly


'rs.close
'set rs=nothing
%>
<HTML>
<HEAD>
<TITLE> �iLDCC�^�~�y��O�E�_���ɤ��ߡj  </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>

</HEAD>
<BODY bottomMargin=0 leftMargin=2 topMargin=0 rightMargin=2 marginheight="0" marginwidth="0"  >

<%
	if not rs.eof then 
%>
<P align="center" class="inputlabel"><font size="4">�B��ҵ{/�~�y�ǲ����y����</font></P>

<TABLE cellSpacing=1 cellPadding=2 align="center" border=1 >
<TR class="inputlabel" bgcolor="#E7E7E7">
	<td nowrap width="30">&nbsp;</td>
	<td nowrap width="80">���</td>
	<td nowrap width="80">�W��</td>
	<td nowrap width="80">�Ǩ�</td>
	<td nowrap width="80">�t�ҦW��</td>
	<td nowrap width="50">�~��</td>
	<td nowrap width="80">�m�W</td>
	<td nowrap width="100">���O</td>
</TR>
	<%
		while not rs.EOF 
			rc=rc +1
			if rc mod 2 = cint(0) then
				vcolor="#E0F7DD"
			else
				vcolor="#FFFFFF"
			end if
	%>
		<tr bgcolor="<%=vcolor%>">
			<td nowrap><%=rc%></td>
			<td nowrap><%=rs("date1")%></td>
			<td nowrap><%=rs("subject")%></td>
			<td nowrap><%=rs("slevel")%></td>
			<td nowrap><%=rs("department")%></td>
			<td nowrap><%=rs("grade")%></td>
			<td nowrap><%=rs("sid")%> - <%=rs("name")%></td>
			<td nowrap><%=replace(replace(rs("category"),"C","�B��ҵ{"),"L","�~�y�ǲ����y")%></td>
		</tr>
	<%	
			rs.movenext
		wend
	%>
</table>
<%
	else
		Response.Write "<FONT class=normal><FONT color=gray>- �S���ŦX���󪺸����� -</FONT></FONT>"
	end if
%>



<P><BR>
</BODY>
</HTML>
<!-- #INCLUDE virtual="/include/lib/conn/syconnclose.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->
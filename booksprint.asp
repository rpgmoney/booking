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

sql = "select a.*,b.itemname from boo_book_software a  "
sql = sql  & "left join ( "
sql = sql &						"select floor+ '( ' +software  + ' )' as itemname ,id from boo_software where category='S' "
sql =sql &						" union "
sql = sql &						" select  software  as itemname ,id  from boo_software where category='T'  "
sql = sql &				  " ) b  on  a.item=b.id  "
sql = sql & "  where a.yn='Y'  and a.signin  is not null "
if sdate<>"" and edate="" then
	sql = sql & " and  Cast(bdate as int) >= '" & sdate& "'  "
end if
if sdate="" and edate<>"" then
	sql = sql & " and  Cast(bdate as int)<= '" & edate& "'  "
end if
if sdate<>"" and edate<>"" then
	sql = sql & " and ( Cast(bdate as int) >= '" & sdate & "'   and Cast(bdate as int) <= '" & edate& "' )"
end if

if scope<>"" then
	sql = sql & " and category in  ('" & scope &"')  "
end if
sql = sql & " order by a.category,bdate"
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
<P align="center" class="inputlabel"><font size="4">�۾ǳn��/�ɥR�Ч�����</font></P>

<TABLE cellSpacing=1 cellPadding=2 align="center" border=1 >
<TR class="inputlabel" bgcolor="#E7E7E7">
	<td nowrap width="30">&nbsp;</td>
	<td nowrap width="80">���ɤ��</td>
	<td nowrap width="80">�Ǩ�</td>
	<td nowrap width="80">�t�ҦW��</td>
	<td nowrap width="50">�~��</td>
	<td nowrap width="80">�m�W</td>
	<td nowrap width="80">�}�l�ɶ�</td>
	<td nowrap width="80">�����ɶ�</td>
	<td nowrap width="80">�X�p����</td>
	<td nowrap width="100">���O</td>
	<td nowrap width="100">�w������</td>
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
			<td nowrap><%=rs("bdate")%></td>
			<td nowrap><%=rs("slevel")%></td>
			<td nowrap><%=rs("department")%></td>
			<td nowrap><%=rs("grade")%></td>
			<td nowrap><%=rs("sid")%> - <%=rs("name")%></td>
			<td nowrap><%=rs("stime")%></td>
			<td nowrap><%=rs("etime")%></td>
			<td nowrap><%=rs("summin")%></td>
			<td nowrap><%=replace(replace(rs("category"),"S","�۾ǳn��"),"T","�ɥR�Ч�")%></td>
			<td nowrap><%=rs("itemname")%></td>
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
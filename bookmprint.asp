<%@Language=VBScript LCID=1033%>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<%

sdate=trim(request("sdate"))
edate=trim(request("edate"))

scope=replace(trim(request("scope")),", ","','")
item=replace(trim(request("item")),", ","','")


set rs=server.CreateObject("adodb.recordset")
set rs2=server.CreateObject("adodb.recordset")

sql =   " select  a.*,b.deptgroup   from boo_book_T_M a   left join   boo_schedule b on a.scid=b.scid  where a.yn='Y' and signin is not null "
if sdate<>"" and edate="" then
	sql = sql & " and  Cast(bdate as int) >= '" & sdate& "'  "
end if
if sdate="" and edate<>"" then
	sql = sql & " and  Cast(bdate as int)<= '" & edate& "'  "
end if
if sdate<>"" and edate<>"" then
	sql = sql & " and ( Cast(bdate as int) >= '" & sdate & "'   and  Cast(bdate as int)<= '" & edate& "' )"
end if
if item<>"" then
	sql = sql & " and item in ( '" & item &"' ) "
end if
if scope<>"" then
	sql = sql & " and a.category in  ( '" & scope &"' ) "
end if
sql = sql & " order by bdate"
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
<P align="center" class="inputlabel"><font size="4">�Юv/�p�Ѯv�������{����</font></P>

<TABLE cellSpacing=1 cellPadding=2 align="center" border=1 >
<TR class="inputlabel" bgcolor="#E7E7E7">
	<td nowrap width="30">&nbsp;</td>
	<td nowrap width="80">���ɤ��</td>
	<td nowrap width="50">�`��</td>
	<td nowrap width="100">�Ǩ�</td>
	<td nowrap width="100">��t</td>
	<td nowrap width="30">�~��</td>
	<td nowrap width="100">�m�W</td>
	<td nowrap width="80">�^�˦��Z</td>
	<td nowrap width="80">���O</td>
	<td nowrap width="100">�Ѯv</td>
	<td nowrap width="150">�E�_�԰ӳ��</td>
	<td nowrap width="80">�w������</td>
	<td nowrap width="100">�D�D�n�w����</td>
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
			<td nowrap><%=rs("timeflag")%></td>
			<td nowrap><%=rs("slevel")%></td>
			<td nowrap><%=rs("department")%></td>
			<td nowrap><%=rs("grade")%></td>
			<td nowrap><%=rs("sid")%> - <%=rs("name")%></td>
			<td nowrap><%=rs("score")%></td>
			<td nowrap><%=replace(replace(rs("category"),"ST","�p�Ѯv"),"T","�Юv")%></td>
			<td nowrap><%=rs("teachername")%></td>
			<td nowrap><%=rs("deptgroup")%>&nbsp;</td>
			<td nowrap><%=rs("item")%></td>
			<td nowrap><%if rs("pid")<>"" then response.write "Y" else response.write "&nbsp;" end if%></td>
			
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
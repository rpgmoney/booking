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
'�Ҧ��X�p
sql = "select  sum(cnt) as totalcntall  from "
sql = sql & " ( "
'�Юv�������{�X�p
sql = sql & "select count(*) as cnt from boo_book_T_M "
sql = sql & " where  yn='Y' and signin is not null "
if sdate<>"" and edate="" then
	sql = sql & " and  Cast(bdate as int) >= " & sdate& "  "
end if
if sdate="" and edate<>"" then
	sql = sql & " and  Cast(bdate as int)<= " & edate& "  "
end if
if sdate<>"" and edate<>"" then
	sql = sql & " and (Cast(bdate as int) >= " & sdate & "   and Cast(bdate as int)<= " & edate& " )"
end if
if slevel<>"" then
	sql = sql & " and slevel='"&slevel&"'"
end if
if department<>"" then
	sql = sql & " and department='"&department&"'"
end if
sql = sql & " union "
'�۾����{�X�p
sql = sql & "select count(*) as cnt from boo_book_software "
sql = sql & " where  yn='Y' and signin is not null "
if sdate<>"" and edate="" then
	sql = sql & " and Cast(bdate as int) >= " & sdate& "  "
end if
if sdate="" and edate<>"" then
	sql = sql & " and  Cast(bdate as int)<= " & edate& "  "
end if
if sdate<>"" and edate<>"" then
	sql = sql & " and (Cast(bdate as int)>= " & sdate & "   and Cast(bdate as int)<= " & edate& " )"
end if
if slevel<>"" then
	sql = sql & " and slevel='"&slevel&"'"
end if
if department<>"" then
	sql = sql & " and department='"&department&"'"
end if
sql = sql & " union "
'�u�@�{���{�X�p
sql = sql & "select count(*) as cnt from boo_book_lecture  a inner join boo_lecture b on a.lid=b.id "
sql = sql & " where  a.yn='Y' and a.signin is not null "
if sdate<>"" and edate="" then
	sql = sql & " and Cast(b.date1 as int)  >= " & sdate& "  "
end if
if sdate="" and edate<>"" then
	sql = sql & " and Cast(b.date1 as int) <= " & edate& "  "
end if
if sdate<>"" and edate<>"" then
	sql = sql & " and ( Cast(b.date1 as int)>= " & sdate & "   and  Cast(b.date1 as int)<= " & edate& " ) "
end if
if slevel<>"" then
	sql = sql & " and a.slevel='"&slevel&"'"
end if
if department<>"" then
	sql = sql & " and a.department='"&department&"' "
end if
sql = sql & " )  a "


'response.write sql

rs.Open sql,msconn,adOpenStatic,adLockReadonly
if not rs.EOF then
	totalcntall = rs("totalcntall")
else
	totalcntall=0
end if
rs.close

%>
<HTML>
<HEAD>
<TITLE> �iLDCC�^�~�y��O�E�_���ɤ��ߡj  </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>

</HEAD>
<BODY bottomMargin=0 leftMargin=2 topMargin=0 rightMargin=2 marginheight="0" marginwidth="0"  >


<P align="center" class="inputlabel"><font size="4">�ϥΥت��έp��</font></P>

<TABLE cellSpacing=1 cellPadding=3 align="center" width="700" border=1 >
<TR class="inputlabel" bgcolor="#E7E7E7">
	<td nowrap width="30">&nbsp;</td>
	<td nowrap width="200">����</td>
	<td nowrap width="100">�ϥΤH����</td>
	<td nowrap width="100">���{���ؤ��</td>
	<td nowrap width="100">�U�����</td>
</TR>
<%
'<!-- ----------------�Юv�������{--------------------------------- -->
sql = " select a.name,isnull(b.cnt,0) as cnt  from  "
sql = sql & " ( "
sql = sql & " select * from boo_slevel where flag='T' "
sql = sql & " ) a left join  "
sql = sql & " ( "
sql = sql & "select item,count(*) as cnt from boo_book_T_M "
sql = sql & " where  yn='Y'  and signin is not null "
if sdate<>"" and edate="" then
	sql = sql & " and Cast(bdate as int) >= " & sdate& "  "
end if
if sdate="" and edate<>"" then
	sql = sql & " and  Cast(bdate as int)<= " & edate& "  "
end if
if sdate<>"" and edate<>"" then
	sql = sql & " and (Cast(bdate as int)>= " & sdate & "   and Cast(bdate as int)<= " & edate& " )"
end if
if slevel<>"" then
	sql = sql & " and slevel='"&slevel&"'"
end if
if department<>"" then
	sql = sql & " and department='"&department&"'"
end if
sql = sql & " group by item"
sql = sql & ") b on a.name=b.item order by a.seq"
'response.write sql
'response.end
rs.Open sql,msconn,adOpenStatic,adLockReadonly
	if not rs.eof then 
%>
<TR bgcolor="#E0F7DD"><TD colspan="5" align="center">�Юv�������{</TD></TR>
<%
		'�Юv�������{�X�p
		sql = "select count(*) as cnt from boo_book_T_M "
		sql = sql & " where  yn='Y' and signin is not null "
		if sdate<>"" and edate="" then
			sql = sql & " and Cast(bdate as int) >= " & sdate& "  "
		end if
		if sdate="" and edate<>"" then
			sql = sql & " and Cast(bdate as int)<= " & edate& "  "
		end if
		if sdate<>"" and edate<>"" then
			sql = sql & " and (Cast(bdate as int)>= " & sdate & "   and  Cast(bdate as int)<= " & edate& " )"
		end if
		if slevel<>"" then
			sql = sql & " and slevel='"&slevel&"'"
		end if
		if department<>"" then
			sql = sql & " and department='"&department&"'"
		end if
		'response.write sql & "<br>"
		rs2.Open sql,msconn,adOpenStatic,adLockReadonly
		if  not rs2.Eof then
			totalsum = rs2("cnt")
		end if
		rs2.Close
		while not rs.EOF 
			rc=rc +1
		%>
		<tr >
			<td nowrap><%=rc%></td>
			<td nowrap><%=rs("name")%></td>
			<td nowrap align="right"><%=rs("cnt")%></td>
			<td nowrap align="right">
			<%
				if totalsum>0 then
					response.write round(rs("cnt")/totalsum,4)*100
				else
					response.write "0"
				end if
			%>
			%</td>
			<td align="right">
			<%
				if totalsum>0 then
					response.write round(rs("cnt")/totalcntall,4)*100
				else
					response.write "0"
				end if
			%>
			%</td>
		</tr>
	<%	
			rs.movenext
		wend
	%>
		<tr>
		<td>&nbsp;</td><td>�X�p</td><td align="right"><%=totalsum%></td><td align="right">100%</td>
		<td align="right" bgcolor="#FFCC66">
		<%
			if totalsum>0 then
				response.write round(totalsum/totalcntall,4)*100
			else
				response.write "0"
			end if
		%>
		%</td>
		</tr>
<%
	else
%>
		<TR ><TD colspan="6" align="center"><FONT color=gray>�S���ŦX���󪺸�����</FONT></TD></TR>
<%
	end if
	rs.close
'<!-- ----------------�Юv�������{end--------------------------------- -->
%>
<%
'<!-- ----------�۾����{--------------------- -->
sql = " select a.name,isnull(b.cnt,0) as cnt  from  "
sql = sql & " ( "
sql = sql & " select name,flag2,seq from boo_slevel where flag='A' "
sql = sql & " ) a left join  "
sql = sql & " ( "
sql = sql & "select category,count(*) as cnt from boo_book_software "
sql = sql & " where  yn='Y'  and signin is not null "
if sdate<>"" and edate="" then
	sql = sql & " and  Cast(bdate as int) >= " & sdate& "  "
end if
if sdate="" and edate<>"" then
	sql = sql & " and  Cast(bdate as int) <= " & edate& "  "
end if
if sdate<>"" and edate<>"" then
	sql = sql & " and (Cast(bdate as int) >= " & sdate & "   and Cast(bdate as int) <= " & edate& " )"
end if

if slevel<>"" then
	sql = sql & " and slevel='"&slevel&"'"
end if
if department<>"" then
	sql = sql & " and department='"&department&"'"
end if
sql = sql & " group by category"
sql = sql & ") b on a.flag2=b.category order by a.seq"
'response.write sql
'response.end
rs.Open sql,msconn,adOpenStatic,adLockReadonly
	if not rs.eof then 
%>
<TR bgcolor="#E0F7DD"><TD colspan="5" align="center">�۾����{</TD></TR>
<%
		'�۾����{�X�p
		sql = "select count(*) as cnt from boo_book_software "
		sql = sql & " where  yn='Y' and signin is not null "
		if sdate<>"" and edate="" then
			sql = sql & " and  Cast(bdate as int) >= " & sdate& "  "
		end if
		if sdate="" and edate<>"" then
			sql = sql & " and Cast(bdate as int) <= " & edate& "  "
		end if
		if sdate<>"" and edate<>"" then
			sql = sql & " and ( Cast(bdate as int) >= " & sdate & "   and  Cast(bdate as int) <= " & edate& " )"
		end if
		if slevel<>"" then
			sql = sql & " and slevel='"&slevel&"'"
		end if
		if department<>"" then
			sql = sql & " and department='"&department&"'"
		end if
		'response.write sql & "<br>"
		rs2.Open sql,msconn,adOpenStatic,adLockReadonly
		if  not rs2.Eof then
			totalsum = rs2("cnt")
		end if
		rs2.Close
		while not rs.EOF 
			rc=rc +1
		%>
		<tr >
			<td nowrap><%=rc%></td>
			<td nowrap><%=rs("name")%></td>
			<td nowrap align="right"><%=rs("cnt")%></td>
			<td nowrap align="right">
			<%
				if totalsum>0 then
					response.write round(rs("cnt")/totalsum,4)*100
				else
					response.write "0"
				end if
			%>
			%</td>
			<td align="right">
			<%
				if totalsum>0 then
					response.write round(rs("cnt")/totalcntall,4)*100
				else
					response.write "0"
				end if
			%>
			%</td>
		</tr>
	<%	
			rs.movenext
		wend
	%>
		<tr>
		<td>&nbsp;</td><td>�X�p</td><td align="right"><%=totalsum%></td><td align="right">100%</td>
		<td align="right" bgcolor="#FFCC66">
		<%
			if totalsum>0 then
				response.write round(totalsum/totalcntall,4)*100
			else
				response.write "0"
			end if
		%>
		%</td>
		</tr>
<%
	else
%>
		<TR ><TD colspan="6" align="center"><FONT color=gray>�S���ŦX���󪺸�����</FONT></TD></TR>
<%
	end if
	rs.close
'<!-- ----------�۾����{end--------------------- -->
%>
<%
'<!-- ----------�u�@�{���{--------------------- -->
sql = " select a.name,isnull(b.cnt,0) as cnt  from  "
sql = sql & " ( "
sql = sql & " select name,flag2,seq from boo_slevel where flag='B' "
sql = sql & " ) a left join  "
sql = sql & " ( "
sql = sql & "select a.category,count(*) as cnt from boo_book_lecture a inner join boo_lecture b on a.lid=b.id "
sql = sql & " where  yn='Y'  and signin is not null "
if sdate<>"" and edate="" then
	sql = sql & " and Cast(date1 as int) >= " & sdate& "  "
end if
if sdate="" and edate<>"" then
	sql = sql & " and Cast(date1 as int) <= " & edate& " "
end if
if sdate<>"" and edate<>"" then
	sql = sql & " and ( Cast(date1 as int) >= " & sdate & "  and Cast(date1 as int) <= " & edate& " )"
end if
if slevel<>"" then
	sql = sql & " and slevel='"&slevel&"'"
end if
if department<>"" then
	sql = sql & " and department='"&department&"'"
end if
sql = sql & " group by a.category"
sql = sql & ") b on a.flag2=b.category order by a.seq"
'response.write sql
'response.end
rs.Open sql,msconn,adOpenStatic,adLockReadonly
	if not rs.eof then 
%>
<TR bgcolor="#E0F7DD"><TD colspan="5" align="center">�u�@�{���{</TD></TR>
<%
		'�u�@�{���{�X�p
		sql = "select count(*) as cnt from boo_book_lecture  a inner join boo_lecture b on a.lid=b.id "
		sql = sql & " where  a.yn='Y' and a.signin is not null "
		if sdate<>"" and edate="" then
			sql = sql & " and  Cast(b.date1 as int) >= " & sdate& "  "
		end if
		if sdate="" and edate<>"" then
			sql = sql & " and  Cast(b.date1 as int) <= " & edate& "  "
		end if
		if sdate<>"" and edate<>"" then
			sql = sql & " and ( Cast(b.date1 as int)>= " & sdate & "   and Cast( b.date1 as int)<= " & edate& ")"
		end if
		if slevel<>"" then
			sql = sql & " and a.slevel='"&slevel&"'"
		end if
		if department<>"" then
			sql = sql & " and a.department='"&department&"'"
		end if
		'response.write sql & "<br>"
		rs2.Open sql,msconn,adOpenStatic,adLockReadonly
		if  not rs2.Eof then
			totalsum = rs2("cnt")
		end if
		rs2.Close
		while not rs.EOF 
			rc=rc +1
		%>
		<tr >
			<td nowrap><%=rc%></td>
			<td nowrap><%=rs("name")%></td>
			<td nowrap align="right"><%=rs("cnt")%></td>
			<td nowrap align="right">
			<%
				if totalsum>0 then
					response.write round(rs("cnt")/totalsum,4)*100
				else
					response.write "0"
				end if
			%>
			%</td>
			<td align="right">
			<%
				if totalsum>0 then
					response.write round(rs("cnt")/totalcntall,4)*100
				else
					response.write "0"
				end if
			%>
			%</td>
		</tr>
	<%	
			rs.movenext
		wend
	%>
		<tr>
		<td>&nbsp;</td><td>�X�p</td><td align="right"><%=totalsum%></td><td align="right">100%</td>
		<td align="right" bgcolor="#FFCC66">
		<%
			if totalsum>0 then
				response.write round(totalsum/totalcntall,4)*100
			else
				response.write "0"
			end if
		%>
		%</td>
		</tr>
<%
	else
%>
		<TR ><TD colspan="6" align="center"><FONT color=gray>�S���ŦX���󪺸�����</FONT></TD></TR>
<%
	end if
	rs.close
'<!-- ----------�u�@�{���{end--------------------- -->
%>
</table>
<P><BR>
</BODY>
</HTML>
<!-- #INCLUDE virtual="/include/lib/conn/syconnclose.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->

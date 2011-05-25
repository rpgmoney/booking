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
'所有合計
sql = "select  sum(cnt) as totalcntall  from "
sql = sql & " ( "
'教師輔導療程合計
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
'自學療程合計
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
'工作坊療程合計
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
<TITLE> 【LDCC英外語能力診斷輔導中心】  </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>

</HEAD>
<BODY bottomMargin=0 leftMargin=2 topMargin=0 rightMargin=2 marginheight="0" marginwidth="0"  >


<P align="center" class="inputlabel"><font size="4">使用目的統計表</font></P>

<TABLE cellSpacing=1 cellPadding=3 align="center" width="700" border=1 >
<TR class="inputlabel" bgcolor="#E7E7E7">
	<td nowrap width="30">&nbsp;</td>
	<td nowrap width="200">項目</td>
	<td nowrap width="100">使用人次數</td>
	<td nowrap width="100">療程項目比例</td>
	<td nowrap width="100">各項比例</td>
</TR>
<%
'<!-- ----------------教師輔導療程--------------------------------- -->
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
<TR bgcolor="#E0F7DD"><TD colspan="5" align="center">教師輔導療程</TD></TR>
<%
		'教師輔導療程合計
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
		<td>&nbsp;</td><td>合計</td><td align="right"><%=totalsum%></td><td align="right">100%</td>
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
		<TR ><TD colspan="6" align="center"><FONT color=gray>沒有符合條件的資料顯示</FONT></TD></TR>
<%
	end if
	rs.close
'<!-- ----------------教師輔導療程end--------------------------------- -->
%>
<%
'<!-- ----------自學療程--------------------- -->
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
<TR bgcolor="#E0F7DD"><TD colspan="5" align="center">自學療程</TD></TR>
<%
		'自學療程合計
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
		<td>&nbsp;</td><td>合計</td><td align="right"><%=totalsum%></td><td align="right">100%</td>
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
		<TR ><TD colspan="6" align="center"><FONT color=gray>沒有符合條件的資料顯示</FONT></TD></TR>
<%
	end if
	rs.close
'<!-- ----------自學療程end--------------------- -->
%>
<%
'<!-- ----------工作坊療程--------------------- -->
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
<TR bgcolor="#E0F7DD"><TD colspan="5" align="center">工作坊療程</TD></TR>
<%
		'工作坊療程合計
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
		<td>&nbsp;</td><td>合計</td><td align="right"><%=totalsum%></td><td align="right">100%</td>
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
		<TR ><TD colspan="6" align="center"><FONT color=gray>沒有符合條件的資料顯示</FONT></TD></TR>
<%
	end if
	rs.close
'<!-- ----------工作坊療程end--------------------- -->
%>
</table>
<P><BR>
</BODY>
</HTML>
<!-- #INCLUDE virtual="/include/lib/conn/syconnclose.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->

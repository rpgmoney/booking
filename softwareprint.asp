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

sql = " select a.floor,a.software,isnull(b.cnt,0) as cnt  from  "
sql = sql & " ( "
sql = sql & " select id,floor,software from boo_software where yn='Y'  "
sql = sql &  " union "
sql = sql &	 " select  id,'自學' as floor,item as software  from boo_self_item where yn='Y'  "
sql = sql & " ) a left join  "
sql = sql & " ( "
sql = sql & "select item,count(*) as cnt from boo_book_software "
sql = sql & " where  yn='Y'  and signin is not null "
if sdate<>"" and edate="" then
	sql = sql & " and Cast(bdate as int) >= " & sdate& "  "
end if
if sdate="" and edate<>"" then
	sql = sql & " and  Cast(bdate as int) <= " & edate& "  "
end if
if sdate<>"" and edate<>"" then
	sql = sql & " and ( Cast(bdate as int) >= " & sdate & "  and Cast(bdate as int) <= " & edate& " )"
end if

if slevel<>"" then
	sql = sql & " and slevel='"&slevel&"'"
end if
if department<>"" then
	sql = sql & " and department='"&department&"'"
end if
sql = sql & " group by item"
sql = sql & ") b on a.id=b.item order by a.floor,a.software"


'response.write sql
'response.end
rs.Open sql,msconn,adOpenStatic,adLockReadonly


'rs.close
'set rs=nothing
%>
<HTML>
<HEAD>
<TITLE> 【LDCC英外語能力診斷輔導中心】  </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>

</HEAD>
<BODY bottomMargin=0 leftMargin=2 topMargin=0 rightMargin=2 marginheight="0" marginwidth="0"  >

<%
	if not rs.eof then 
%>
<P align="center" class="inputlabel"><font size="4">自學療程使用分析表</font></P>

<TABLE cellSpacing=1 cellPadding=2 align="center" border=1 >
<TR class="inputlabel" bgcolor="#E7E7E7">
	<td nowrap>&nbsp;</td>
	<td nowrap>自學療程項目</td>
	<td nowrap >使用人次數</td>
	<td nowrap >百分比</td>
	<td nowrap>&nbsp;</td>
</TR>
<%
		'合計
		sql = "select count(*) as cnt from boo_book_software "
		sql = sql & " where  yn='Y'  and signin is not null "
		if sdate<>"" and edate="" then
			sql = sql & " and Cast(bdate as int) >= " & sdate& "  "
		end if
		if sdate="" and edate<>"" then
			sql = sql & " and  Cast(bdate as int)<= " & edate& "  "
		end if
		if sdate<>"" and edate<>"" then
			sql = sql & " and ( Cast(bdate as int) >= " & sdate & "   and Cast(bdate as int)<= " & edate& " )"
		end if

		if slevel<>"" then
			sql = sql & " and slevel='"&slevel&"'"
		end if
		if department<>"" then
			sql = sql & " and department='"&department&"'"
		end if
		rs2.Open sql,msconn,adOpenStatic,adLockReadonly
		if  not rs2.Eof then
			totalsum = rs2("cnt")
		end if
		rs2.Close

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
			<td nowrap><%=rs("floor")%> - <%=rs("software")%></td>
			<td nowrap><%=rs("cnt")%></td>
			<td nowrap><% if totalsum > 1 then response.write   round(rs("cnt")/totalsum,4)*100 else response.write "0"  end if %>%</td>
			<td><img src="images/bar.gif" height="20" width="<%=cdbl(rs("cnt"))%>" ></td>
		</tr>
			
	<%	
			rs.movenext
		wend
	%>
		<tr>
		<td>&nbsp;</td><td>合計</td><td><%=totalsum%></td><td>100%</td>
		<td><img src="images/bar.gif" height="20" width="100" ></td>
		</tr>
</table>
	

<%
	else
		Response.Write "<FONT class=normal><FONT color=gray>- 沒有符合條件的資料顯示 -</FONT></FONT>"
	end if
%>



<P><BR>
</BODY>
</HTML>
<!-- #INCLUDE virtual="/include/lib/conn/syconnclose.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->
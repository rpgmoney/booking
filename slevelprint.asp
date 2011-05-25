<%@Language=VBScript LCID=1033%>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->

<%

sdate=trim(request("sdate"))
edate=trim(request("edate"))

scope=trim(request("scope"))
item=replace(trim(request("item")),", ","','")

set rs=server.CreateObject("adodb.recordset")
set rs2=server.CreateObject("adodb.recordset")

if scope<>"" then
	sql = ""
	if instr(scope,"1") <>0 then
		sql = sql &  " select slevel   from boo_book_T_M  where yn='Y' and signin is not null "
		if sdate<>"" and edate="" then
			sql = sql & " and Cast(bdate as int) >= '" & sdate& "'  "
		end if
		if sdate="" and edate<>"" then
			sql = sql & " and  Cast(bdate as int) <= '" & edate& "'  "
		end if
		if sdate<>"" and edate<>"" then
			sql = sql & " and ( Cast(bdate as int) >= '" & sdate & "'   and Cast(bdate as int) <= '" & edate& "' )"
		end if
		if item<>"" then
			sql = sql & " and item in ( '" & item &"' ) "
		end if
	end if
	if instr(scope,"2") <>0 then
		if sql<>"" then sql=sql & " union all " end if
		sql = sql &  " select slevel  from boo_book_software  where yn='Y' and signin is not null "
		if sdate<>"" and edate="" then
			sql = sql & " and  Cast( bdate as int) >= '" & sdate& "'  "
		end if
		if sdate="" and edate<>"" then
			sql = sql & " and  Cast(bdate as int) <= '" & edate& "'  "
		end if
		if sdate<>"" and edate<>"" then
			sql = sql & " and ( Cast(bdate as int) >= '" & sdate & "'   and Cast(bdate as int) <= '" & edate& "' )"
		end if
	end if
	
	if instr(scope,"3") <>0 then
		if sql<>"" then sql=sql & " union all " end if
		sql = sql &  " select slevel  from boo_book_lecture a inner join boo_lecture b on a.lid=b.id where yn='Y' and signin is not null "
		if sdate<>"" and edate="" then
			sql = sql & " and  Cast(date1 as int) >= '" & sdate& "'  "
		end if
		if sdate="" and edate<>"" then
			sql = sql & " and Cast( date1 as int) <= '" & edate& "'  "
		end if
		if sdate<>"" and edate<>"" then
			sql = sql & " and ( Cast(date1 as int) >= '" & sdate & "'   and Cast(date1 as int) <= '" & edate& "' )"
		end if

	end if
	if sql<>"" then
		sqlc =  "select a.name,isnull(b.cnt,0) as cnt from boo_slevel a left join "
		sqlc = sqlc & " ("
		sqlc = sqlc  & "select slevel ,count(*) as cnt  from ( " & sql & " ) a group by slevel "
		sqlc = sqlc & " ) b on a.name=b.slevel where a.flag='S' order by a.seq"
	end if

else
	sqlc = "select * from boo_book_T_M where 1=0 "
end if


'response.write sqlc
'response.end
rs.Open sqlc,msconn,adOpenStatic,adLockReadonly


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
<P align="center" class="inputlabel"><font size="4">學制人數統計分析表</font></P>

<TABLE cellSpacing=1 cellPadding=2 align="center" border=1 >
<TR class="inputlabel" bgcolor="#E7E7E7">
	<td nowrap>&nbsp;</td>
	<td nowrap>學制</td>
	<td nowrap >使用人次數</td>
	<td nowrap >百分比</td>
	<td nowrap>&nbsp;</td>
</TR>
<%
		'合計
		sqls = ""
		if instr(scope,"1") <>0 then
			sqls = sqls &  " select id   from boo_book_T_M  where yn='Y' and signin is not null "
			if sdate<>"" and edate="" then
				sqls = sqls & " and  Cast(bdate as int)  >= '" & sdate& "'  "
			end if
			if sdate="" and edate<>"" then
				sqls = sqls & " and  Cast(bdate as int) <= '" & edate& "'  "
			end if
			if sdate<>"" and edate<>"" then
				sqls = sqls & " and ( Cast(bdate as int) >= '" & sdate & "'   and Cast(bdate as int) <= '" & edate& "' )"
			end if
			if item<>"" then
				sql = sql & " and item in ('"&item&"')"
			end if
		end if
		if instr(scope,"2") <>0 then
			if sqls<>"" then sqls=sqls & " union all " end if
			sqls = sqls &  " select id  from boo_book_software  where yn='Y' and signin is not null "
			if sdate<>"" and edate="" then
				sqls = sqls & " and  Cast(bdate as int)  >= '" & sdate& "'  "
			end if
			if sdate="" and edate<>"" then
				sqls = sqls & " and  Cast(bdate as int) <= '" & edate& "'  "
			end if
			if sdate<>"" and edate<>"" then
				sqls = sqls & " and ( Cast(bdate as int) >= '" & sdate & "'   and  Cast(bdate as int) <= '" & edate& "' )"
			end if
		end if
		if instr(scope,"3") <>0 then
			if sqls<>"" then sqls=sqls & " union all " end if
			sqls = sqls &  " select slevel  from boo_book_lecture a inner join boo_lecture b on a.lid=b.id where yn='Y' and signin is not null "
			if sdate<>"" and edate="" then
				sqls = sqls & " and  Cast(date1 as int)  >= '" & sdate& "'  "
			end if
			if sdate="" and edate<>"" then
				sqls = sqls & " and  Cast(date1 as int) <= '" & edate& "'  "
			end if
			if sdate<>"" and edate<>"" then
				sqls = sqls & " and ( Cast(date1 as int) >= '" & sdate & "'   and Cast(date1 as int) <= '" & edate& "' )"
			end if


		end if
		if sqls<>"" then
			sqlg = sqlg  & "select count(*) as cnt  from ( " & sqls & " ) a "

		end if
'		response.write sqlg
'response.end

		rs2.Open sqlg,msconn,adOpenStatic,adLockReadonly
		if  not rs2.Eof then
			totalsum = rs2("cnt")
		end if
		rs2.Close
		chart=round(300/totalsum,3)'圖形*數
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
			<td nowrap><%=rs("name")%></td>
			<td nowrap><%=rs("cnt")%></td>
			<td nowrap><%=round(rs("cnt")/totalsum,4)*100%>%</td>
			<td><img src="images/bar.gif" height="20" width="<%=cdbl(rs("cnt"))*chart%>" ></td>
		</tr>
			
	<%	
			rs.movenext
		wend
	%>
		<tr>
		<td>&nbsp;</td><td>合計</td><td><%=totalsum%></td><td>100%</td>
		<td><img src="images/bar.gif" height="20" width="<%=totalsum*chart%>" ></td>
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
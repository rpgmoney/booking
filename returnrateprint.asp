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



'select a.sid,a.name,a.slevel,a.grade,a.class1,a.department,a.score,b.backdate,a.bdate,b.notice,b.tid  
' from boo_book_T_M a inner join boo_diagnosis b on a.id=b.tid where 1=1  
'and sid   in
' (
'select sid  from boo_book_T_M  c   where c.item='診斷' and c.YN='Y'   
'and  cast(left(c.bdate,2)+1911 as varchar(4)) + '/' + substring(c.backdate,3,2) + '/'+ right(c.backdate,2) as datetime)>dateadd(d,12,cast( cast(left(c.backdate,2)+1911 as varchar(4)) + '/' + substring(c.backdate,3,2) + '/'+ right(c.backdate,2) as datetime))  

' )  order by backdate


'select  backdate,dateadd(d,12,cast( cast(left(backdate,2)+1911 as varchar(4)) + '/' + substring(backdate,3,2) + '/'+ right(backdate,2) as datetime))   from boo_diagnosis



'select *  from boo_diagnosis

%>
<HTML>
<HEAD>
<TITLE> 【LDCC英外語能力診斷輔導中心】  </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>

</HEAD>
<BODY bottomMargin=0 leftMargin=2 topMargin=0 rightMargin=2 marginheight="0" marginwidth="0"  >


<P align="center" class="inputlabel"><font size="4">每月回診比率統計</font></P>


<%

sql = "select a.yymm,a.cnt,b.cnt1 "
sql = sql & " from "
sql = sql & " ( "
'到期應回診人數
		sql = sql & " select left(b.backdate,len(b.backdate)-2) as yymm,count(*) as cnt "
		sql = sql & "  from boo_book_T_M a inner join boo_diagnosis b on a.id=b.tid where 1=1  "
		if sdate<>"" and edate="" then
			sql = sql & " and Cast(backdate as int) >= '" & sdate& "'  "
		end if
		if sdate="" and edate<>"" then
			sql = sql & " and  Cast(backdate as int)<= '" & edate& "'  "
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
		sql = sql & " group by  left(b.backdate,len(b.backdate)-2) "
sql = sql & " ) a left join "
sql = sql & " (  "
'到期已回診人數
sql = sql & "select left(b.backdate,len(b.backdate)-2) as yymm,count(*) as cnt1 "
 sql = sql & "from boo_book_T_M a inner join boo_diagnosis b on a.id=b.tid where 1=1  " 
sql = sql & "and sid   in "
sql = sql & " ( "
sql = sql &		"select sid  from boo_book_T_M  c   where c.item='診斷' and c.YN='Y'    "
sql = sql &		"and a.sid=c.sid "
sql = sql &		"and  cast( cast(left(c.bdate,len(c.bdate)-4)+1911 as varchar(4)) + '/' + substring(c.bdate,len(c.bdate)-3,2) + '/'+ right(c.bdate,2) as  datetime )  > "
sql = sql &		"dateadd(d,-12,cast( cast(left( b.backdate,len(b.backdate)-4)+1911 as varchar(4)) + '/' + substring(b.backdate,len(b.backdate)-3,2) + '/'+ right(b.backdate,2) as datetime))   "
sql = sql &		"and  cast( cast(left( c.bdate,len(c.bdate)-4)+1911 as varchar(4)) + '/' + substring(c.bdate,len(c.bdate)-3,2) + '/'+ right(c.bdate,2) as  datetime )  < "
sql = sql &		"dateadd(d,12,cast( cast(left(  b.backdate,len(b.backdate)-4)+1911 as varchar(4)) + '/' + substring(b.backdate,len(b.backdate)-3,2) + '/'+ right(b.backdate,2) as datetime))  "
sql = sql & " ) "
sql = sql & "group by left(b.backdate,len(b.backdate)-2)   "
sql = sql & " ) b on a.yymm=b.yymm "
'response.write sql
'response.end
rs.Open sql,msconn,adOpenStatic,adLockReadonly
	if not rs.eof then 
		totalcnt = 0
		lableStr=""
		cntStr = ""
		cnt1Str = ""
		cnt2Str = ""
		while not rs.EOF 
			rc=rc +1
			totalcnt = cdbl(totalcnt) + cdbl(ifnull(rs("cnt"),0))
			lableStr = lableStr & "<td align=center>" &  left(rs("yymm"),2) & "&nbsp;年&nbsp;" & right(rs("yymm"),2) & "&nbsp;月&nbsp;" & "</td>"
			'應到診
			cntStr = cntStr & "<td align=right>" & rs("cnt") & "</td>"
			'已回診人數
			cnt1Str = cnt1Str & "<td align=right>" & ifnull(rs("cnt1"),"0") & "</td>"
			'到期未回診人數
			cnt2 = cdbl(rs("cnt") ) - cdbl(ifnull(rs("cnt1"),"0"))
			cnt2Str = cnt2Str & "<td align=right>" & cnt2 & "</td>"
			'回診百分比
			if rs("cnt1") >0 and rs("cnt") >0 then
				cnt3 = round(cdbl(rs("cnt1"))/cdbl(rs("cnt")),3)*100
			else
				cnt3 = 0
			end if
			cnt3Str = cnt3Str & "<td align=right>" & cnt3 & "%</td>"
			'未回診百分比
			if  cnt2 >0 and rs("cnt") >0 then
				cnt5 = round(cdbl(cnt2)/cdbl(rs("cnt")),3)*100
			else
				cnt5 = 0
			end if
			cnt5Str = cnt5Str & "<td align=right>" & cnt5 & "%</td>"
			rs.movenext
		wend
	%>
	<TABLE cellSpacing=1 cellPadding=3 align="center" width="700" border=1 >
	<TR class="inputlabel" bgcolor="#E7E7E7">
		<td nowrap >年月</td>
		<%=lableStr%>
	</TR>
	<TR >
		<td nowrap class="inputlabel" >到期應回診人數</td>
		<%=cntStr%>
	</TR>
	<TR >
		<td nowrap class="inputlabel" >到期已回診人數</td>
		<%=cnt1Str%>
	</TR>
	<TR >
		<td nowrap class="inputlabel" >到期未回診人數</td>
		<%=cnt2Str%>
	</TR>
	<TR >
		<td nowrap class="inputlabel" >回診百分比</td>
		<%=cnt3Str%>
	</TR>
	<TR >
		<td nowrap class="inputlabel" >未回診百分比</td>
		<%=cnt5Str%>
	</TR>
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

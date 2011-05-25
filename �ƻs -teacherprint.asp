<%@Language=VBScript LCID=1033%>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->

<%

sdate=trim(request("sdate"))
edate=trim(request("edate"))
yms=trim(request("yms"))
slevel=trim(request("slevel"))
department=trim(request("department"))



set rs=server.CreateObject("adodb.recordset")
set rs2=server.CreateObject("adodb.recordset")

sql =" select a.group1,a.teachername,c.cnt1,b.cnt,a.alltime from"
sql = sql &" ("
'時數
sql = sql &		" select  a.group1,a.teachername,sum(Cast(a.timeflag as decimal(2,1) )) as alltime    from "
sql = sql &		" ("
sql = sql &				" select b.group1,a.bdate,replace(replace(replace(a.timeflag,'A',1),'U',0.5),'B',0.5) as timeflag  ,teachername "
sql = sql &				" from boo_book_T_M a left join  "
sql = sql &				" ( "
sql = sql &					" select distinct yms,teacher,category,group1  from boo_schedule where yms='"&yms&"' and category='T'  "
sql = sql &				" )  b on a.teachername=b.teacher and a.yms=b.yms where  a.yn='Y' and signin is not null and a.yms='"&yms&"' and a.category='T' "
				if sdate<>"" and edate="" then
					sql = sql & " and  Cast(a.bdate as int) >= '" & sdate& "'  "
				end if
				if sdate="" and edate<>"" then
					sql = sql & " and  Cast(a.bdate as int)<= '" & edate& "'  "
				end if
				if sdate<>"" and edate<>"" then
					sql = sql & " and ( Cast(a.bdate as int)>= '" & sdate & "'   and Cast(a.bdate as int)<= '" & edate& "' )"
				end if
				if slevel<>"" then
					sql = sql & " and a.slevel='"&slevel&"'"
				end if
				if department<>"" then
					sql = sql & " and a.department='"&department&"'"
				end if
sql = sql &		" ) a group by a.group1,a.teachername  "
sql = sql &  ") a inner  join "

sql = sql &" ("
'人次數
sql = sql &		" select  a.group1,a.teachername,count(*)  cnt   from "
sql = sql &		" ("
sql = sql &				" select b.group1,a.bdate,teachername,sid "
sql = sql &				" from boo_book_T_M a left join  "
sql = sql &				" ( "
sql = sql &				" select distinct yms,teacher,category,group1  from boo_schedule where yms='"&yms&"' and category='T'  "
sql = sql &				" )  b on a.teachername=b.teacher and a.yms=b.yms where  a.yn='Y' and signin is not null and a.yms='"&yms&"' and a.category='T' "
				if sdate<>"" and edate="" then
					sql = sql & " and  Cast(a.bdate as int) >= '" & sdate& "'  "
				end if
				if sdate="" and edate<>"" then
					sql = sql & " and  Cast(a.bdate as int)<= '" & edate& "'  "
				end if
				if sdate<>"" and edate<>"" then
					sql = sql & " and ( Cast(a.bdate as int)>= '" & sdate & "'   and  Cast(a.bdate as int) <= '" & edate& "' )"
				end if
				if slevel<>"" then
					sql = sql & " and a.slevel='"&slevel&"'"
				end if
				if department<>"" then
					sql = sql & " and a.department='"&department&"'"
				end if
sql = sql &		 " ) a group by a.group1,a.teachername  "
sql = sql &	 " ) b on a.group1=b.group1 and a.teachername=b.teachername "
sql = sql &  " inner  join "
sql = sql & " ( "
'人數
sql = sql &		" select  a.group1,a.teachername,count(*)  cnt1   from "
sql = sql &		" ("
sql = sql &				" select distinct b.group1,teachername,sid "
sql = sql &				" from boo_book_T_M a left join  "
sql = sql &				" ( "
sql = sql &				" select distinct yms,teacher,category,group1  from boo_schedule where yms='"&yms&"' and category='T'  "
sql = sql &				" )  b on a.teachername=b.teacher and a.yms=b.yms where  a.yn='Y' and signin is not null and a.yms='"&yms&"' and a.category='T' "
				if sdate<>"" and edate="" then
					sql = sql & " and  Cast(a.bdate as int) >= '" & sdate& "'  "
				end if
				if sdate="" and edate<>"" then
					sql = sql & " and Cast(a.bdate as int)<= '" & edate& "'  "
				end if
				if sdate<>"" and edate<>"" then
					sql = sql & " and ( Cast(a.bdate as int)>= '" & sdate & "'   and Cast(a.bdate as int)<= '" & edate& "' )"
				end if
				if slevel<>"" then
					sql = sql & " and a.slevel='"&slevel&"'"
				end if
				if department<>"" then
					sql = sql & " and a.department='"&department&"'"
				end if
sql = sql &		 " ) a group by a.group1,a.teachername  "
sql = sql &	 ") c on a.group1=c.group1 and a.teachername=c.teachername "
sql = sql & " order by a.group1,a.teachername "

response.write sql 
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
<P align="center" class="inputlabel"><font size="4">專案教師總統計</font></P>

<TABLE cellSpacing=1 cellPadding=10 align="center" border=1 >
<TR class="inputlabel" bgcolor="#E7E7E7">
	<td nowrap>&nbsp;</td>
	<td nowrap>系科</td>
	<td nowrap>教師</td>
	<td nowrap >受輔學生人數</td>
	<td nowrap >輔導次數</td>
	<td nowrap>輔導時數</td>
</TR>
<%
		
		alltime_T = 0
		cnt1_T = 0
		cnt_T = 0
		tmpgroup=""
		while not rs.EOF 
			rc=rc +1
			alltime_T = cdbl(alltime_T) + cdbl(rs("alltime"))
			cnt1_T = cdbl(cnt1_T) + cdbl(rs("cnt1"))
			cnt_T = cdbl(cnt_T) + cdbl(rs("cnt"))
			if rc mod 2 = cint(0) then
				vcolor="#E0F7DD"
			else
				vcolor="#FFFFFF"
			end if
		%>
		<tr bgcolor="<%=vcolor%>">
			<td nowrap><%=rc%></td>
			<%if  tmpgroup <> rs("group1") then%>
			<td nowrap><%=rs("group1")%></td>
			<%else%>
			<td nowrap>&nbsp;</td>
			<%end if%>
			<td nowrap><%=rs("teachername")%></td>
			<td nowrap align="right"><%=rs("cnt1")%></td>
			<td nowrap align="right"><%=rs("cnt")%></td>
			<td nowrap align="right"><%=rs("alltime")%></td>
		</tr>
			
	<%	
			tmpgroup = rs("group1")
			rs.movenext
		wend
	%>
		<tr>
		<td>&nbsp;</td><td colspan="2">合計</td><td align="right"><%=cnt1_T%></td><td align="right"><%=cnt_T%></td>
		<td align="right"><%=alltime_T%></td>
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
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

function ReportHeader()
	dim rs0,sql
	set rs0 = server.CreateObject("adodb.recordset")
	on error resume next
	Str="<TABLE cellSpacing=0 cellPadding=3 border=1   ><TR class=""inputlabel""><TD rowspan=2>&nbsp;</TD><TD rowspan=2>系科</TD><TD rowspan=2>教師</TD><TD rowspan=2>&nbsp;</TD><TD colspan=7 align=center>LDCC英外語能力診斷輔導中心</TD><TD colspan=7 align=center>ELC英語學習中心</TD><TD rowspan=2>合計</TD>"
	
	Str = Str & "</TR>"
	Str = Str & "</TR><TD>診斷</TD><TD>諮商</TD><TD>口語</TD><TD>簡報</TD><TD>詩歌</TD><TD>寫作</TD><TD>閱讀</TD><TD>診斷</TD><TD>諮商</TD><TD>口語</TD><TD>簡報</TD><TD>詩歌</TD><TD>寫作</TD><TD>閱讀</TD>"
	Str = Str & "</TR>"

	sql = "select distinct a.teachername,b.group1  from boo_book_T_M a left join  "
	sql = sql &				" ( "
	sql = sql &					" select distinct yms,teacher,category,group1,deptgroup   from boo_schedule where yms='"&yms&"' and category='T'  "
	sql = sql &				" )  b on a.teachername=b.teacher and a.yms=b.yms where  b.deptgroup is not null and  a.yn='Y' and signin is not null and a.yms='"&yms&"' and a.category='T' "
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
	sql = sql &  " order by b.group1 " 

'		response.write sql
	rs0.Open sql,msconn,adOpenStatic,adLockReadonly
	while not rs0.EOF
		i = i + 1 
		Str = Str & "<tr nowrap><td rowspan=3>"&i&"</td><td rowspan=3 nowrap>"&rs0("group1")&"</td><td rowspan=3 nowrap>"&rs0("teachername")&"</td><td nowrap>受輔學生人數</td>"
		Str = Str & "<TD >:L"&rs0("teachername")&"人診斷</TD><TD >:L"&rs0("teachername")&"人諮商</TD><TD>:L"&rs0("teachername")&"人口語</TD><TD>:L"&rs0("teachername")&"人簡報</TD><TD>:L"&rs0("teachername")&"人詩歌</TD><TD>:L"&rs0("teachername")&"人寫作</TD><TD>:L"&rs0("teachername")&"人閱讀</TD><TD>:E"&rs0("teachername")&"人診斷</TD><TD>:E"&rs0("teachername")&"人諮商</TD><TD>:E"&rs0("teachername")&"人口語</TD><TD>:E"&rs0("teachername")&"人簡報</TD><TD>:E"&rs0("teachername")&"人詩歌</TD><TD>:E"&rs0("teachername")&"人寫作</TD><TD>:E"&rs0("teachername")&"人閱讀</TD><TD>:"&rs0("teachername")&"人合計</TD>"
		Str = Str & "</tr>"
		Str = Str & "<tr nowrap><td>輔導次數</td>"
		Str = Str & "<TD>:L"&rs0("teachername")&"次診斷</TD><TD>:L"&rs0("teachername")&"次諮商</TD><TD>:L"&rs0("teachername")&"次口語</TD><TD>:L"&rs0("teachername")&"次簡報</TD><TD>:L"&rs0("teachername")&"次詩歌</TD><TD>:L"&rs0("teachername")&"次寫作</TD><TD>:L"&rs0("teachername")&"次閱讀</TD><TD>:E"&rs0("teachername")&"次診斷</TD><TD>:E"&rs0("teachername")&"次諮商</TD><TD>:E"&rs0("teachername")&"次口語</TD><TD>:E"&rs0("teachername")&"次簡報</TD><TD>:E"&rs0("teachername")&"次詩歌</TD><TD>:E"&rs0("teachername")&"次寫作</TD><TD>:E"&rs0("teachername")&"次閱讀</TD><TD>:"&rs0("teachername")&"次合計</TD>"
		Str = Str & "</tr>"
		Str = Str & "<tr ><td>輔導時數</td>"
		Str = Str & "<TD nowrap>:L"&rs0("teachername")&"時診斷</TD><TD nowrap>:L"&rs0("teachername")&"時諮商</TD><TD nowrap>:L"&rs0("teachername")&"時口語</TD><TD nowrap>:L"&rs0("teachername")&"時簡報</TD><TD nowrap>:L"&rs0("teachername")&"時詩歌</TD><TD nowrap>:L"&rs0("teachername")&"時寫作</TD><TD nowrap>:L"&rs0("teachername")&"時閱讀</TD><TD nowrap>:E"&rs0("teachername")&"時診斷</TD><TD nowrap>:E"&rs0("teachername")&"時諮商</TD><TD nowrap>:E"&rs0("teachername")&"時口語</TD><TD nowrap>:E"&rs0("teachername")&"時簡報</TD><TD nowrap>:E"&rs0("teachername")&"時詩歌</TD><TD nowrap>:E"&rs0("teachername")&"時寫作</TD><TD nowrap>:E"&rs0("teachername")&"時閱讀</TD><TD nowrap>:"&rs0("teachername")&"時合計</TD>"
		Str = Str & "</tr>"

		rs0.MoveNext
	wend



	Str = Str & "</TABLE>"
	ReportHeader=Str
	
end function



set rs=server.CreateObject("adodb.recordset")
set rs2=server.CreateObject("adodb.recordset")



'人時
sql = "select a.group1,a.deptgroup,a.teachername,a.item,sum(Cast(a.timeflag as decimal(2,1) )) as alltime from "
sql =  sql   &   " ( "
sql =  sql   &   "	 select b.group1,b.deptgroup,a.bdate,replace(replace(replace(a.timeflag,'A',1),'U',0.5),'B',0.5) as timeflag ,teachername,item from boo_book_T_M a "
sql =  sql   &   "	 left join  "
sql =  sql   &   "	(  "
sql =  sql   &   "		select distinct yms,deptgroup,teacher,category,group1 from boo_schedule where yms='992' and category='T'  "
sql =  sql   &   "	) b on a.teachername=b.teacher and a.yms=b.yms   where b.deptgroup is not null and a.yn='Y' and signin is not null and a.yms='992' and a.category='T' "
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

sql = sql&   ") a group by a.group1,a.deptgroup,a.teachername,a.item order by  a.deptgroup "


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
<P align="center" class="inputlabel"><font size="4">駐診教師總統計</font></P>

<%

	ReportStr = ReportHeader()
	while not rs.EOF
		if  rs("deptgroup") = "LDCC英外語能力診斷輔導中心" then
			ReportStr=replace(ReportStr,":L"&rs("teachername")&"時"&rs("item"),rs("alltime"))

		elseif  rs("deptgroup") = "ELC英語學習中心" then
			ReportStr=replace(ReportStr,":E"&rs("teachername")&"時"&rs("item"),rs("alltime"))
		end if
		rs.movenext
	wend
	rs.Close
	'人次數
	sql =		" select  a.group1,a.deptgroup,a.item ,a.teachername,count(*) as  cnt   from "
	sql = sql &		" ("
	sql = sql &				" select b.group1,b.deptgroup,a.bdate,teachername,sid,item  "
	sql = sql &				" from boo_book_T_M a left join  "
	sql = sql &				" ( "
	sql = sql &				" select distinct yms,deptgroup,teacher,category,group1  from boo_schedule where yms='"&yms&"' and category='T'  "
	sql = sql &				" )  b on a.teachername=b.teacher and a.yms=b.yms where b.deptgroup is not null and  a.yn='Y' and signin is not null and a.yms='"&yms&"' and a.category='T' "
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
	sql = sql &		 " ) a group by a.group1,a.teachername,a.deptgroup,a.item order by  a.deptgroup "

'response.write sql 
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
	while not rs.EOF
		if  rs("deptgroup") = "LDCC英外語能力診斷輔導中心" then
			ReportStr=replace(ReportStr,":L"&rs("teachername")&"次"&rs("item"),rs("cnt"))

		elseif  rs("deptgroup") = "ELC英語學習中心" then
			ReportStr=replace(ReportStr,":E"&rs("teachername")&"次"&rs("item"),rs("cnt"))
		end if
		rs.movenext
	wend
	rs.Close
	'人數
	sql = 		" select  a.group1,a.deptgroup,a.teachername,a.item,count(*)  cnt1   from "
	sql = sql &		" ("
	sql = sql &				" select distinct b.group1,b.deptgroup,teachername,sid,a.item "
	sql = sql &				" from boo_book_T_M a left join  "
	sql = sql &				" ( "
	sql = sql &				" select distinct yms,deptgroup,teacher,category,group1  from boo_schedule where yms='"&yms&"' and category='T'  "
	sql = sql &				" )  b on a.teachername=b.teacher and a.yms=b.yms where b.deptgroup is not null and  a.yn='Y' and signin is not null and a.yms='"&yms&"' and a.category='T' "
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
	sql = sql &		 " ) a group by a.group1,a.teachername,a.deptgroup,a.item  "
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
	while not rs.EOF
		if  rs("deptgroup") = "LDCC英外語能力診斷輔導中心" then
			ReportStr=replace(ReportStr,":L"&rs("teachername")&"人"&rs("item"),rs("cnt1"))

		elseif  rs("deptgroup") = "ELC英語學習中心" then
			ReportStr=replace(ReportStr,":E"&rs("teachername")&"人"&rs("item"),rs("cnt1"))
		end if
		rs.movenext
	wend
	rs.Close
'response.write sql
	'清空
	sql = "select distinct a.teachername,b.group1  from boo_book_T_M a left join  "
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
			sql = sql & " and ( Cast(a.bdate as int)>= '" & sdate & "'   and  Cast(a.bdate as int) <= '" & edate& "' )"
		end if
		if slevel<>"" then
			sql = sql & " and a.slevel='"&slevel&"'"
		end if
		if department<>"" then
			sql = sql & " and a.department='"&department&"'"
		end if
	sql = sql &  " order by b.group1 " 
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
	while not rs.EOF
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"人診斷","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"人諮商","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"人口語","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"人簡報","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"人詩歌","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"人寫作","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"人閱讀","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"次診斷","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"次諮商","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"次口語","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"次簡報","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"次詩歌","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"次寫作","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"次閱讀","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"時診斷","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"時諮商","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"時口語","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"時簡報","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"時詩歌","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"時寫作","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"時閱讀","&nbsp;")
		
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"人診斷","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"人諮商","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"人口語","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"人簡報","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"人詩歌","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"人寫作","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"人閱讀","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"次診斷","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"次諮商","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"次口語","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"次簡報","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"次詩歌","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"次寫作","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"次閱讀","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"時診斷","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"時諮商","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"時口語","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"時簡報","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"時詩歌","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"時寫作","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"時閱讀","&nbsp;")

		rs.movenext
	wend
	rs.Close
'合計
	sql =" select a.group1,a.teachername,c.cnt1,b.cnt,a.alltime from"
	sql = sql &" ("
	'時數
	sql = sql &		" select  a.group1,a.teachername,sum(Cast(a.timeflag as decimal(2,1) )) as alltime    from "
	sql = sql &		" ("
	sql = sql &				" select b.group1,a.bdate,replace(replace(replace(a.timeflag,'A',1),'U',0.5),'B',0.5) as timeflag  ,teachername "
	sql = sql &				" from boo_book_T_M a left join  "
	sql = sql &				" ( "
	sql = sql &					" select distinct yms,teacher,category,group1,deptgroup  from boo_schedule where yms='"&yms&"' and category='T'  "
	sql = sql &				" )  b on a.teachername=b.teacher and a.yms=b.yms where b.deptgroup is not null and  a.yn='Y' and signin is not null and a.yms='"&yms&"' and a.category='T' "
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
	sql = sql &				" select distinct yms,teacher,category,group1,deptgroup  from boo_schedule where yms='"&yms&"' and category='T'  "
	sql = sql &				" )  b on a.teachername=b.teacher and a.yms=b.yms where  b.deptgroup is not null and a.yn='Y' and signin is not null and a.yms='"&yms&"' and a.category='T' "
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
	sql = sql &				" select distinct yms,teacher,category,group1,deptgroup  from boo_schedule where yms='"&yms&"' and category='T'  "
	sql = sql &				" )  b on a.teachername=b.teacher and a.yms=b.yms where  b.deptgroup is not null and  a.yn='Y' and signin is not null and a.yms='"&yms&"' and a.category='T' "
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

'response.write  "<br>" & sql
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
	while not rs.EOF
		ReportStr=replace(ReportStr,":"&rs("teachername")&"時合計",rs("alltime"))
		ReportStr=replace(ReportStr,":"&rs("teachername")&"次合計",rs("cnt"))
		ReportStr=replace(ReportStr,":"&rs("teachername")&"人合計",rs("cnt1"))
		rs.movenext
	wend
	rs.Close


	response.write ReportStr

%>

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
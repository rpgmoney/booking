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
	Str="<TABLE cellSpacing=0 cellPadding=3 border=1   ><TR class=""inputlabel""><TD rowspan=2>&nbsp;</TD><TD rowspan=2>�t��</TD><TD rowspan=2>�Юv</TD><TD rowspan=2>&nbsp;</TD><TD colspan=7 align=center>LDCC�^�~�y��O�E�_���ɤ���</TD><TD colspan=7 align=center>ELC�^�y�ǲߤ���</TD><TD rowspan=2>�X�p</TD>"
	
	Str = Str & "</TR>"
	Str = Str & "</TR><TD>�E�_</TD><TD>�԰�</TD><TD>�f�y</TD><TD>²��</TD><TD>�ֺq</TD><TD>�g�@</TD><TD>�\Ū</TD><TD>�E�_</TD><TD>�԰�</TD><TD>�f�y</TD><TD>²��</TD><TD>�ֺq</TD><TD>�g�@</TD><TD>�\Ū</TD>"
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
		Str = Str & "<tr nowrap><td rowspan=3>"&i&"</td><td rowspan=3 nowrap>"&rs0("group1")&"</td><td rowspan=3 nowrap>"&rs0("teachername")&"</td><td nowrap>�����ǥͤH��</td>"
		Str = Str & "<TD >:L"&rs0("teachername")&"�H�E�_</TD><TD >:L"&rs0("teachername")&"�H�԰�</TD><TD>:L"&rs0("teachername")&"�H�f�y</TD><TD>:L"&rs0("teachername")&"�H²��</TD><TD>:L"&rs0("teachername")&"�H�ֺq</TD><TD>:L"&rs0("teachername")&"�H�g�@</TD><TD>:L"&rs0("teachername")&"�H�\Ū</TD><TD>:E"&rs0("teachername")&"�H�E�_</TD><TD>:E"&rs0("teachername")&"�H�԰�</TD><TD>:E"&rs0("teachername")&"�H�f�y</TD><TD>:E"&rs0("teachername")&"�H²��</TD><TD>:E"&rs0("teachername")&"�H�ֺq</TD><TD>:E"&rs0("teachername")&"�H�g�@</TD><TD>:E"&rs0("teachername")&"�H�\Ū</TD><TD>:"&rs0("teachername")&"�H�X�p</TD>"
		Str = Str & "</tr>"
		Str = Str & "<tr nowrap><td>���ɦ���</td>"
		Str = Str & "<TD>:L"&rs0("teachername")&"���E�_</TD><TD>:L"&rs0("teachername")&"���԰�</TD><TD>:L"&rs0("teachername")&"���f�y</TD><TD>:L"&rs0("teachername")&"��²��</TD><TD>:L"&rs0("teachername")&"���ֺq</TD><TD>:L"&rs0("teachername")&"���g�@</TD><TD>:L"&rs0("teachername")&"���\Ū</TD><TD>:E"&rs0("teachername")&"���E�_</TD><TD>:E"&rs0("teachername")&"���԰�</TD><TD>:E"&rs0("teachername")&"���f�y</TD><TD>:E"&rs0("teachername")&"��²��</TD><TD>:E"&rs0("teachername")&"���ֺq</TD><TD>:E"&rs0("teachername")&"���g�@</TD><TD>:E"&rs0("teachername")&"���\Ū</TD><TD>:"&rs0("teachername")&"���X�p</TD>"
		Str = Str & "</tr>"
		Str = Str & "<tr ><td>���ɮɼ�</td>"
		Str = Str & "<TD nowrap>:L"&rs0("teachername")&"�ɶE�_</TD><TD nowrap>:L"&rs0("teachername")&"�ɿ԰�</TD><TD nowrap>:L"&rs0("teachername")&"�ɤf�y</TD><TD nowrap>:L"&rs0("teachername")&"��²��</TD><TD nowrap>:L"&rs0("teachername")&"�ɸֺq</TD><TD nowrap>:L"&rs0("teachername")&"�ɼg�@</TD><TD nowrap>:L"&rs0("teachername")&"�ɾ\Ū</TD><TD nowrap>:E"&rs0("teachername")&"�ɶE�_</TD><TD nowrap>:E"&rs0("teachername")&"�ɿ԰�</TD><TD nowrap>:E"&rs0("teachername")&"�ɤf�y</TD><TD nowrap>:E"&rs0("teachername")&"��²��</TD><TD nowrap>:E"&rs0("teachername")&"�ɸֺq</TD><TD nowrap>:E"&rs0("teachername")&"�ɼg�@</TD><TD nowrap>:E"&rs0("teachername")&"�ɾ\Ū</TD><TD nowrap>:"&rs0("teachername")&"�ɦX�p</TD>"
		Str = Str & "</tr>"

		rs0.MoveNext
	wend



	Str = Str & "</TABLE>"
	ReportHeader=Str
	
end function



set rs=server.CreateObject("adodb.recordset")
set rs2=server.CreateObject("adodb.recordset")



'�H��
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
<TITLE> �iLDCC�^�~�y��O�E�_���ɤ��ߡj  </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>

</HEAD>
<BODY bottomMargin=0 leftMargin=2 topMargin=0 rightMargin=2 marginheight="0" marginwidth="0"  >

<%
	if not rs.eof then 
%>
<P align="center" class="inputlabel"><font size="4">�n�E�Юv�`�έp</font></P>

<%

	ReportStr = ReportHeader()
	while not rs.EOF
		if  rs("deptgroup") = "LDCC�^�~�y��O�E�_���ɤ���" then
			ReportStr=replace(ReportStr,":L"&rs("teachername")&"��"&rs("item"),rs("alltime"))

		elseif  rs("deptgroup") = "ELC�^�y�ǲߤ���" then
			ReportStr=replace(ReportStr,":E"&rs("teachername")&"��"&rs("item"),rs("alltime"))
		end if
		rs.movenext
	wend
	rs.Close
	'�H����
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
		if  rs("deptgroup") = "LDCC�^�~�y��O�E�_���ɤ���" then
			ReportStr=replace(ReportStr,":L"&rs("teachername")&"��"&rs("item"),rs("cnt"))

		elseif  rs("deptgroup") = "ELC�^�y�ǲߤ���" then
			ReportStr=replace(ReportStr,":E"&rs("teachername")&"��"&rs("item"),rs("cnt"))
		end if
		rs.movenext
	wend
	rs.Close
	'�H��
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
		if  rs("deptgroup") = "LDCC�^�~�y��O�E�_���ɤ���" then
			ReportStr=replace(ReportStr,":L"&rs("teachername")&"�H"&rs("item"),rs("cnt1"))

		elseif  rs("deptgroup") = "ELC�^�y�ǲߤ���" then
			ReportStr=replace(ReportStr,":E"&rs("teachername")&"�H"&rs("item"),rs("cnt1"))
		end if
		rs.movenext
	wend
	rs.Close
'response.write sql
	'�M��
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
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"�H�E�_","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"�H�԰�","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"�H�f�y","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"�H²��","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"�H�ֺq","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"�H�g�@","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"�H�\Ū","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"���E�_","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"���԰�","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"���f�y","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"��²��","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"���ֺq","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"���g�@","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"���\Ū","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"�ɶE�_","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"�ɿ԰�","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"�ɤf�y","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"��²��","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"�ɸֺq","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"�ɼg�@","&nbsp;")
		ReportStr=replace(ReportStr,":L"&rs("teachername")&"�ɾ\Ū","&nbsp;")
		
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"�H�E�_","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"�H�԰�","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"�H�f�y","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"�H²��","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"�H�ֺq","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"�H�g�@","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"�H�\Ū","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"���E�_","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"���԰�","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"���f�y","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"��²��","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"���ֺq","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"���g�@","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"���\Ū","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"�ɶE�_","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"�ɿ԰�","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"�ɤf�y","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"��²��","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"�ɸֺq","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"�ɼg�@","&nbsp;")
		ReportStr=replace(ReportStr,":E"&rs("teachername")&"�ɾ\Ū","&nbsp;")

		rs.movenext
	wend
	rs.Close
'�X�p
	sql =" select a.group1,a.teachername,c.cnt1,b.cnt,a.alltime from"
	sql = sql &" ("
	'�ɼ�
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
	'�H����
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
	'�H��
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
		ReportStr=replace(ReportStr,":"&rs("teachername")&"�ɦX�p",rs("alltime"))
		ReportStr=replace(ReportStr,":"&rs("teachername")&"���X�p",rs("cnt"))
		ReportStr=replace(ReportStr,":"&rs("teachername")&"�H�X�p",rs("cnt1"))
		rs.movenext
	wend
	rs.Close


	response.write ReportStr

%>

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
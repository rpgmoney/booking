<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<!-- #INCLUDE virtual="/include/lib/asp/jsCalendar.asp" -->
<!-- #INCLUDE file="lib/parameter.inc" -->
<%
AColor="#FF0000"'�w���C��
PColor="blue"'�f�ֳq�L
validate=trim(request("validate"))
sid=trim(request("sid"))
BOOK_DATE=trim(request("BOOK_DATE"))

today1=year(dateadd("d",1,date()))-1911 & right("0" & month(dateadd("d",1,date())),2) & right("0" & day(dateadd("d",1,date())),2)

if BOOK_DATE="" or isnull(BOOK_DATE) or isempty(BOOK_DATE) then
	BOOK_DATE=today1

end if

set rs = server.CreateObject("adodb.recordset")
function dateformat(vdate)
	if vdate<>"" then
		if  Cstr(left(vdate,1)) ="9"  then
		dateformat=cStr(cint(left(vdate,2))+1911 ) & "/" & mid(vdate,3,2) & "/" & right(vdate,2)
		else
		dateformat=cStr(cint(left(vdate,3))+1911 ) & "/" & mid(vdate,4,2) & "/" & right(vdate,2)
		end if
	end if
	
end function
%>
<HTML>
<HEAD>
<TITLE> �iLDCC�^�~�y��O�E�_���ɤ��ߡj </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/stm31.js" type="text/javascript"></SCRIPT>
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>
<script language="javascript">

function changedate(v)
{
	
	var  frmlistform = document.getElementById("form1");
	frmlistform.BOOK_DATE.value=v;
	frmlistform.submit();
}
function getValue(vbdate,vbtime,vteachername,vlanguagecode)
{
	form2.teachername.value=vteachername;
	form2.bdate.value=vbdate;
	form2.btime.value=vbtime;
	form2.languagecode.value=vlanguagecode;
	form2.submit();
}
</script>
</HEAD>
<BODY bottomMargin=0 leftMargin=0 topMargin=0 rightMargin=0 marginheight="0" marginwidth="0"  bgcolor="#f4c60d">
<form name="form2" method=post action="bookingt.asp">
<input type="hidden" value="" name="teachername">
<input type="hidden" value="" name="bdate">
<input type="hidden" value="" name="btime">
<input type="hidden" value="ST" name="category">
<input type="hidden" value="1" name="showflag">
<input type="hidden" value="" name="languagecode">
</form>
<TABLE cellSpacing=1 cellPadding=0 width="760"  height="100%" align="center" >
<TR><TD>
<TABLE cellSpacing=0 cellPadding=0 width="760"  height="100%" align="center" bgColor=#ffffff border=0>
<TR height="70"><TD><img src="images\top.jpg" border="0"></TD></TR>
<TR height="25" bgcolor="#333333">
	<TD align="center">
	<TABLE cellSpacing=0 cellPadding=0 border=0 width="100%">
	<TR>
	<TD align="left"><!-- #INCLUDE FILE="lib\link.inc" --></TD>
	<TD align="right"><!-- #INCLUDE FILE="lib\promsg.inc" --></TD>
	</TR>
	</TABLE>
	</TD>
</TR>
<TR valign="top">
	<TD>
<!-- ---------------------------------------------------------------------------------------- -->
	<P><BR>
	<TABLE cellSpacing=3 cellPadding=3 border=0 width="100%">
	<TR >
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">�d�ߤp�Ѯv�������{�Z�� </TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>
	
	<TR>
		<TD></TD><TD valign="top">
		<form name="form1" method=post action="booksteacher.asp">
		<P><BR>
		<table border=0 cellpadding=0 cellspacing=2 width="98%"  align="center">
		<tr valign=top> 
		<td>
			<table border=0 cellpadding=0 cellspacing=2  align="left">
			
			<td>�d�ߤ���G</td>                     
			<td> 
			<input type="text" value="<%=BOOK_DATE%>" maxlength="6" name="BOOK_DATE" id="BOOK_DATE" class="inputtext">
			<img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('BOOK_DATE')" class="showhand">
			</td>
			<td><input  type="submit" value="�d��" ></td>
			</tr>
			</table>
			<P><BR><BR>
		</td>                         
		</tr>
		<tr valign=top> 
		<td>
			<table border=0 cellpadding=0 cellspacing=2  align="left">
			<td align="left">
			<%
			sql = "select * from boo_language where yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				response.write "&nbsp;&nbsp;<font color='"&rs("showcolor")&"'>�� </font>&nbsp;" & rs("name")
				rs.MoveNext
			wend
			rs.close
			sql = "select * from boo_noopen where yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			Strnoopendate=""
			while not rs.EOF
				if  left(rs("noopendate"),1) = "9" then
					Strnoopendate = Strnoopendate &left(rs("noopendate"),2)+1911 & "/" & mid(rs("noopendate"),3,2) & "/" & right(rs("noopendate"),2)  & ","
				else
					Strnoopendate = Strnoopendate &left(rs("noopendate"),3)+1911 & "/" & mid(rs("noopendate"),4,2) & "/" & right(rs("noopendate"),2)  & ","
				end if
				rs.MoveNext
			wend
			rs.close
			%>
			</td>
			</table>
		</td>                         
		</tr>
		<tr>
		<td>
			<table border=0 cellpadding=0 cellspacing=2 width="100%" >
			<tr>
				<td align="left">
				<font color="#CC9900">��</font>����N��w�w���A<font color="blue">��</font>�Ŧ�N��i�w���A<font color="#999999">��</font>�Ǧ�N���b�w������
				</td>
				<td>
					<table border=0 cellpadding=0 cellspacing=2 align="right" >
					<tr>
					<td align="right"  class="showhand" onclick="changedate('<%=datetoNumformat(dateadd("d",-7,NumberToDateFormat(BOOK_DATE)))%>')"><a href="#"><font color="red">&nbsp;<<�e�@�g&nbsp;</font></a></td>
					<td>&nbsp;�U&nbsp;</td>
					<td align="right" class="showhand" onclick="changedate('<%=datetoNumformat(dateadd("d",7,NumberToDateFormat(BOOK_DATE)))%>')"><a href="#"><font color="red">&nbsp;��@�g&nbsp;>></font></a></td>
					</tr>
					</table>
				</td>
			</tr>
			</table>
			
		
		</td> 
		</form>
		</tr>
		<%
			set dic = server.CreateObject("scripting.dictionary")
			dic.Add "1","��"
			dic.Add "2","�@"
			dic.Add "3","�G"
			dic.Add "4","�T"
			dic.Add "5","�|"
			dic.Add "6","��"
			dic.Add "7","��"
			set rs = server.CreateObject("adodb.recordset")
			'tmpdate=left(BOOK_DATE,2)+1911 & "/" & mid(BOOK_DATE,3,2) & "/" & right(BOOK_DATE,2)
			tmpdate=dateformat(BOOK_DATE)
			StrDateTop=""'���D������M�P��
			StrStatus0810=""'8�I
			StrStatus0910=""'9�I
			StrStatus1010=""'10�I
			StrStatus1110=""'11�I
			StrStatus1310=""'13�I
			StrStatus1410=""'14�I
			StrStatus1510=""'15�I
			StrStatus1610=""'16�I

			Str0810=""'8�I
			Str0910=""'9�I
			Str1010=""'10�I
			Str1110=""'11�I
			Str1310=""'13�I
			Str1410=""'14�I
			Str1510=""'15�I
			Str1610=""'16�I
			
			'response.write "BOOK_ROOM=" & BOOK_ROOM
			
			for i=0 to 6
				ww = weekday(tmpdate)
				if ww="1" then
					tmpColor="#FFEEEE"
				elseif ww="7" then
					tmpColor="#ECF9F2"
				else
					tmpColor="#FFFFF4"
				end if
					
				tmpdate = year(tmpdate) & "/" & right("0" & month(tmpdate),2) & "/" & right("0" & day(tmpdate),2)	

				
				if ww="1" or ww="7" then
					'Str1010 = Str1010 & "<td bgcolor="&tmpColor&">&nbsp;</td>"
					'Str1110 = Str1110 & "<td bgcolor="&tmpColor&">&nbsp;</td>"

					'Str1310 = Str1310 & "<td bgcolor="&tmpColor&">&nbsp;</td>"
					'Str1410 = Str1410 & "<td bgcolor="&tmpColor&">&nbsp;</td>"
					'Str1510 = Str1510 & "<td bgcolor="&tmpColor&">&nbsp;</td>"
					'Str1610 = Str1610 & "<td bgcolor="&tmpColor&">&nbsp;</td>"
				

				else
					'���D������M�P��
					
					StrDateTop = StrDateTop & "<td bgcolor=""#c1e0a3"" align=""center"">" & tmpdate & "<br>�P��" & dic.Item(cstr(ww)) & "</td>"
					if  InStr(Strnoopendate,Cstr(tmpdate))  then 				
						
						Str0810 = Str0810 & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						Str0910 = Str0910 & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						Str1010 = Str1010 & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						Str1110 = Str1110 & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						Str1310 = Str1310 & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						Str1410 = Str1410 & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						Str1510 = Str1510 & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						Str1610 = Str1610 & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						Str1710 = Str1710 & "<td bgcolor=#FFCCFF>&nbsp;</td>"
					else
						'8�I
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate,c.showcolor  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername  and b.category='ST'  and b.btime='0810' and b.timeflag in ('A') and b.bdate='"&datetoNumformat(tmpdate)&"' and b.yn='Y' "
						sql = sql & "left join boo_language c on a.languagecode = c.code"
						sql = sql & " where b.pid is null and a.category='ST'and a.btime='0810' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y' and  a.yms='"&par_yms&"' "

						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus0810=""
						
						while not rs.EOF
							if rs("name")="N" then
								if (cdbl(datetoNumformat(tmpdate)) <= cdbl(datetoNumformat(dateadd("d",14,date())))  and  cdbl(datetoNumformat(tmpdate)) > cdbl(datetoNumformat(date())) )  or  session("classify")="A" then
									StrStatus0810 = StrStatus0810 & "<font color='"&rs("showcolor")&"'>��</font>"  & "<a title='�i�w��' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=0810&category=ST&teachername="&rs("teacher")&"&showflag=1&languagecode="&rs("languagecode")&"'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								else
									StrStatus0810 = StrStatus0810 & "<font color='"&rs("showcolor")&"'>��</font>"   & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'�w����T
								StrStatus0810 = StrStatus0810 & "<font color='"&rs("showcolor")&"'>��</font>"  & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');""  title='�w���H�G"&rs("name")&"("&rs("department")&"�A"&rs("grade")&"�~�šA"&rs("class1")&"�Z)'><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str0810 = Str0810 & "<td bgcolor="&tmpColor&">" & StrStatus0810 & "&nbsp;</td>"


						'9�I
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate,c.showcolor  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername  and b.category='ST'  and b.btime='0910' and b.timeflag in ('A') and b.bdate='"&datetoNumformat(tmpdate)&"' and b.yn='Y' "
						sql = sql & "left join boo_language c on a.languagecode = c.code"
						sql = sql & " where b.pid is null and a.category='ST'and a.btime='0910' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y' and  a.yms='"&par_yms&"'"
						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus0910=""
						
						while not rs.EOF
							if rs("name")="N" then
								if  (  cdbl(datetoNumformat(tmpdate)) <= cdbl(datetoNumformat(dateadd("d",14,date())))  and  cdbl(datetoNumformat(tmpdate)) > cdbl(datetoNumformat(date())) )  or  session("classify")="A" then
									StrStatus0910 = StrStatus0910 & "<font color='"&rs("showcolor")&"'>��</font>"  & "<a title='�i�w��' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=0910&category=ST&teachername="&rs("teacher")&"&showflag=1&languagecode="&rs("languagecode")&"'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								else
									response.write "tmpdate=" & tmpdate
									StrStatus0910 = StrStatus0910 & "<font color='"&rs("showcolor")&"'>��</font>"   & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'�w����T
								StrStatus0910 = StrStatus0910 & "<font color='"&rs("showcolor")&"'>��</font>"  & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');""  title='�w���H�G"&rs("name")&"("&rs("department")&"�A"&rs("grade")&"�~�šA"&rs("class1")&"�Z)'><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str0910 = Str0910 & "<td bgcolor="&tmpColor&">" & StrStatus0910 & "&nbsp;</td>"


						'10�I
						'sql = "select * from boo_schedule where category='T' and btime='1010' and bweek='"&cint(ww)-1&"'  and yn='Y'"
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate,c.showcolor  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername  and b.category='ST'  and b.btime='1010' and b.timeflag in ('A') and b.bdate='"&datetoNumformat(tmpdate)&"' and b.yn='Y' "
						sql = sql & "left join boo_language c on a.languagecode = c.code"
						sql = sql & " where b.pid is null and a.category='ST'and a.btime='1010' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y' and  a.yms='"&par_yms&"'"
						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1010=""
						
						while not rs.EOF
							if rs("name")="N" then
								if ( cdbl(datetoNumformat(tmpdate)) <= cdbl(datetoNumformat(dateadd("d",14,date())))   and  cdbl(datetoNumformat(tmpdate)) > cdbl(datetoNumformat(date())))  or  session("classify")="A" then
									'StrStatus1010 = StrStatus1010 & "<font color='"&rs("showcolor")&"'>��</font>"  & "<a title='�i�w��' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1010&category=ST&teachername="&rs("teacher")&"&showflag=1&languagecode="&rs("languagecode")&"'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
									StrStatus1010 = StrStatus1010 & "<font color='"&rs("showcolor")&"'>��</font>"  & "<a title='�i�w��' href=""#"" onclick='getValue("""&datetoNumformat(tmpdate)&""",""1010"","""&rs("teacher")&""","""&rs("languagecode")&""")'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								else
									StrStatus1010 = StrStatus1010 & "<font color='"&rs("showcolor")&"'>��</font>"   & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'�w����T
								StrStatus1010 = StrStatus1010 & "<font color='"&rs("showcolor")&"'>��</font>"  & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');""  title='�w���H�G"&rs("name")&"("&rs("department")&"�A"&rs("grade")&"�~�šA"&rs("class1")&"�Z)'><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1010 = Str1010 & "<td bgcolor="&tmpColor&">" & StrStatus1010 & "&nbsp;</td>"

						'11�I
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate,c.showcolor  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername and b.btime='1110' and b.timeflag in ('A') and b.bdate='"&datetoNumformat(tmpdate)&"' and b.yn='Y'"
						sql = sql & "left join boo_language c on a.languagecode = c.code"
						sql = sql & " where b.pid is null and  a.category='ST' and a.btime='1110' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y' and  a.yms='"&par_yms&"'"

						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1110=""
						while not rs.EOF
							if rs("name")="N" then
								if  (cdbl(datetoNumformat(tmpdate)) <= cdbl(datetoNumformat(dateadd("d",14,date())))  and  cdbl(datetoNumformat(tmpdate)) >  cdbl(datetoNumformat(date())))  or  session("classify")="A" then
									'StrStatus1110 = StrStatus1110 & "<font color='"&rs("showcolor")&"'>��</font>" & "<a title='�i�w��' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1110&category=ST&teachername="&rs("teacher")&"&showflag=1&languagecode="&rs("languagecode")&"'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
									StrStatus1110 = StrStatus1110 & "<font color='"&rs("showcolor")&"'>��</font>"  & "<a title='�i�w��' href=""#"" onclick='getValue("""&datetoNumformat(tmpdate)&""",""1110"","""&rs("teacher")&""","""&rs("languagecode")&""")'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								else
									StrStatus1110 = StrStatus1110 & "<font color='"&rs("showcolor")&"'>��</font>" & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'�w����T
								StrStatus1110 = StrStatus1110 & "<font color='"&rs("showcolor")&"'>��</font>"  & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');"" title='�w���H�G"&rs("name")&"("&rs("department")&"�A"&rs("grade")&"�~�šA"&rs("class1")&"�Z)'><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1110 = Str1110 & "<td bgcolor="&tmpColor&">" & StrStatus1110 & "&nbsp;</td>"

						

						'13�I
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate,c.showcolor  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername and b.btime='1310' and b.timeflag in ('A') and b.bdate='"&datetoNumformat(tmpdate)&"' and b.yn='Y' "
						sql = sql & "left join boo_language c on a.languagecode = c.code"
						sql = sql & " where b.pid is null and  a.category='ST' and a.btime='1310' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y' and  a.yms='"&par_yms&"'"
						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1310=""
						while not rs.EOF
							if rs("name")="N" then
								if  (cdbl(datetoNumformat(tmpdate)) <= cdbl(datetoNumformat(dateadd("d",14,date())))  and  cdbl(datetoNumformat(tmpdate)) > cdbl(datetoNumformat(date())))  or  session("classify")="A" then
									'StrStatus1310 = StrStatus1310  & "<font color='"&rs("showcolor")&"'>��</font>" & "<a title='�i�w��' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1310&category=ST&teachername="&rs("teacher")&"&showflag=1&languagecode="&rs("languagecode")&"'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
									StrStatus1310 = StrStatus1310 & "<font color='"&rs("showcolor")&"'>��</font>"  & "<a title='�i�w��' href=""#"" onclick='getValue("""&datetoNumformat(tmpdate)&""",""1310"","""&rs("teacher")&""","""&rs("languagecode")&""")'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								else
									StrStatus1310 = StrStatus1310  & "<font color='"&rs("showcolor")&"'>��</font>" & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'�w����T
								StrStatus1310 = StrStatus1310 & "<font color='"&rs("showcolor")&"'>��</font>"  & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');"" title='�w���H�G"&rs("name")&"("&rs("department")&"�A"&rs("grade")&"�~�šA"&rs("class1")&"�Z)'><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1310 = Str1310 & "<td bgcolor="&tmpColor&">" & StrStatus1310 & "&nbsp;</td>"
						'14�I
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate,c.showcolor  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername and b.btime='1410' and b.timeflag in ('A') and b.bdate='"&datetoNumformat(tmpdate)&"' and b.yn='Y' "
						sql = sql & "left join boo_language c on a.languagecode = c.code"
						sql = sql & " where b.pid is null and  a.category='ST' and a.btime='1410' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y' and  a.yms='"&par_yms&"'"
						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1410=""
						while not rs.EOF
							if rs("name")="N" then
								if  (cdbl(datetoNumformat(tmpdate)) <= cdbl(datetoNumformat(dateadd("d",14,date())))  and  cdbl(datetoNumformat(tmpdate)) > cdbl(datetoNumformat(date())))  or  session("classify")="A" then
									'StrStatus1410 = StrStatus1410  & "<font color='"&rs("showcolor")&"'>��</font>" & "<a title='�i�w��' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1410&category=ST&teachername="&rs("teacher")&"&showflag=1&languagecode="&rs("languagecode")&"'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
									StrStatus1410 = StrStatus1410 & "<font color='"&rs("showcolor")&"'>��</font>"  & "<a title='�i�w��' href=""#"" onclick='getValue("""&datetoNumformat(tmpdate)&""",""1410"","""&rs("teacher")&""","""&rs("languagecode")&""")'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								else
									StrStatus1410 = StrStatus1410  & "<font color='"&rs("showcolor")&"'>��</font>" & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'�w����T
								StrStatus1410 = StrStatus1410 & "<font color='"&rs("showcolor")&"'>��</font>"  & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');"" title='�w���H�G"&rs("name")&"("&rs("department")&"�A"&rs("grade")&"�~�šA"&rs("class1")&"�Z)'><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1410 = Str1410 & "<td bgcolor="&tmpColor&">" & StrStatus1410 & "&nbsp;</td>"

						'15�I
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate,c.showcolor  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername and b.btime='1510' and b.timeflag in ('A') and b.bdate='"&datetoNumformat(tmpdate)&"'  and b.yn='Y' "
						sql = sql & "left join boo_language c on a.languagecode = c.code"
						sql = sql & " where b.pid is null and  a.category='ST' and a.btime='1510' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y' and  a.yms='"&par_yms&"'"
						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1510=""
						while not rs.EOF
							if rs("name")="N" then
								if  (cdbl(datetoNumformat(tmpdate)) <=  cdbl(datetoNumformat(dateadd("d",14,date())))  and  cdbl(datetoNumformat(tmpdate)) > cdbl(datetoNumformat(date())))  or  session("classify")="A" then
									'StrStatus1510 = StrStatus1510 & "<font color='"&rs("showcolor")&"'>��</font>"  & "<a title='�i�w��' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1510&category=ST&teachername="&rs("teacher")&"&showflag=1&languagecode="&rs("languagecode")&"' ><font color='blue'>" & rs("teacher") & "</font></a><br>" 
									StrStatus1510 = StrStatus1510 & "<font color='"&rs("showcolor")&"'>��</font>"  & "<a title='�i�w��' href=""#"" onclick='getValue("""&datetoNumformat(tmpdate)&""",""1510"","""&rs("teacher")&""","""&rs("languagecode")&""")'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								else
									StrStatus1510 = StrStatus1510 & "<font color='"&rs("showcolor")&"'>��</font>" & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'�w����T
								StrStatus1510 = StrStatus1510 & "<font color='"&rs("showcolor")&"'>��</font>"  & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');"" title='�w���H�G"&rs("name")&"("&rs("department")&"�A"&rs("grade")&"�~�šA"&rs("class1")&"�Z)'><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1510 = Str1510 & "<td bgcolor="&tmpColor&">" & StrStatus1510 & "&nbsp;</td>"

						'16�I
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate,c.showcolor  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername and b.btime='1610' and b.timeflag in ('A') and b.bdate='"&datetoNumformat(tmpdate)&"'  and b.yn='Y' "
						sql = sql & "left join boo_language c on a.languagecode = c.code"
						sql = sql & " where b.pid is null and  a.category='ST' and a.btime='1610' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y' and  a.yms='"&par_yms&"'"
						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1610=""
						while not rs.EOF
							if rs("name")="N" then
								if  (cdbl(datetoNumformat(tmpdate)) <= cdbl(datetoNumformat(dateadd("d",14,date())))  and  cdbl(datetoNumformat(tmpdate)) > cdbl(datetoNumformat(date())))  or  session("classify")="A" then
									'StrStatus1610 = StrStatus1610  & "<font color='"&rs("showcolor")&"'>��</font>" & "<a title='�i�w��' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1610&category=ST&teachername="&rs("teacher")&"&showflag=1&languagecode="&rs("languagecode")&"'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
									StrStatus1610 = StrStatus1610 & "<font color='"&rs("showcolor")&"'>��</font>"  & "<a title='�i�w��' href=""#"" onclick='getValue("""&datetoNumformat(tmpdate)&""",""1610"","""&rs("teacher")&""","""&rs("languagecode")&""")'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								else
									StrStatus1610 = StrStatus1610  & "<font color='"&rs("showcolor")&"'>��</font>" & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'�w����T
								StrStatus1610 = StrStatus1610 & "<font color='"&rs("showcolor")&"'>��</font>"  & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');"" title='�w���H�G"&rs("name")&"("&rs("department")&"�A"&rs("grade")&"�~�šA"&rs("class1")&"�Z)'><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1610 = Str1610 & "<td bgcolor="&tmpColor&">" & StrStatus1610 & "&nbsp;</td>"

						'17�I
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate,c.showcolor  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername and b.btime='1710' and b.timeflag in ('A') and b.bdate='"&datetoNumformat(tmpdate)&"'  and b.yn='Y'  "
						sql = sql & "left join boo_language c on a.languagecode = c.code"
						sql = sql & " where b.pid is null and  a.category='ST' and a.btime='1710' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y' and  a.yms='"&par_yms&"'"
						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1710=""
						while not rs.EOF
							if rs("name")="N" then
								if  (cdbl(datetoNumformat(tmpdate))  <=  cdbl(datetoNumformat(dateadd("d",14,date())))  and  cdbl(datetoNumformat(tmpdate)) > cdbl(datetoNumformat(date())))  or  session("classify")="A" then
									StrStatus1710 = StrStatus1710  & "<font color='"&rs("showcolor")&"'>��</font>" & "<a title='�i�w��' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1710&category=ST&teachername="&rs("teacher")&"&showflag=1&languagecode="&rs("languagecode")&"'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								else
									StrStatus1710 = StrStatus1710  & "<font color='"&rs("showcolor")&"'>��</font>" & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'�w����T
								StrStatus1710 = StrStatus1710 & "<font color='"&rs("showcolor")&"'>��</font>"  & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');"" title='�w���H�G"&rs("name")&"("&rs("department")&"�A"&rs("grade")&"�~�šA"&rs("class1")&"�Z)'><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1710 = Str1710 & "<td bgcolor="&tmpColor&">" & StrStatus1710 & "&nbsp;</td>"


					end if	
				end if
				tmpdate=dateadd("d",1,tmpdate)
			Next

			'response.write Str1010

			set rs=nothing
			set dic = nothing
		%>
		<tr valign=top> 
			  <td> 
			  
				<table border=1 cellpadding=2 cellspacing=2 width="100%" bgcolor="#FFFFF4" bordercolor="#326916" align="center">
				<tr valign=top> 
				<td bgcolor="#c1e0a3" align="center" colspan="2">���<br>�ɬq</td>
				<%=StrDateTop%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#c1e0a3" valign="center" rowspan="4">�W<br>��</td>
				<td bgcolor="#E5F6D4" align="center" >08:10<br>�x<br>09:00</td>
				<%=Str0810%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#E5F6D4" valign="center">09:10<br>�x<br>10:00</td>
				<%=Str0910%>
				</tr>
				<tr valign=top> 
				<td bgcolor="#E5F6D4" valign="center">10:10<br>�x<br>11:00</td>
				<%=Str1010%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#E5F6D4" valign="center">11:10<br>�x<br>12:00</td>
				<%=Str1110%>
				</tr>
				<tr valign=top  > 
				<td bgcolor="#FFFFF4" align="center" colspan="10">Lunch Recess</td>
				
				</tr>
				<tr valign=top > 
				<td bgcolor="#c1e0a3" valign="center" rowspan="5">�U<br>��</td>
				<td bgcolor="#E5F6D4" valign="center" >13:10<br>�x<br>14:00</td>
				<%=Str1310%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#E5F6D4" valign="center">14:10<br>�x<br>15:00</td>
				<%=Str1410%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#E5F6D4" valign="center" >15:10<br>�x<br>16:00</td>
				<%=Str1510%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#E5F6D4" valign="center" >16:10<br>�x<br>17:00</td>
				<%=Str1610%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#E5F6D4" valign="center" >17:10<br>�x<br>18:00</td>
				<%=Str1710%>
				</tr>
				</table>
			</td>                         
		  </tr>
		</table>
		<BR>
		</TD>
	</TR>
	</TABLE>
<!-- ---------------------------------------------------------------------------------------- -->
	</TD>
</TR>
<TR bgcolor="#333333" height="30">
	<TD class="T1">
	<!-- #include file="lib\bottom.inc" -->
	
	</TD>
</TR>
</TABLE>

</TD>
</TR>
</TABLE>

</BODY>
</HTML>
<!-- #INCLUDE virtual="/include/lib/conn/syconnclose.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->
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
response.Write("------------------" & session("classify"))
AColor="#FF0000"'預約顏色
PColor="blue"'審核通過
validate=trim(request("validate"))
sid=trim(request("sid"))
BOOK_DATE=trim(request("BOOK_DATE"))

'today1=year(dateadd("d",1,date()))-1911 & right("0" & month(dateadd("d",1,date())),2) & right("0" & day(dateadd("d",1,date())),2)
today=year(date())-1911 & right("0" & month(date()),2) & right("0" & day(date()),2)

if BOOK_DATE="" or isnull(BOOK_DATE) or isempty(BOOK_DATE) then
	BOOK_DATE=today

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
<TITLE> 【LDCC英外語能力診斷輔導中心】 </TITLE>
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

</script>
</HEAD>
<BODY bottomMargin=0 leftMargin=0 topMargin=0 rightMargin=0 marginheight="0" marginwidth="0"  bgcolor="#f4c60d">
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">查詢教師輔導療程班表 </TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>
	
	<TR>
		<TD></TD><TD valign="top">
		<form name="form1" method=post action="bookteacher.asp">
		<input type="hidden" value="CheckAccount" name="validate">
		<P><BR>
		<table border=0 cellpadding=0 cellspacing=2 width="98%"  align="center">
		<tr valign=top> 
		<td>
			<table border=0 cellpadding=0 cellspacing=2  align="left">
			
			<td>查詢日期：</td>                     
			<td> 
			<input type="text" value="<%=BOOK_DATE%>" maxlength="6" name="BOOK_DATE" id="BOOK_DATE" class="inputtext">
			<img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('BOOK_DATE')" class="showhand">
			</td>
			<td><input  type="submit" value="查詢" ></td>
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
			sql = "select * from boo_skill where yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				response.write "<font color='red'><B>" & rs("code") & "</B></font>：" & rs("name") & "&nbsp;"
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
'response.write Strnoopendate
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
				<font color="#CC9900">■</font>黃色代表已預約，<font color="blue">■</font>藍色代表可預約，<font color="#999999">■</font>灰色代表未在預約期限，<font color="#339900">■</font>綠色可當日親洽預約
				</td>
				<td>
					<table border=0 cellpadding=0 cellspacing=2 align="right" >
					<tr>
					<td align="right"  class="showhand" onclick="changedate('<%=datetoNumformat(dateadd("d",-7,NumberToDateFormat(BOOK_DATE)))%>')"><a href="#"><font color="red">&nbsp;<<前一週&nbsp;</font></a></td>
					<td>&nbsp;｜&nbsp;</td>
					<td align="right" class="showhand" onclick="changedate('<%=datetoNumformat(dateadd("d",7,NumberToDateFormat(BOOK_DATE)))%>')"><a href="#"><font color="red">&nbsp;後一週&nbsp;>></font></a></td>
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
			dic.Add "1","日"
			dic.Add "2","一"
			dic.Add "3","二"
			dic.Add "4","三"
			dic.Add "5","四"
			dic.Add "6","五"
			dic.Add "7","六"
			set rs = server.CreateObject("adodb.recordset")
			tmpdate=dateformat(BOOK_DATE)
			'response.Write(tmpdate)
			StrDateTop=""'標題的日期和星期
			StrStatus1010=""'10點
			StrStatus1010B=""

			StrStatus1110=""'11點
			StrStatus1110B=""
			
			'新增加中午時段, 2011/05/03, shihchi
			StrStatus1210=""'12點
			StrStatus1210B=""
			
			StrStatus1310=""'13點
			StrStatus1410=""'14點
			StrStatus1510=""'15點
			StrStatus1610=""'16點

			Str1010=""'10點
			Str1010B=""'10點
			Str1110=""'11點
			Str1110B=""
			
			'新增加中午時段, 2011/05/03, shihchi
			Str1210=""'12點
			Str1210B=""
			
			Str1310=""'13點
			Str1410=""'14點
			Str1510=""'15點
			Str1610=""'16點
			
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
				if ww="1" or ww="7"  then
					'Str1010 = Str1010 & "<td bgcolor="&tmpColor&">&nbsp;</td>"
					'Str1110 = Str1110 & "<td bgcolor="&tmpColor&">&nbsp;</td>"

					'Str1310 = Str1310 & "<td bgcolor="&tmpColor&">&nbsp;</td>"
					'Str1410 = Str1410 & "<td bgcolor="&tmpColor&">&nbsp;</td>"
					'Str1510 = Str1510 & "<td bgcolor="&tmpColor&">&nbsp;</td>"
					'Str1610 = Str1610 & "<td bgcolor="&tmpColor&">&nbsp;</td>"

				
				else
					'標題的日期和星期
					
					StrDateTop = StrDateTop & "<td bgcolor=""#c1e0a3"" align=""center"">" & tmpdate & "<br>星期" & dic.Item(cstr(ww)) & "</td>"
					'如果tmpdate在Strnoopendate裡面有的話, 就不顯示
					if  InStr(Strnoopendate,Cstr(tmpdate))  then 				
						Str1010 = Str1010 & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						Str1010B = Str1010B & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						Str1110 = Str1110 & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						Str1110B = Str1110B & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						
						'新增加中午時段, 2011/05/03, shihchi
						Str1210 = Str1210 & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						Str1210B = Str1210B & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						
						Str1310 = Str1310 & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						Str1310B = Str1310B & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						Str1410 = Str1410 & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						Str1410B = Str1410B & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						Str1510 = Str1510 & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						Str1510B = Str1510B & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						Str1610 = Str1610 & "<td bgcolor=#FFCCFF>&nbsp;</td>"
						Str1610B = Str1610B & "<td bgcolor=#FFCCFF>&nbsp;</td>"
					else
						'dd=year(tmpdate)-1911 & right("0" & month(tmpdate),2) & right("0" & day(tmpdate),2)
						'10點
						'上半節
						'sql = "select * from boo_schedule where category='T' and btime='1010' and bweek='"&cint(ww)-1&"'  and yn='Y'"
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername  and b.yn in ('Y','A')   and b.btime='1010' and b.timeflag in ('U','A') and b.bdate='"&datetoNumformat(tmpdate)&"' "
						sql = sql & " where b.pid is null and a.category='T' and a.btime='1010' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y'  and  a.yms='"&par_yms&"' "
						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1010=""
						
						while not rs.EOF
							if rs("name")="N" then
								if  ( cdbl(datetoNumformat(tmpdate)) <=  cdbl(datetoNumformat(dateadd("d",7,date())))  and  cdbl(datetoNumformat(tmpdate)) >  cdbl(datetoNumformat(date()))) or  session("classify")="A" then
									StrStatus1010 = StrStatus1010 & "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<a title='可預約' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1010&category=T&teachername="&rs("teacher")&"&scid="&rs("scid")&"&showflag=1'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								elseif  cdbl(datetoNumformat(tmpdate)) =  cdbl(datetoNumformat(date())) then
									StrStatus1010 = StrStatus1010& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#339900'>" & rs("teacher") & "</font>" & "<br>" 
								else
									StrStatus1010 = StrStatus1010& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'預約資訊
								StrStatus1010 = StrStatus1010 & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');""  title='預約人："&rs("name")&"("&rs("department")&"，"&rs("grade")&"年級，"&rs("class1")&"班)'><font color='#FF3300'>" & rs("skillcode") & "</font><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1010 = Str1010 & "<td bgcolor="&tmpColor&">" & StrStatus1010 & "&nbsp;</td>"
						'下半節
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername  and b.yn in ('Y','A')  and b.btime='1010' and b.timeflag in ('B','A') and b.bdate='"&datetoNumformat(tmpdate)&"' "
						sql = sql & " where b.pid is null and  a.category='T' and a.btime='1010' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y'   and  a.yms='"&par_yms&"' "
						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1010B=""
						
						while not rs.EOF
							if rs("name")="N" then
								if  ( cdbl(datetoNumformat(tmpdate))  <=  cdbl(datetoNumformat(dateadd("d",7,date())))   and   cdbl(datetoNumformat(tmpdate))  >  cdbl(datetoNumformat(date())) )  or  session("classify")="A"  then
									StrStatus1010B = StrStatus1010B & "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<a title='可預約' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1010&category=T&teachername="&rs("teacher")&"&scid="&rs("scid")&"&showflag=1'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								elseif  cdbl(datetoNumformat(tmpdate)) =  cdbl(datetoNumformat(date())) then
									StrStatus1010B = StrStatus1010B& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#339900'>" & rs("teacher") & "</font>" & "<br>" 
								else
									StrStatus1010B = StrStatus1010B & "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'預約資訊
								StrStatus1010B = StrStatus1010B & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');"" title='預約人："&rs("name")&"("&rs("department")&"，"&rs("grade")&"年級，"&rs("class1")&"班)'><font color='#FF3300'>" & rs("skillcode") & "</font><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1010B = Str1010B & "<td bgcolor="&tmpColor&">" & StrStatus1010B & "&nbsp;</td>"
						'11點
						'上半節
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername  and b.yn in ('Y','A')  and b.btime='1110' and b.timeflag in ('U','A') and b.bdate='"&datetoNumformat(tmpdate)&"' "
						sql = sql & " where b.pid is null and  a.category='T' and a.btime='1110' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y'   and  a.yms='"&par_yms&"' "

						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1110=""
						while not rs.EOF
							if rs("name")="N" then
								if ( cdbl(datetoNumformat(tmpdate)) <=  cdbl(datetoNumformat(dateadd("d",7,date())))  and   cdbl(datetoNumformat(tmpdate))  >  cdbl(datetoNumformat(date())) )  or  session("classify")="A"  then
									StrStatus1110 = StrStatus1110 & "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<a title='可預約' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1110&category=T&teachername="&rs("teacher")&"&scid="&rs("scid")&"&showflag=1'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								elseif  cdbl(datetoNumformat(tmpdate)) =  cdbl(datetoNumformat(date())) then
									StrStatus1110 = StrStatus1110& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#339900'>" & rs("teacher") & "</font>" & "<br>" 
								else
									StrStatus1110 = StrStatus1110& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'預約資訊
								StrStatus1110 = StrStatus1110 & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');"" title='預約人："&rs("name")&"("&rs("department")&"，"&rs("grade")&"年級，"&rs("class1")&"班)'><font color='#FF3300'>" & rs("skillcode") & "</font><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1110 = Str1110 & "<td bgcolor="&tmpColor&">" & StrStatus1110 & "&nbsp;</td>"

						'下半節
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername  and b.yn in ('Y','A')  and b.btime='1110' and b.timeflag in ('B','A') and b.bdate='"&datetoNumformat(tmpdate)&"' "
						sql = sql & " where b.pid is null and  a.category='T' and a.btime='1110' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y'   and  a.yms='"&par_yms&"' "

						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1110B=""
						while not rs.EOF
							if rs("name")="N" then
								if ( cdbl(datetoNumformat(tmpdate)) <=  cdbl(datetoNumformat(dateadd("d",7,date())))  and  cdbl(datetoNumformat(tmpdate)) >  cdbl(datetoNumformat(date())) )  or  session("classify")="A"  then
									StrStatus1110B = StrStatus1110B & "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<a title='可預約' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1110&category=T&teachername="&rs("teacher")&"&scid="&rs("scid")&"&showflag=1'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								elseif  cdbl(datetoNumformat(tmpdate)) =  cdbl(datetoNumformat(date())) then
									StrStatus1110B = StrStatus1110B& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#339900'>" & rs("teacher") & "</font>" & "<br>" 
								else
									StrStatus1110B = StrStatus1110B& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'預約資訊
								StrStatus1110B = StrStatus1110B & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');"" title='預約人："&rs("name")&"("&rs("department")&"，"&rs("grade")&"年級，"&rs("class1")&"班)'><font color='#FF3300'>" & rs("skillcode") & "</font><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1110B = Str1110B & "<td bgcolor="&tmpColor&">" & StrStatus1110B & "&nbsp;</td>"
						
						
						
						
						
						'新增加中午時段, 2011/05/03, shihchi
						'12點
						'上半節
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername  and b.yn in ('Y','A')  and b.btime='1210' and b.timeflag in ('U','A') and b.bdate='"&datetoNumformat(tmpdate)&"' "
						sql = sql & " where b.pid is null and  a.category='T' and a.btime='1210' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y'   and  a.yms='"&par_yms&"' "

						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1210=""
						while not rs.EOF
							if rs("name")="N" then
								if ( cdbl(datetoNumformat(tmpdate)) <=  cdbl(datetoNumformat(dateadd("d",7,date())))  and   cdbl(datetoNumformat(tmpdate))  >  cdbl(datetoNumformat(date())) )  or  session("classify")="A"  then
									StrStatus1210 = StrStatus1210 & "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<a title='可預約' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1210&category=T&teachername="&rs("teacher")&"&scid="&rs("scid")&"&showflag=1'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								elseif  cdbl(datetoNumformat(tmpdate)) =  cdbl(datetoNumformat(date())) then
									StrStatus1210 = StrStatus1210& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#339900'>" & rs("teacher") & "</font>" & "<br>" 
								else
									StrStatus1210 = StrStatus1210& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'預約資訊
								StrStatus1210 = StrStatus1210 & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');"" title='預約人："&rs("name")&"("&rs("department")&"，"&rs("grade")&"年級，"&rs("class1")&"班)'><font color='#FF3300'>" & rs("skillcode") & "</font><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1210 = Str1210 & "<td bgcolor="&tmpColor&">" & StrStatus1210 & "&nbsp;</td>"

						'下半節
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername  and b.yn in ('Y','A')  and b.btime='1210' and b.timeflag in ('B','A') and b.bdate='"&datetoNumformat(tmpdate)&"' "
						sql = sql & " where b.pid is null and  a.category='T' and a.btime='1210' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y'   and  a.yms='"&par_yms&"' "

						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1210B=""
						while not rs.EOF
							if rs("name")="N" then
								if ( cdbl(datetoNumformat(tmpdate)) <=  cdbl(datetoNumformat(dateadd("d",7,date())))  and  cdbl(datetoNumformat(tmpdate)) >  cdbl(datetoNumformat(date())) )  or  session("classify")="A"  then
									StrStatus1210B = StrStatus1210B & "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<a title='可預約' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1210&category=T&teachername="&rs("teacher")&"&scid="&rs("scid")&"&showflag=1'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								elseif  cdbl(datetoNumformat(tmpdate)) =  cdbl(datetoNumformat(date())) then
									StrStatus1210B = StrStatus1210B& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#339900'>" & rs("teacher") & "</font>" & "<br>" 
								else
									StrStatus1210B = StrStatus1210B& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'預約資訊
								StrStatus1210B = StrStatus1210B & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');"" title='預約人："&rs("name")&"("&rs("department")&"，"&rs("grade")&"年級，"&rs("class1")&"班)'><font color='#FF3300'>" & rs("skillcode") & "</font><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1210B = Str1210B & "<td bgcolor="&tmpColor&">" & StrStatus1210B & "&nbsp;</td>"
						
						
						
						
						

						'13點
						'上半節
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername  and b.yn in ('Y','A')  and b.btime='1310' and b.timeflag in ('U','A') and b.bdate='"&datetoNumformat(tmpdate)&"' "
						sql = sql & " where b.pid is null and  a.category='T' and a.btime='1310' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y'  and  a.yms='"&par_yms&"' "
						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1310=""
						while not rs.EOF
							if rs("name")="N" then
								if  ( cdbl(datetoNumformat(tmpdate)) <=  cdbl(datetoNumformat(dateadd("d",7,date())))  and   cdbl(datetoNumformat(tmpdate)) >  cdbl(datetoNumformat(date())) )  or  session("classify")="A"  then
									StrStatus1310 = StrStatus1310 & "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<a title='可預約' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1310&category=T&teachername="&rs("teacher")&"&scid="&rs("scid")&"&showflag=1'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								elseif  cdbl(datetoNumformat(tmpdate)) =  cdbl(datetoNumformat(date())) then
									StrStatus1310 = StrStatus1310& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#339900'>" & rs("teacher") & "</font>" & "<br>" 
								else
									StrStatus1310 = StrStatus1310& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'預約資訊
								StrStatus1310 = StrStatus1310 & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');"" title='預約人："&rs("name")&"("&rs("department")&"，"&rs("grade")&"年級，"&rs("class1")&"班)'><font color='#FF3300'>" & rs("skillcode") & "</font><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1310 = Str1310 & "<td bgcolor="&tmpColor&">" & StrStatus1310 & "&nbsp;</td>"

						'下半節
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername  and b.yn in ('Y','A')  and b.btime='1310' and b.timeflag in ('B','A') and b.bdate='"&datetoNumformat(tmpdate)&"' "
						sql = sql & " where b.pid is null and  a.category='T' and a.btime='1310' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y'  and  a.yms='"&par_yms&"' "

						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1310B=""
						while not rs.EOF
							if rs("name")="N" then
								if  ( cdbl(datetoNumformat(tmpdate))  <=  cdbl(datetoNumformat(dateadd("d",7,date())))  and  cdbl(datetoNumformat(tmpdate)) >  cdbl(datetoNumformat(date())))   or  session("classify")="A"   then
									StrStatus1310B = StrStatus1310B & "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<a title='可預約' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1310&category=T&teachername="&rs("teacher")&"&scid="&rs("scid")&"&showflag=1'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								elseif  cdbl(datetoNumformat(tmpdate)) =  cdbl(datetoNumformat(date())) then
									StrStatus1310B = StrStatus1310B& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#339900'>" & rs("teacher") & "</font>" & "<br>" 
								else
									StrStatus1310B = StrStatus1310B& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'預約資訊
								StrStatus1310B = StrStatus1310B & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');"" title='預約人："&rs("name")&"("&rs("department")&"，"&rs("grade")&"年級，"&rs("class1")&"班)'><font color='#FF3300'>" & rs("skillcode") & "</font><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1310B = Str1310B & "<td bgcolor="&tmpColor&">" & StrStatus1310B & "&nbsp;</td>"
						
						'14點
						'上半節
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername  and b.yn in ('Y','A')  and b.btime='1410' and b.timeflag in ('U','A') and b.bdate='"&datetoNumformat(tmpdate)&"' "
						sql = sql & " where b.pid is null and  a.category='T' and a.btime='1410' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y'  and  a.yms='"&par_yms&"' "
						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1410=""
						while not rs.EOF
							if rs("name")="N" then
								if ( cdbl(datetoNumformat(tmpdate)) <=  cdbl(datetoNumformat(dateadd("d",7,date())))  and  cdbl(datetoNumformat(tmpdate))  > cdbl(datetoNumformat(date())) )   or  session("classify")="A"  then
									StrStatus1410 = StrStatus1410 & "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<a title='可預約' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1410&category=T&teachername="&rs("teacher")&"&scid="&rs("scid")&"&showflag=1'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								elseif  cdbl(datetoNumformat(tmpdate)) =  cdbl(datetoNumformat(date())) then
									StrStatus1410 = StrStatus1410& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#339900'>" & rs("teacher") & "</font>" & "<br>" 
								else
									StrStatus1410 = StrStatus1410& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'預約資訊
								StrStatus1410 = StrStatus1410 & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');"" title='預約人："&rs("name")&"("&rs("department")&"，"&rs("grade")&"年級，"&rs("class1")&"班)'><font color='#FF3300'>" & rs("skillcode") & "</font><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1410 = Str1410 & "<td bgcolor="&tmpColor&">" & StrStatus1410 & "&nbsp;</td>"

						'下半節
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername  and b.yn in ('Y','A')  and b.btime='1410' and b.timeflag in ('B','A') and b.bdate='"&datetoNumformat(tmpdate)&"' "
						sql = sql & " where b.pid is null and  a.category='T' and a.btime='1410' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y'  and  a.yms='"&par_yms&"' "

						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1410B=""
						while not rs.EOF
							if rs("name")="N" then
								if  ( cdbl(datetoNumformat(tmpdate))  <=  cdbl(datetoNumformat(dateadd("d",7,date())))  and  cdbl(datetoNumformat(tmpdate))  >  cdbl(datetoNumformat(date())) )   or  session("classify")="A"   then
									StrStatus1410B = StrStatus1410B & "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<a title='可預約' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1410&category=T&teachername="&rs("teacher")&"&scid="&rs("scid")&"&showflag=1'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								elseif  cdbl(datetoNumformat(tmpdate)) =  cdbl(datetoNumformat(date())) then
									StrStatus1410B = StrStatus1410B& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#339900'>" & rs("teacher") & "</font>" & "<br>" 
								else
									StrStatus1410B = StrStatus1410B& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'預約資訊
								StrStatus1410B = StrStatus1410B & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');"" title='預約人："&rs("name")&"("&rs("department")&"，"&rs("grade")&"年級，"&rs("class1")&"班)'><font color='#FF3300'>" & rs("skillcode") & "</font><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1410B = Str1410B & "<td bgcolor="&tmpColor&">" & StrStatus1410B & "&nbsp;</td>"



						
						'15點
						'上半節
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername  and b.yn in ('Y','A')  and b.btime='1510' and b.timeflag in ('U','A') and b.bdate='"&datetoNumformat(tmpdate)&"' "
						sql = sql & " where b.pid is null and  a.category='T' and a.btime='1510' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y'  and  a.yms='"&par_yms&"' "
						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1510=""
						while not rs.EOF
							if rs("name")="N" then
								if  ( cdbl(datetoNumformat(tmpdate)) <=  cdbl(datetoNumformat(dateadd("d",7,date())))  and  cdbl(datetoNumformat(tmpdate))  >  cdbl(datetoNumformat(date())) )   or  session("classify")="A"  then
									StrStatus1510 = StrStatus1510 & "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<a title='可預約' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1510&category=T&teachername="&rs("teacher")&"&scid="&rs("scid")&"&showflag=1'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								elseif  cdbl(datetoNumformat(tmpdate)) =  cdbl(datetoNumformat(date())) then
									StrStatus1510 = StrStatus1510& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#339900'>" & rs("teacher") & "</font>" & "<br>" 
								else
									StrStatus1510 = StrStatus1510& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'預約資訊
								StrStatus1510 = StrStatus1510 & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');"" title='預約人："&rs("name")&"("&rs("department")&"，"&rs("grade")&"年級，"&rs("class1")&"班)'><font color='#FF3300'>" & rs("skillcode") & "</font><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1510 = Str1510 & "<td bgcolor="&tmpColor&">" & StrStatus1510 & "&nbsp;</td>"

						'下半節
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername  and b.yn in ('Y','A')  and b.btime='1510' and b.timeflag in ('B','A') and b.bdate='"&datetoNumformat(tmpdate)&"' "
						sql = sql & " where b.pid is null and  a.category='T' and a.btime='1510' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y'  and  a.yms='"&par_yms&"' "

						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1510B=""
						while not rs.EOF
							if rs("name")="N" then
								if  ( cdbl(datetoNumformat(tmpdate)) <=  cdbl(datetoNumformat(dateadd("d",7,date())))  and  cdbl(datetoNumformat(tmpdate))  >  cdbl(datetoNumformat(date())))   or  session("classify")="A"   then
									StrStatus1510B = StrStatus1510B & "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<a title='可預約' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1510&category=T&teachername="&rs("teacher")&"&scid="&rs("scid")&"&showflag=1'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								elseif  cdbl(datetoNumformat(tmpdate)) =  cdbl(datetoNumformat(date())) then
									StrStatus1510B = StrStatus1510B& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#339900'>" & rs("teacher") & "</font>" & "<br>" 
								else
									StrStatus1510B = StrStatus1510B& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'預約資訊
								StrStatus1510B = StrStatus1510B & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');"" title='預約人："&rs("name")&"("&rs("department")&"，"&rs("grade")&"年級，"&rs("class1")&"班)'><font color='#FF3300'>" & rs("skillcode") & "</font><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1510B = Str1510B & "<td bgcolor="&tmpColor&">" & StrStatus1510B & "&nbsp;</td>"
						'16點
						'上半節
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername  and b.yn in ('Y','A')  and b.btime='1610' and b.timeflag in ('U','A') and b.bdate='"&datetoNumformat(tmpdate)&"' "
						sql = sql & " where b.pid is null and  a.category='T' and a.btime='1610' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y'  and  a.yms='"&par_yms&"' "
						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1610=""
						while not rs.EOF
							if rs("name")="N" then
								if  ( cdbl(datetoNumformat(tmpdate)) <=  cdbl(datetoNumformat(dateadd("d",7,date())))  and  cdbl(datetoNumformat(tmpdate))  >  cdbl(datetoNumformat(date())))  or  session("classify")="A"   then
									StrStatus1610 = StrStatus1610 & "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<a title='可預約' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1610&category=T&teachername="&rs("teacher")&"&scid="&rs("scid")&"&showflag=1'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								elseif  cdbl(datetoNumformat(tmpdate)) =  cdbl(datetoNumformat(date())) then
									StrStatus1610 = StrStatus1610& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#339900'>" & rs("teacher") & "</font>" & "<br>" 
								else
									StrStatus1610 = StrStatus1610& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'預約資訊
								StrStatus1610 = StrStatus1610 & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');"" title='預約人："&rs("name")&"("&rs("department")&"，"&rs("grade")&"年級，"&rs("class1")&"班)'><font color='#FF3300'>" & rs("skillcode") & "</font><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1610 = Str1610 & "<td bgcolor="&tmpColor&">" & StrStatus1610 & "&nbsp;</td>"

						'下半節
						sql ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate  from boo_schedule a "
						sql = sql & "left join boo_book_T_M b on a.teacher=b.teachername  and b.yn in ('Y','A')  and b.btime='1610' and b.timeflag in ('B','A') and b.bdate='"&datetoNumformat(tmpdate)&"' "
						sql = sql & " where b.pid is null and  a.category='T' and a.btime='1610' and a.bweek='"&cint(ww)-1&"'  and a.yn='Y'  and  a.yms='"&par_yms&"' "

						'response.write sql & "<br>"
						rs.Open sql,msconn,adOpenStatic,adLockReadonly
						StrStatus1610B=""
						while not rs.EOF
							if rs("name")="N" then
								if  ( cdbl(datetoNumformat(tmpdate))  <=  cdbl(datetoNumformat(dateadd("d",7,date())))  and  cdbl(datetoNumformat(tmpdate))  >  cdbl(datetoNumformat(date())) )   or  session("classify")="A"  then
									StrStatus1610B = StrStatus1610B & "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<a title='可預約' href='bookingt.asp?bdate="&datetoNumformat(tmpdate)&"&btime=1610&category=T&teachername="&rs("teacher")&"&scid="&rs("scid")&"&showflag=1'><font color='blue'>" & rs("teacher") & "</font></a><br>" 
								elseif  cdbl(datetoNumformat(tmpdate)) =  cdbl(datetoNumformat(date())) then
									StrStatus1610B = StrStatus1610B& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#339900'>" & rs("teacher") & "</font>" & "<br>" 
								else
									StrStatus1610B = StrStatus1610B& "<font color='#FF3300'>" & rs("skillcode") & "</font>"  & "<font color='#999999'>" & rs("teacher") & "</font>" & "<br>" 
								end if
							else
								'預約資訊
								StrStatus1610B = StrStatus1610B & "<a href='#' onclick=""window.showModalDialog('bookingtveiw.asp?id="&rs("bookid")&"','','dialogWidth=650px;dialogHeight=650px;status=no');"" title='預約人："&rs("name")&"("&rs("department")&"，"&rs("grade")&"年級，"&rs("class1")&"班)'><font color='#FF3300'>" & rs("skillcode") & "</font><font color='#CC9900'>"  & rs("teacher") & "</font></a>" & "<br>" 
							end if
							
							rs.MoveNext
						wend 
						rs.close
						Str1610B = Str1610B & "<td bgcolor="&tmpColor&">" & StrStatus1610B & "&nbsp;</td>"
			
					end if
				
				end if 'if ww="1" or ww="7" then
				
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
				<td bgcolor="#c1e0a3" align="center" colspan="3">日期<br>時段</td>
				<%=StrDateTop%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#c1e0a3" valign="center" rowspan="4">上<br>午</td>
				<td bgcolor="#E5F6D4" align="center" rowspan="2">10:10<br>│<br>11:00</td>
				<td bgcolor="#E5F6D4" align="center">10:10<br>│<br>10:35</td>
				<%=Str1010%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#E5F6D4" valign="center">10:35<br>│<br>11:00</td>
				<%=Str1010B%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#E5F6D4" valign="center" rowspan="2">11:10<br>│<br>12:00</td>
				<td bgcolor="#E5F6D4" valign="center">11:10<br>│<br>11:35</td>
				<%=Str1110%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#E5F6D4" valign="center">11:35<br>│<br>12:00</td>
				<%=Str1110B%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#c1e0a3" valign="center" rowspan="2">中<br>午</td>
				<td bgcolor="#E5F6D4" valign="center" rowspan="2">12:10<br>│<br>13:00</td>
				<td bgcolor="#E5F6D4" valign="center">12:10<br>│<br>12:35</td>
				<%=Str1210%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#E5F6D4" valign="center">12:35<br>│<br>13:00</td>
				<%=Str1210B%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#c1e0a3" valign="center" rowspan="10">下<br>午</td>
				<td bgcolor="#E5F6D4" valign="center" rowspan="2">13:10<br>│<br>14:00</td>
				<td bgcolor="#E5F6D4" valign="center">13:10<br>│<br>13:35</td>
				<%=Str1310%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#E5F6D4" valign="center">13:35<br>│<br>14:00</td>
				<%=Str1310B%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#E5F6D4" valign="center" rowspan="2">14:10<br>│<br>15:00</td>
				<td bgcolor="#E5F6D4" valign="center">14:10<br>│<br>14:35</td>
				<%=Str1410%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#E5F6D4" valign="center">14:35<br>│<br>15:00</td>
				<%=Str1410B%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#E5F6D4" valign="center" rowspan="2">15:10<br>│<br>16:00</td>
				<td bgcolor="#E5F6D4" valign="center">15:10<br>│<br>15:35</td>
				<%=Str1510%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#E5F6D4" valign="center">15:35<br>│<br>16:00</td>
				<%=Str1510B%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#E5F6D4" valign="center" rowspan="2">16:10<br>│<br>17:00</td>
				<td bgcolor="#E5F6D4" valign="center">16:10<br>│<br>16:35</td>
				<%=Str1610%>
				</tr>
				<tr valign=top > 
				<td bgcolor="#E5F6D4" valign="center">16:35<br>│<br>17:00</td>
				<%=Str1610B%>
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
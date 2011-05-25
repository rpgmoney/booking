<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE file="include/conn/syconn.asp" -->
<!-- #INCLUDE file="include/conn/msconn.asp" -->
<!-- #INCLUDE file="include/asp/jsCalendar.asp" -->
<!-- #INCLUDE file="include/inc/lib.inc" -->
<%
validate=trim(request("validate"))
sid=trim(request("sid"))
BOOK_DATE=trim(request("BOOK_DATE"))

today=year(date())-1911 & right("0" & month(date()),2) & right("0" & day(date()),2)

if BOOK_DATE="" or isnull(BOOK_DATE) or isempty(BOOK_DATE) then
	BOOK_DATE=today
end if




set rs = server.CreateObject("adodb.recordset")


function dateformat(vdate)
	if vdate<>"" then
		dateformat=cint(left(vdate,2))+1911 & "/" & mid(vdate,3,2) & "/" & right(vdate,2)
	end if
end function
function datetoNumformat(vdate)
	if vdate<>"" then
		datetoNumformat=Year(vdate)-1911 & right("0"& month(vdate),2) & right("0" & day(vdate),2)
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
			
		<%
			
			set rs = server.CreateObject("adodb.recordset")
			'星期一
			str1010_1=""
			sql = "select * from boo_schedule where category='T' and btime='1010' and bweek='1' and yn='Y' "
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1010_1 = str1010_1 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1110_1=""
			sql = "select * from boo_schedule where category='T' and btime='1110' and bweek='1'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1110_1 = str1110_1 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1310_1=""
			sql = "select * from boo_schedule where category='T' and btime='1310' and bweek='1'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1310_1 = str1310_1 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1410_1=""
			sql = "select * from boo_schedule where category='T' and btime='1410' and bweek='1'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1410_1 = str1410_1 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1510_1=""
			sql = "select * from boo_schedule where category='T' and btime='1510' and bweek='1'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1510_1 = str1510_1 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1610_1=""
			sql = "select * from boo_schedule where category='T' and btime='1610' and bweek='1'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1610_1 = str1610_1 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			'星期二
			str1010_2=""
			sql = "select * from boo_schedule where category='T' and btime='1010' and bweek='2'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1010_2 = str1010_2 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1110_2=""
			sql = "select * from boo_schedule where category='T' and btime='1110' and bweek='2'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1110_2 = str1110_2 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1310_2=""
			sql = "select * from boo_schedule where category='T' and btime='1310' and bweek='2'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1310_2 = str1310_2 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1410_2=""
			sql = "select * from boo_schedule where category='T' and btime='1410' and bweek='2'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1410_2 = str1410_2 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1510_2=""
			sql = "select * from boo_schedule where category='T' and btime='1510' and bweek='2'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1510_2 = str1510_2 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1610_2=""
			sql = "select * from boo_schedule where category='T' and btime='1610' and bweek='2'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1610_2 = str1610_2 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			'星期三
			str1010_3=""
			sql = "select * from boo_schedule where category='T' and btime='1010' and bweek='3'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1010_3 = str1010_3 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1110_3=""
			sql = "select * from boo_schedule where category='T' and btime='1110' and bweek='3'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1110_3 = str1110_3 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1310_3=""
			sql = "select * from boo_schedule where category='T' and btime='1310' and bweek='3'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1310_3 = str1310_3 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1410_3=""
			sql = "select * from boo_schedule where category='T' and btime='1410' and bweek='3'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1410_3 = str1410_3 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1510_3=""
			sql = "select * from boo_schedule where category='T' and btime='1510' and bweek='3'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1510_3 = str1510_3 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1610_3=""
			sql = "select * from boo_schedule where category='T' and btime='1610' and bweek='3'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1610_3 = str1610_3 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			'星期四
			str1010_4=""
			sql = "select * from boo_schedule where category='T' and btime='1010' and bweek='4'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1010_4 = str1010_4 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1110_4=""
			sql = "select * from boo_schedule where category='T' and btime='1110' and bweek='4'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1110_4 = str1110_4 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1310_4=""
			sql = "select * from boo_schedule where category='T' and btime='1310' and bweek='4'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1310_4 = str1310_4 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1410_4=""
			sql = "select * from boo_schedule where category='T' and btime='1410' and bweek='4'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1410_4 = str1410_4 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1510_4=""
			sql = "select * from boo_schedule where category='T' and btime='1510' and bweek='4'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1510_4 = str1510_4 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1610_4=""
			sql = "select * from boo_schedule where category='T' and btime='1610' and bweek='4'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1610_4 = str1610_4 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			'星期五
			str1010_5=""
			sql = "select * from boo_schedule where category='T' and btime='1010' and bweek='5'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1010_5 = str1010_5 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1110_5=""
			sql = "select * from boo_schedule where category='T' and btime='1110' and bweek='5'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1110_5 = str1110_5 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1310_5=""
			sql = "select * from boo_schedule where category='T' and btime='1310' and bweek='5'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1310_5 = str1310_5 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1410_5=""
			sql = "select * from boo_schedule where category='T' and btime='1410' and bweek='5'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1410_5 = str1410_5 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1510_5=""
			sql = "select * from boo_schedule where category='T' and btime='1510' and bweek='5'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1510_5 = str1510_5 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			str1610_5=""
			sql = "select * from boo_schedule where category='T' and btime='1610' and bweek='5'  and yn='Y'"
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			while not rs.EOF
				str1610_5 = str1610_5 & rs("teacher") & "<br>"
				rs.MoveNext
			wend 
			rs.close
			set rs=nothing
			
		%>
	
			  
				<table border=1 cellpadding=0 cellspacing=2 width="720" bgcolor="#FFFFF4" bordercolor="#326916" align="center">
				<tr valign=top align="center"> 
				<td bgcolor="#c1e0a3" align="center" colspan="3" width="120">星期<br>時段</td>
				<td bgcolor="#E5F6D4" width="120">Monday<br>星期一</td><td bgcolor="#E5F6D4" width="120">Tuesday<br>星期二</td><td bgcolor="#E5F6D4" width="120">Wednesday<br>星期三</td><td bgcolor="#E5F6D4" width="120">Thursday<br>星期四</td><td bgcolor="#E5F6D4" width="120">Friday<br>星期五</td>
				</tr>
				<tr valign=top  > 
				<td bgcolor="#c1e0a3" align="center" rowspan="2">上午</td>
				<td bgcolor="#E5F6D4" align="center">10:10</td>
				<td bgcolor="#E5F6D4" align="center">11:00</td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1010_1%></td>
				<td bgcolor="#FFFFF4" align="center">&nbsp;<%=str1010_2%></td>
				<td bgcolor="#FFFFF4" align="center">&nbsp;<%=str1010_3%></td>
				<td bgcolor="#FFFFF4" align="center">&nbsp;<%=str1010_4%></td>
				<td bgcolor="#FFFFF4" align="center">&nbsp;<%=str1010_5%></td>
				</tr>
				<tr valign=top > 
				<td bgcolor="#E5F6D4" align="center">11:10</td>
				<td bgcolor="#E5F6D4" align="center">12:00</td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1110_1%></td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1110_2%></td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1110_3%></td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1110_4%></td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1110_5%></td>
				</tr>
				<tr valign=top  > 
				<td bgcolor="#FFFFF4" align="center" colspan="8">Lunch Recess</td>
				
				</tr>
				<tr valign=top  > 
				<td bgcolor="#c1e0a3" align="center" rowspan="5">下午</td>
				<td bgcolor="#E5F6D4" align="center">13:10</td>
				<td bgcolor="#E5F6D4" align="center">14:00</td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1310_1%></td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1310_2%></td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1310_3%></td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1310_4%></td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1310_5%></td>
				</tr>
				<tr valign=top > 
				<td bgcolor="#E5F6D4" align="center">14:10</td>
				<td bgcolor="#E5F6D4" align="center">15:00</td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1410_1%></td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1410_2%></td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1410_3%></td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1410_4%></td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1410_5%></td>
				</tr>
				<tr valign=top > 
				<td bgcolor="#E5F6D4" align="center">15:10</td>
				<td bgcolor="#E5F6D4" align="center">16:00</td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1510_1%></td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1510_2%></td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1510_3%></td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1510_4%></td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1510_5%></td>
				</tr>
				<tr valign=top  > 
				<td bgcolor="#E5F6D4" align="center">16:10</td>
				<td bgcolor="#E5F6D4" align="center">17:00</td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1610_1%></td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1610_2%></td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1610_3%></td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1610_4%></td>
				<td bgcolor="#FFFFF4" align="center" >&nbsp;<%=str1610_5%></td>
				</tr>
				
				</table>
		


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
<!-- #INCLUDE file="include/conn/syconnclose.asp" -->
<!-- #INCLUDE file="include/conn/msconnclose.asp" -->
<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<!-- #INCLUDE file="lib/parameter.inc" -->

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
	<TR >
		<TD></TD><TD>
		<%
		sql = "select * from boo_skill where yn='Y'"
		rs.Open sql,msconn,adOpenStatic,adLockReadonly
		while not rs.EOF
			response.write "&nbsp;&nbsp;<font color='#CC9900'><B>" & rs("code") & "：" & rs("name") & "</B></font>"
			rs.MoveNext
		wend


		rs.close

		%>
		
		
		</TD>
	</TR>
	<TR>
		<TD></TD><TD valign="top">
			
		<%
			
			StrDateTop=""'標題的日期和星期
			StrStatus1010=""'10點

			StrStatus1110=""'11點
			StrStatus1310=""'13點
			StrStatus1410=""'14點
			StrStatus1510=""'15點
			StrStatus1610=""'16點

			Str1010=""'10點
			Str1110=""'11點
			Str1310=""'13點
			Str1410=""'14點
			Str1510=""'15點
			Str1610=""'16點
			
			'新增加中午時段, 2011/05/04, shihchi
			StrStatus1210=""'12點
			Str1210=""'12點
			
			'response.write "BOOK_ROOM=" & BOOK_ROOM
			
			for i=1 to 5
				ww = i
				tmpColor="#FFFFF4"
				
					
				
				
				'10點
				sql = "select * from boo_schedule where category='T' and btime='1010' and bweek='"&cint(ww)&"'  and yn='Y' and  yms='"&par_yms&"' "

				'response.write sql & "<br>"
				rs.Open sql,msconn,adOpenStatic,adLockReadonly
				StrStatus1010=""
				while not rs.EOF
					StrStatus1010 = StrStatus1010 & "<font color='#FF3300'>" & rs("skillcode") & "</font>" & rs("teacher") & "<br>" 
					rs.MoveNext
				wend 
				rs.close
				Str1010 = Str1010 & "<td bgcolor="&tmpColor&">" & StrStatus1010 & "&nbsp;</td>"
				'11點
				sql = "select * from boo_schedule where category='T' and btime='1110' and bweek='"&cint(ww)&"'  and yn='Y' and  yms='"&par_yms&"'"

				'response.write sql & "<br>"
				rs.Open sql,msconn,adOpenStatic,adLockReadonly
				StrStatus1110=""
				while not rs.EOF
					StrStatus1110 = StrStatus1110 & "<font color='#FF3300'>" & rs("skillcode") & "</font>" & rs("teacher") & "<br>" 
					rs.MoveNext
				wend 
				rs.close
				Str1110 = Str1110 & "<td bgcolor="&tmpColor&">" & StrStatus1110 & "&nbsp;</td>"
				
				
				
				'新增加中午時段, 2011/05/04, shihchi
				'12點
				sql = "select * from boo_schedule where category='T' and btime='1210' and bweek='"&cint(ww)&"'  and yn='Y' and  yms='"&par_yms&"'"

				'response.write sql & "<br>"
				rs.Open sql,msconn,adOpenStatic,adLockReadonly
				StrStatus1210=""
				while not rs.EOF
					StrStatus1210 = StrStatus1210 & "<font color='#FF3300'>" & rs("skillcode") & "</font>" & rs("teacher") & "<br>" 
					rs.MoveNext
				wend 
				rs.close
				Str1210 = Str1210 & "<td bgcolor="&tmpColor&">" & StrStatus1210 & "&nbsp;</td>"
				
				
				
				
				
				
				'13點
				sql = "select * from boo_schedule where category='T' and btime='1310' and bweek='"&cint(ww)&"'  and yn='Y' and  yms='"&par_yms&"'"
				'response.write sql & "<br>"
				rs.Open sql,msconn,adOpenStatic,adLockReadonly
				StrStatus1310=""
				while not rs.EOF
					StrStatus1310 = StrStatus1310 & "<font color='#FF3300'>" & rs("skillcode") & "</font>" & rs("teacher") & "<br>" 
					rs.MoveNext
				wend 
				rs.close
				Str1310 = Str1310 & "<td bgcolor="&tmpColor&">" & StrStatus1310 & "&nbsp;</td>"
				
				'14點
				sql = "select * from boo_schedule where category='T' and btime='1410' and bweek='"&cint(ww)&"'  and yn='Y' and  yms='"&par_yms&"'"
				'response.write sql & "<br>"
				rs.Open sql,msconn,adOpenStatic,adLockReadonly
				StrStatus1410=""
				while not rs.EOF
					StrStatus1410 = StrStatus1410 & "<font color='#FF3300'>" & rs("skillcode") & "</font>" & rs("teacher") & "<br>" 
					rs.MoveNext
				wend 
				rs.close
				Str1410 = Str1410 & "<td bgcolor="&tmpColor&">" & StrStatus1410 & "&nbsp;</td>"
				'15點
				sql = "select * from boo_schedule where category='T' and btime='1510' and bweek='"&cint(ww)&"'  and yn='Y' and  yms='"&par_yms&"'"
				'response.write sql & "<br>"
				rs.Open sql,msconn,adOpenStatic,adLockReadonly
				StrStatus1510=""
				while not rs.EOF
					StrStatus1510 = StrStatus1510 & "<font color='#FF3300'>" & rs("skillcode") & "</font>" & rs("teacher") & "<br>" 
					rs.MoveNext
				wend 
				rs.close
				Str1510 = Str1510 & "<td bgcolor="&tmpColor&">" & StrStatus1510 & "&nbsp;</td>"
				'16點
				sql = "select * from boo_schedule where category='T' and btime='1610' and bweek='"&cint(ww)&"'  and yn='Y' and  yms='"&par_yms&"'"
				'response.write sql & "<br>"
				rs.Open sql,msconn,adOpenStatic,adLockReadonly
				StrStatus1610=""
				while not rs.EOF
					StrStatus1610 = StrStatus1610 & "<font color='#FF3300'>" & rs("skillcode") & "</font>" & rs("teacher") & "<br>" 
					rs.MoveNext
				wend 
				rs.close
				Str1610 = Str1610 & "<td bgcolor="&tmpColor&">" & StrStatus1610 & "&nbsp;</td>"
				
			Next
			set rs=nothing
			
		%>
	
			  
				<table border=1 cellpadding=2 cellspacing=2 width="720" bgcolor="#FFFFF4" bordercolor="#326916" align="center">
				<tr valign=top align="center"> 
				<td bgcolor="#c1e0a3" align="center" colspan="3" width="120">星期<br>時段</td>
				<td bgcolor="#E5F6D4" width="120">Monday<br>星期一</td><td bgcolor="#E5F6D4" width="120">Tuesday<br>星期二</td><td bgcolor="#E5F6D4" width="120">Wednesday<br>星期三</td><td bgcolor="#E5F6D4" width="120">Thursday<br>星期四</td><td bgcolor="#E5F6D4" width="120">Friday<br>星期五</td>
				</tr>
				<tr valign=top  height="50"> 
				<td bgcolor="#c1e0a3" align="center" rowspan="2">上午</td>
				<td bgcolor="#E5F6D4" align="center">10:10</td>
				<td bgcolor="#E5F6D4" align="center">11:00</td>
				<%=Str1010%>
				</tr>
				<tr valign=top height="50"> 
				<td bgcolor="#E5F6D4" align="center">11:10</td>
				<td bgcolor="#E5F6D4" align="center">12:00</td>
				<%=Str1110%>
				</tr>
				<tr valign=top  height="50"> 
				<td bgcolor="#c1e0a3" align="center" rowspan="1">中午</td>
				<td bgcolor="#E5F6D4" align="center">12:10</td>
				<td bgcolor="#E5F6D4" align="center">13:00</td>
				<%=Str1210%>
				</tr>
				<tr valign=top  height="50"> 
				<td bgcolor="#c1e0a3" align="center" rowspan="5">下午</td>
				<td bgcolor="#E5F6D4" align="center">13:10</td>
				<td bgcolor="#E5F6D4" align="center">14:00</td>
				<%=Str1310%>
				</tr>
				<tr valign=top  height="50"> 
				<td bgcolor="#E5F6D4" align="center">14:10</td>
				<td bgcolor="#E5F6D4" align="center">15:00</td>
				<%=Str1410%>
				</tr>
				<tr valign=top  height="50"> 
				<td bgcolor="#E5F6D4" align="center">15:10</td>
				<td bgcolor="#E5F6D4" align="center">16:00</td>
				<%=Str1510%>
				</tr>
				<tr valign=top  height="50"> 
				<td bgcolor="#E5F6D4" align="center">16:10</td>
				<td bgcolor="#E5F6D4" align="center">17:00</td>
				<%=Str1610%>
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
<!-- #INCLUDE virtual="/include/lib/conn/syconnclose.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->
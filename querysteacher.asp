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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">�d�ߤp�Ѯv�Z�� </TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>
	<TR >
		<TD></TD><TD>
		<%
		sql = "select * from boo_language where yn='Y'"
		rs.Open sql,msconn,adOpenStatic,adLockReadonly
		while not rs.EOF
			response.write "&nbsp;&nbsp;<font color='"&rs("showcolor")&"'>�� </font>&nbsp;" & rs("name")
			rs.MoveNext
		wend


		rs.close

		%>
		
		
		</TD>
	</TR>
	<TR>
		<TD></TD><TD valign="top">
			
		<%
			
			StrDateTop=""'���D������M�P��
			StrStatus0810=""
			StrStatus0910=""
			StrStatus1010=""'10�I
			StrStatus1110=""'11�I
			StrStatus1310=""'13�I
			StrStatus1410=""'14�I
			StrStatus1510=""'15�I
			StrStatus1610=""'16�I
			
			Str0810=""
			Str0910=""
			Str1010=""'10�I
			Str1110=""'11�I
			Str1310=""'13�I
			Str1410=""'14�I
			Str1510=""'15�I
			Str1610=""'16�I
			
			'response.write "BOOK_ROOM=" & BOOK_ROOM
			
			for i=1 to 5
				ww = i
				tmpColor="#FFFFF4"
				
					
				'08�I
				sql = "select a.*,b.showcolor from boo_schedule a left join boo_language b on a.languagecode = b.code where a.category='ST' and a.btime='0810' and a.bweek='"&cint(ww)&"'  and a.yn='Y'  and  yms='"&par_yms&"' "

				'response.write sql & "<br>"
				rs.Open sql,msconn,adOpenStatic,adLockReadonly
				StrStatus0810=""
				while not rs.EOF
					StrStatus0810 = StrStatus0810 & "<font color='"&rs("showcolor")&"'>��</font>"  &  rs("teacher") & "<br>" 
					rs.MoveNext
				wend 
				rs.close
				Str0810 = Str0810 & "<td bgcolor="&tmpColor&">" & StrStatus0810 & "&nbsp;</td>"
				
				'09�I
				sql = "select a.*,b.showcolor from boo_schedule a left join boo_language b on a.languagecode = b.code where a.category='ST' and a.btime='0910' and a.bweek='"&cint(ww)&"'  and a.yn='Y'  and  yms='"&par_yms&"' "

				'response.write sql & "<br>"
				rs.Open sql,msconn,adOpenStatic,adLockReadonly
				StrStatus0910=""
				while not rs.EOF
					StrStatus0910 = StrStatus0910 & "<font color='"&rs("showcolor")&"'>��</font>"  &  rs("teacher") & "<br>" 
					rs.MoveNext
				wend 
				rs.close
				Str0910 = Str0910 & "<td bgcolor="&tmpColor&">" & StrStatus0910 & "&nbsp;</td>"
				'10�I
				sql = "select a.*,b.showcolor from boo_schedule a left join boo_language b on a.languagecode = b.code where a.category='ST' and a.btime='1010' and a.bweek='"&cint(ww)&"'  and a.yn='Y'  and  yms='"&par_yms&"' "

				'response.write sql & "<br>"
				rs.Open sql,msconn,adOpenStatic,adLockReadonly
				StrStatus1010=""
				while not rs.EOF
					StrStatus1010 = StrStatus1010 & "<font color='"&rs("showcolor")&"'>��</font>"  &  rs("teacher") & "<br>" 
					rs.MoveNext
				wend 
				rs.close
				Str1010 = Str1010 & "<td bgcolor="&tmpColor&">" & StrStatus1010 & "&nbsp;</td>"
				'11�I
				sql = "select a.*,b.showcolor from boo_schedule a left join boo_language b on a.languagecode = b.code where a.category='ST' and a.btime='1110' and a.bweek='"&cint(ww)&"'  and a.yn='Y'  and  yms='"&par_yms&"' "

				'response.write sql & "<br>"
				rs.Open sql,msconn,adOpenStatic,adLockReadonly
				StrStatus1110=""
				while not rs.EOF
					StrStatus1110 = StrStatus1110 & "<font color='"&rs("showcolor")&"'>��</font>"  &  rs("teacher") & "<br>" 
					rs.MoveNext
				wend 
				rs.close
				Str1110 = Str1110 & "<td bgcolor="&tmpColor&">" & StrStatus1110 & "&nbsp;</td>"
				'13�I
				sql = "select a.*,b.showcolor from boo_schedule a left join boo_language b on a.languagecode = b.code where a.category='ST' and a.btime='1310' and a.bweek='"&cint(ww)&"'  and a.yn='Y'  and  yms='"&par_yms&"' "
				'response.write sql & "<br>"
				rs.Open sql,msconn,adOpenStatic,adLockReadonly
				StrStatus1310=""
				while not rs.EOF
					StrStatus1310 = StrStatus1310 & "<font color='"&rs("showcolor")&"'>��</font>"  &  rs("teacher") & "<br>" 
					rs.MoveNext
				wend 
				rs.close
				Str1310 = Str1310 & "<td bgcolor="&tmpColor&">" & StrStatus1310 & "&nbsp;</td>"
				
				'14�I
				sql = "select a.*,b.showcolor from boo_schedule a left join boo_language b on a.languagecode = b.code where a.category='ST' and a.btime='1410' and a.bweek='"&cint(ww)&"'  and a.yn='Y'  and  yms='"&par_yms&"' "
				'response.write sql & "<br>"
				rs.Open sql,msconn,adOpenStatic,adLockReadonly
				StrStatus1410=""
				while not rs.EOF
					StrStatus1410 = StrStatus1410 & "<font color='"&rs("showcolor")&"'>��</font>"  &  rs("teacher") & "<br>" 
					rs.MoveNext
				wend 
				rs.close
				Str1410 = Str1410 & "<td bgcolor="&tmpColor&">" & StrStatus1410 & "&nbsp;</td>"
				'15�I
				sql = "select a.*,b.showcolor from boo_schedule a left join boo_language b on a.languagecode = b.code where a.category='ST' and a.btime='1510' and a.bweek='"&cint(ww)&"'  and a.yn='Y'  and  yms='"&par_yms&"' "
				'response.write sql & "<br>"
				rs.Open sql,msconn,adOpenStatic,adLockReadonly
				StrStatus1510=""
				while not rs.EOF
					StrStatus1510 = StrStatus1510 & "<font color='"&rs("showcolor")&"'>��</font>"  &  rs("teacher") & "<br>" 
					rs.MoveNext
				wend 
				rs.close
				Str1510 = Str1510 & "<td bgcolor="&tmpColor&">" & StrStatus1510 & "&nbsp;</td>"
				'16�I
				sql = "select a.*,b.showcolor from boo_schedule a left join boo_language b on a.languagecode = b.code where a.category='ST' and a.btime='1610' and a.bweek='"&cint(ww)&"'  and a.yn='Y'  and  yms='"&par_yms&"' "
				'response.write sql & "<br>"
				rs.Open sql,msconn,adOpenStatic,adLockReadonly
				StrStatus1610=""
				while not rs.EOF
					StrStatus1610 = StrStatus1610 & "<font color='"&rs("showcolor")&"'>��</font>"  &  rs("teacher") & "<br>" 
					rs.MoveNext
				wend 
				rs.close
				Str1610 = Str1610 & "<td bgcolor="&tmpColor&">" & StrStatus1610 & "&nbsp;</td>"

				'17�I
				sql = "select a.*,b.showcolor from boo_schedule a left join boo_language b on a.languagecode = b.code where a.category='ST' and a.btime='1710' and a.bweek='"&cint(ww)&"'  and a.yn='Y'  and  yms='"&par_yms&"' "
				'response.write sql & "<br>"
				rs.Open sql,msconn,adOpenStatic,adLockReadonly
				StrStatus1710=""
				while not rs.EOF
					StrStatus1710 = StrStatus1710 & "<font color='"&rs("showcolor")&"'>��</font>"  &  rs("teacher") & "<br>" 
					rs.MoveNext
				wend 
				rs.close
				Str1710 = Str1710 & "<td bgcolor="&tmpColor&">" & StrStatus1710 & "&nbsp;</td>"
				
			Next
			set rs=nothing
			
		%>
	
			  
				<table border=1 cellpadding=2 cellspacing=2 width="720" bgcolor="#FFFFF4" bordercolor="#326916" align="center">
				<tr valign=top align="center"> 
				<td bgcolor="#c1e0a3" align="center" colspan="3" width="120">�P��<br>�ɬq</td>
				<td bgcolor="#E5F6D4" width="120">Monday<br>�P���@</td><td bgcolor="#E5F6D4" width="120">Tuesday<br>�P���G</td><td bgcolor="#E5F6D4" width="120">Wednesday<br>�P���T</td><td bgcolor="#E5F6D4" width="120">Thursday<br>�P���|</td><td bgcolor="#E5F6D4" width="120">Friday<br>�P����</td>
				</tr>
				<tr valign=top  height="50"> 
				<td bgcolor="#c1e0a3" align="center" rowspan="4">�W��</td>
				<td bgcolor="#E5F6D4" align="center">08:10</td>
				<td bgcolor="#E5F6D4" align="center">09:00</td>
				<%=Str0810%>
				</tr>
				<tr valign=top  height="50"> 
				<td bgcolor="#E5F6D4" align="center">09:10</td>
				<td bgcolor="#E5F6D4" align="center">10:00</td>
				<%=Str0910%>
				</tr>
				<tr valign=top  height="50"> 
				<td bgcolor="#E5F6D4" align="center">10:10</td>
				<td bgcolor="#E5F6D4" align="center">11:00</td>
				<%=Str1010%>
				</tr>
				<tr valign=top  height="50"> 
				<td bgcolor="#E5F6D4" align="center">11:10</td>
				<td bgcolor="#E5F6D4" align="center">12:00</td>
				<%=Str1110%>
				</tr>
				<tr valign=top  > 
				<td bgcolor="#FFFFF4" align="center" colspan="10">Lunch Recess</td>
				
				</tr>
				<tr valign=top  height="50"> 
				<td bgcolor="#c1e0a3" align="center" rowspan="5">�U��</td>
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
				<tr valign=top  height="50"> 
				<td bgcolor="#E5F6D4" align="center">17:10</td>
				<td bgcolor="#E5F6D4" align="center">18:00</td>
				<%=Str1710%>
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
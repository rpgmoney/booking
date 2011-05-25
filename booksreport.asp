<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<!-- #INCLUDE virtual="/include/lib/asp/jsCalendar.asp" -->


<HTML>
<HEAD>
<TITLE> 【LDCC英外語能力診斷輔導中心】 </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/stm31.js" type="text/javascript"></SCRIPT>
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>
<script language="javascript">

function check_input()
{
    var errmsg=""
	
	if (form1.sdate.value=="")
       errmsg += "使用起始日期不能為空白\n";
    if (form1.edate.value=="")
        errmsg += "使用結束日期不能為空白\n";
	
	
    if (errmsg=="")
        return true;
    else
    {
        alert(errmsg);
        return false;
    }
}

function check_all(title,num)
{
	//alert('title=' + title + '\n' + 'num=' + num);
	tmpobj = document.getElementById(title+'_c');

	if (tmpobj .checked==false){
		for (i=0;i<=num;i++){
			tmpobj1 = document.getElementById(title+i);
			tmpobj1.checked=false;
		}
	}
	else
	{
		for (i=0;i<=num;i++){
			tmpobj2 = document.getElementById(title+i);
			tmpobj2.checked=true;
		}
	}
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">自學軟體/補充教材紀錄</TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="booksprint.asp" target="new" onsubmit="return check_input();">
			<input type="hidden" value="add" name="validate">
			<input type="hidden" id="nextrec" name="nextrec">
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>使用日期起迄：</TD>
					</TR>
					<TR>
						<TD>
						<TABLE cellSpacing=0 cellPadding=0 border=0 >
						<TD><input type="text" id="sdate" name="sdate" value="<%=sdate%>"  maxlength="6" class="inputtext"><img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('sdate')" class="showhand"></TD>
						<TD class="inputlabel">&nbsp;&nbsp;~&nbsp;&nbsp;</TD>
						<TD><input type="text" id="edate" name="edate" value="<%=edate%>"  maxlength="6" class="inputtext"><img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('edate')" class="showhand"></TD>
						</table>
						</TD>
						
					</TR>
				</TABLE>
			</TD></TR>
		
			<tr>
					<td class="inputlabel">範圍：</td>
			</tr>
			<tr>
					<td>
						<TABLE cellSpacing=0 cellPadding=0 border=0 >
						<TD><input type="checkbox" name="scope" id="scope0" value="S"></TD><TD>自學軟體{</TD>
						<TD><input type="checkbox" name="scope" id="scope1" value="T"></TD><TD>補充教材</TD>
						
						<TD>﹝</TD>
						<TD><input type="checkbox" name="scope_c" id="scope_c" class="<%=class_normal%>" onclick="check_all('scope','1')" <%if scope_c<>""  then Response.Write "checked" end if%>></TD>
						<TD>全選</TD><TD>﹞</TD>
						</table>
					</td>
			</tr>
			
			<TR>
			<TD>
			<BR>
			<input  type="submit" value="列印" class="inputbutton" >
			</TD>
			</TR>
			</form>
			</TABLE>
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
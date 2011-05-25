<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<!-- #INCLUDE virtual="/include/lib/asp/jsCalendar.asp" -->
<%


'系所
StrDepartment="<option value=''> - 全部 - </option>"
set rsLoad = server.CreateObject("adodb.recordset")
sql ="select * from s90_unit where unt_std='Y' order by unt_sort_seq  "
rsLoad.Open sql,syconn,adOpenStatic,adLockReadonly

while not rsLoad.EOF
	if department=rsLoad("unt_name_abr") then
		StrDepartment=StrDepartment&"<option value="""&rsLoad("unt_name_abr")&""" selected>"&  rsLoad("unt_name_abr")&"</option>"
	else
		StrDepartment=StrDepartment&"<option value="""&rsLoad("unt_name_abr")&""" >"&  rsLoad("unt_name_abr")&"</option>"
	end if 
	rsLoad.MoveNext 
wend
rsLoad.close
'學制
Strslevel="<option value=''> - 全部 - </option>"
sql ="select * from s90_degree "
rsLoad.Open sql,syconn,adOpenStatic,adLockReadonly

while not rsLoad.EOF
	if slevel=rsLoad("dgr_name") then
		Strslevel=Strslevel&"<option value="""&rsLoad("dgr_name")&""" selected>"&  rsLoad("dgr_name")&"</option>"
	else
		Strslevel=Strslevel&"<option value="""&rsLoad("dgr_name")&""" >"&  rsLoad("dgr_name")&"</option>"
	end if 
	rsLoad.MoveNext 
wend
rsLoad.close

set rsLoad=nothing



%>
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
	
	//if (form1.topic.value=="")
  //      errmsg += "口語題目不能為空白\n";
  //  if (form1.category.value=="")
  //      errmsg += "類別不能為空白\n";
//	 if (form1.rank.value=="")
 //       errmsg += "等級不能為空白\n";
//    if (form1.yn.value=="")
//        errmsg += "開放否不能為空白\n";
	
	
	
    if (errmsg=="")
        return true;
    else
    {
        alert(errmsg);
        return false;
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">學習問卷調查 - 學習風格明細表</TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="stylereportprint.asp" target="new" onsubmit="return check_input();">
			<input type="hidden" value="add" name="validate">
			<input type="hidden" id="nextrec" name="nextrec">
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>註冊日期起迄：</TD>
					</TR>
					<TR>
						<TD>
						<TABLE cellSpacing=0 cellPadding=0 border=0 >
						<TD><input type="text" id="sRegDate" name="sRegDate" value="<%=sRegDate%>"  maxlength="6" class="inputtext"><img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('sRegDate')" class="showhand"></TD>
						<TD class="inputlabel">&nbsp;&nbsp;~&nbsp;&nbsp;</TD>
						<TD><input type="text" id="eRegDate" name="eRegDate" value="<%=eRegDate%>"  maxlength="6" class="inputtext"><img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('eRegDate')" class="showhand"></TD>
						</table>
						</TD>
						
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>填寫問卷日期起迄：</TD>
					</TR>
					<TR>
						<TD>
						<TABLE cellSpacing=0 cellPadding=0 border=0 >
						<TD><input type="text" id="sIniDate" name="sIniDate" value="<%=sIniDate%>"  maxlength="6" class="inputtext"><img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('sIniDate')" class="showhand"></TD>
						<TD class="inputlabel">&nbsp;&nbsp;~&nbsp;&nbsp;</TD>
						<TD><input type="text" id="eIniDate" name="eIniDate" value="<%=eIniDate%>"  maxlength="6" class="inputtext"><img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('eIniDate')" class="showhand"></TD>
						</table>
						</TD>
						
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>學制：</TD>
						<TD>系所：</TD>
						<TD>就學中：</TD>
					</TR>
					<TR>
						<TD>
							<select name="slevel" class="inputtext">
							<%=Strslevel%>
							</select>
						</TD>
						<TD>
							<select name="department" class="inputtext">
							<%=StrDepartment%>
							</select>
						</TD>
						<TD>
						<select name="yn" class="inputtext">
						<option value=""> - 全部 -</option>
						<option value="Y" <%if yn="Y" then response.write "selected" end if%>>Y</option>
						<option value="N" <%if yn="N" then response.write "selected" end if%>>N</option>
						</select>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			
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
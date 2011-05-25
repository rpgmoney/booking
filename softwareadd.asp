<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<%
validate=trim(request("validate"))
nextrec=trim(request("nextrec"))
floor=trim(request("floor")) 
software=trim(request("software"))
yn=trim(request("yn"))
category=trim(request("category")) 




sender=ifnull(trim(request("sender")),"software.asp")

set rs = server.CreateObject("adodb.recordset")
if validate="add" then
	sql = "select * from boo_software where floor='"&floor&"' and software='"&software&"' "
	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
    if rs.EOF then
        rs.AddNew
		if floor<>"" then
            rs("floor")=floor
        end if
		if software<>"" then
            rs("software")=software
        end if
		if category<>"" then
            rs("category")=category
        end if
		if yn<>"" then
            rs("yn")=yn
        end if
		rs("initdate") = date()
		rs("inituid") = session("sid")


		rs.Update
        if Err.Number=0 then 
			if nextrec="Y" then
				validate=""
				nextrec=""
				floor=""
				software=""
				
			else
				response.redirect sender
			end if
		else
			showmessage= Err.Description
		end if

	else
		showmessage="資料重覆。"
	end if

	rs.close
	
end if


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
	
	if (form1.floor.value=="")
        errmsg += "層樓不能為空白\n";
    if (form1.software.value=="")
        errmsg += "軟體名稱不能為空白\n";
	 if (form1.yn.value=="")
        errmsg += "開放否不能為空白\n";

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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">新增自學療程軟體項目</TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="softwareadd.asp" onsubmit="return check_input();">
			<input type="hidden" value="add" name="validate">
			<input type="hidden" id="nextrec" name="nextrec">
			<input type="hidden" id="category" name="category" value="S">
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>層樓：</TD>
						<TD>軟體名稱：</TD>
					</TR>
					<TR>
						<TD>
						<select name="floor" class="inputtext">
						<option value=""> - 請指定 -</option>
						<option value="2F" <%if floor="2F" then response.write "selected" end if%>>2F</option>
						<option value="3F" <%if floor="3F" then response.write "selected" end if%>>3F</option>
						</select>

						</TD>
						<TD>
						<input type="text" value="<%=software%>" maxlength="50" size="55" name="software" class="inputtext" >
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						
						<TD>開放否：</TD>
					</TR>
					<TR>
						<TD>
						<select name="yn" class="inputtext">
						<option value=""> - 請指定 -</option>
						<option value="Y" <%if yn="Y" then response.write "selected" end if%>>開放</option>
						<option value="N" <%if yn="N" then response.write "selected" end if%>>關閉</option>
						</select>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			
			<TR>
			<TD>
			<BR>
			<input  type="submit" value="新增" class="inputbutton" >
			<input  type="submit" onclick="form1.nextrec.value='Y'" value="新增後繼續新增" class="inputbutton">
			<input  type="button" value="返回" class="inputbutton" onclick="window.location='<%=sender%>'">
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
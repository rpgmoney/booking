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
id=trim(request("id"))
item=trim(request("item"))
hours=trim(request("hours"))
yn=trim(request("yn"))



sender=ifnull(trim(request("sender")),"selflist.asp")

set rs = server.CreateObject("adodb.recordset")
if validate="edit" then
	sql = "select * from boo_self_item where id = '"&id&"' "
'	response.end
	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
    if not rs.EOF then
		if item<>"" then
            rs("item")=item
        end if
		if hours<>"" then
            rs("hours")=hours
        end if
		if yn<>"" then
            rs("yn")=yn
        end if
		rs("initdate") = date()
		rs("inituid") = session("sid")


		rs.Update
        if Err.Number=0 then 
			response.redirect "selflist.asp"
		else
			showmessage= Err.Description
		end if

	else
		showmessage="�䤣��ӵ���ơC"
	end if

	rs.close
else
	sql = "select * from boo_self_item where id='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
	if rs.EOF then
		response.redirect sender
	else
		id=trim(rs("id"))
		item=trim(rs("item"))
		hours=trim(rs("hours"))
		yn=trim(rs("yn"))


	end if

end if


%>
<HTML>
<HEAD>
<TITLE> �iLDCC�^�~�y��O�E�_���ɤ��ߡj </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/stm31.js" type="text/javascript"></SCRIPT>
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>
<script language="javascript">

function check_input()
{
    var errmsg=""
	
	if (form1.item.value=="")
        errmsg += "���ؤ��ର�ť�\n";
    if (form1.hours.value=="")
        errmsg += "���w�ɼƤ��ର�ť�\n";
    if (form1.yn.value=="")
        errmsg += "�}��_���ର�ť�\n";
	
	
	
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">�s��۾����{����</TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="selfedit.asp" onsubmit="return check_input();">
			<input type="hidden" value="edit" name="validate">
			<input type="hidden" id="id" name="id" value="<%=id%>">
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>���ءG</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=item%>" maxlength="100" size="55" name="item" class="inputtext" >
						</TD>
						
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>�@�ѭ��w�ɼơG</TD>
						<TD>�}��_�G</TD>
					</TR>
					<TR>
						<TD>
						<select name="hours" class="inputtext">
						<option value=""> - �Ы��w -</option>
						<option value="999" <%if hours="999" then response.write "selected" end if%>>����</option>
						<%=NumOption2(1,8,hours)%>
						</select>
						</TD>
						
						<TD>
						<select name="yn" class="inputtext">
						<option value=""> - �Ы��w -</option>
						<option value="Y" <%if yn="Y" then response.write "selected" end if%>>�}��</option>
						<option value="N" <%if yn="N" then response.write "selected" end if%>>����</option>
						</select>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			
			<TR>
			<TD>
			<BR>
			<input  type="submit" value="�x�s" class="inputbutton" >
			<input  type="button" value="��^" class="inputbutton" onclick="window.location='<%=sender%>'">
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
<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<!-- #INCLUDE virtual="/include/lib/asp/jsCalendar.asp" -->
<%
validate=trim(request("validate"))
id=trim(request("id"))
tactile=trim(request("tactile")) 
individual=trim(request("individual")) 
visual=trim(request("visual")) 
auditory=trim(request("auditory")) 
kinesthetic=trim(request("kinesthetic")) 
group1=trim(request("group1")) 


set rs = server.CreateObject("adodb.recordset")
if validate="edit" then
	sql = "select * from boo_style_analysis where id='1'  "
	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
    if  not rs.EOF then
		if tactile<>"" then
            rs("tactile")=tactile
        else
			rs("tactile")=null
		end if
		if individual<>"" then
            rs("individual")=individual
        else
			rs("individual")=null
		end if
		if visual<>"" then
            rs("visual")=visual
        else
			rs("visual")=null
		end if
		if auditory<>"" then
            rs("auditory")=auditory
        else
			rs("auditory")=null
		end if
		if kinesthetic<>"" then
            rs("kinesthetic")=kinesthetic
        else
			rs("kinesthetic")=null
		end if
		if group1<>"" then
            rs("group1")=group1
        else
			rs("group1")=null
		end if
		
		rs("initdate") = date()
		rs("inituid") = session("sid")

		rs.Update
        if Err.Number<>0 then 
			showmessage= Err.Description
		end if

	else
		showmessage="�䤣��ӵ���ơC"
	end if

	rs.close
else

	sql = "select * from boo_style_analysis where id='1' "
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
	if rs.EOF then
		response.redirect sender
	else
		id=trim(rs("id"))
		tactile=trim(rs("tactile")) 
		individual=trim(rs("individual")) 
		visual=trim(rs("visual"))
		auditory=trim(rs("auditory"))
		kinesthetic=trim(rs("kinesthetic"))
		group1=trim(rs("group1"))
	end if
	rs.close
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
	
//    if (form1.subject.value=="")
//        errmsg += "�W�٤��ର�ť�\n";
//	if (form1.place.value=="")
//        errmsg += "�a�I���ର�ť�\n";
//	if (form1.speaker.value=="")
//        errmsg += "�D���H/�Ѯv���ର�ť�\n";
//	if (form1.content.value=="")
 //       errmsg += "���y���e/�ҵ{���e���ର�ť�\n";
//	if (form1.sdate.value=="")
 //       errmsg += "�}�l���W������ର�ť�\n";
//		if (form1.edate.value=="")
//        errmsg += "�������W������ର�ť�\n";

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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">�����R�w�q</TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="styleanalysis.asp" onsubmit="return check_input();">
			<input type="hidden" value="edit" name="validate">
			
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>Ĳı�ǲߪ�tactile�G</TD>
					</TR>
					<TR>
						<TD>
						<textarea name="tactile" cols="80" rows="6" class="inputtext" ><%=tactile%></textarea>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>�ӤH�ǲߪ�individual�G</TD>
					</TR>
					<TR>
						<TD>
						<textarea name="individual" cols="80" rows="6" class="inputtext" ><%=individual%></textarea>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>��ı�ǲߪ�visual�G</TD>
					</TR>
					<TR>
						<TD>
						<textarea name="visual" cols="80" rows="6" class="inputtext" ><%=visual%></textarea>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>ťı�ǲߪ�auditory�G</TD>
					</TR>
					<TR>
						<TD>
						<textarea name="auditory" cols="80" rows="6" class="inputtext" ><%=auditory%></textarea>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>��ı�ǲߪ�kinesthetic�G</TD>
					</TR>
					<TR>
						<TD>
						<textarea name="kinesthetic" cols="80" rows="6" class="inputtext" ><%=kinesthetic%></textarea>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>����ǲߪ�group�G</TD>
					</TR>
					<TR>
						<TD>
						<textarea name="group1" cols="80" rows="6" class="inputtext" ><%=group1%></textarea>
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
<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<%

today=year(date())-1911 & right("0" & month(date()),2) & right("0" & day(date()),2)
bdate=today

id = trim(request("id"))
validate=trim(request("validate"))
deptname=trim(request("deptname"))
sid=trim(request("sid"))
name=trim(request("name"))
classify=trim(request("classify"))
enable=trim(request("enable"))

sender=ifnull(trim(request("sender")),"auth.asp")


set rs = server.CreateObject("adodb.recordset")
if validate="edit" then
	sql = "select * from boo_profile where id='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
    if not rs.EOF then
		if sid<>"" then
            rs("sid")=sid
        end if
		if name<>"" then
            rs("name")=name
        end if
		if deptname<>"" then
            rs("department")=deptname
        end if
		if classify<>"" then
            rs("classify")=classify
        else
			rs("classify")=null
		end if
		
		if enable<>"" then
            rs("enable")=enable
        end if

		rs.Update
        if Err.Number=0 then 

				response.redirect sender  
        else
            showmessage= Err.Description
        end if

	else
		showmessage="�䤣��ӵ���ơC"
	end if
	rs.Close
else
	sql = "select * from boo_profile where id='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
	deptname=trim(rs("department"))
	sid=trim(rs("sid"))
	name=trim(rs("name"))
	classify=trim(rs("classify"))
	enable=trim(rs("enable"))

	rs.Close
end if


%>
<HTML>
<HEAD>
<TITLE> �iLDCC�^�~�y��O�E�_���ɤ��ߡj  </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/stm31.js" type="text/javascript"></SCRIPT>
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>
<script language="javascript">

function check_input()
{
    var errmsg=""
	
    if (form1.deptname.value=="")
        errmsg += "�t�Ҥ��ର�ť�\n";
	 if (form1.sid.value=="")
        errmsg += "�Ǹ����ର�ť�\n";

    if (errmsg=="")
        return true;
    else
    {
        alert(errmsg);
        return false;
    }
}

function clickblock(id)
{
	obj=document.getElementById("remark");
	if (obj!=null)
	{
		if (id==1)
		{
			obj.style.display="block";
			//obj1.src="images/icon-rectup.gif"
		}
		else
		{
			obj.style.display="none";
			//obj1.src="images/icon-rectdown.gif"
		}
	}
}

function ChkSid()
{
	vWinCal2 = window.open("lib/checksid.asp?sid="+form1.sid.value  ,"iframe_query", "width=200,height=50,status=no,resizable=no,scrollbars=no");
	vWinCal2.opener = form1;
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">�s��Ѯv/�p�Ѯv�W��</TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=2 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="authedit.asp" onsubmit="return check_input();">
			<input type="hidden" value="edit" name="validate">
			<input type="hidden" value="<%=id%>" name="id">
			<input type="hidden" value="<%=sender%>" name="sender">
			
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>�Ǹ��G</TD>
						<TD>�m�W�G</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=sid%>" maxlength="25" size="35" onblur="ChkSid()" name="sid" class="inputtext"  >
						</TD>
						<TD>
						<input type="text" value="<%=name%>" maxlength="25" size="35" name="name" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly >
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>���</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=deptname%>" maxlength="25" size="65"  name="deptname" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly  >
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD >�v��</TD>
					</TR>
					<TR >
						<TD>
						<select name="classify" class="inputtext">
						<option value="T">�Ѯv</option>
						</select>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD >�O�_�}��</TD>
					</TR>
					<TR >
						<TD>
						<select name="enable" class="inputtext">
						<option value="Y">�}��</option>
						<option value="N">����</option>
						</select>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR>
			<TD>
			<BR>
			<input  type="submit" value="�x�s" class="inputbutton" >
			<!-- <input  type="button" value="�R��" class="inputbutton" > -->
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
<iframe style="display:none"  name="iframe_query" id="iframe_query"></iframe>
</BODY>
</HTML>
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconnclose.asp" -->
<%@Language=VBScript LCID=1033%>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<%
 
'if  Request.ServerVariables("URL") ="/ldcc/login.asp" then
'	response.redirect "http://ldcc.wtuc.edu.tw/login.asp"
'end if
'response.write Request.ServerVariables("HTTP_HOST") 
validate=trim(request("validate"))
loginID=trim(request("loginID"))
pwd=trim(request("pwd"))
if validate="CheckAccount" then
	on error resume next
	set rs = server.CreateObject("adodb.recordset")
	set rs1 = server.CreateObject("adodb.recordset")
	session("questionnaire")=""

	'��check�b���K�X sts_id ��  ���O   '02'-->�ǥ�     01-->¾��
	sql = 	" select usr_pwd,emp_idno as idno,emp_id as sid,emp_name as name,emp_email as email,'01' as status  from s10_employee a ,s90_auth_usr b "
	sql = sql &		" where (emp_leavedate is null or emp_leavedate='' ) "
	sql = sql &		" and ( emp_id='"&loginID&"' or emp_name = '"&loginID&"')   and a.emp_idno = b.usr_id "
	'sql = sql &		" and b.usr_pwd='"&pwd&"' "
	sql = sql &		" union "
	sql = sql &		" select usr_pwd,s30_student.std_id as idno ,s30_student.std_id as sid, s30_student.std_name as name,s30_student.std_email  as email,'02' as status from s30_student , s30_sturec , s90_yms ,s90_auth_usr"
	sql = sql &		" where s90_yms.yms_mark = 'Y' and  s90_yms.yms_year = s30_sturec.yms_year "
	sql = sql &		" and  s90_yms.yms_smester = s30_sturec.yms_sms and s30_student.std_key = s30_sturec.std_key  "
	sql = sql &		" and (s30_student.std_id='"&loginID&"' or s30_student.std_name ='"&loginID&"') and s30_student.std_id = s90_auth_usr.usr_id  "
	'sql = sql &		"and s90_auth_usr.usr_pwd='"&pwd&"' " 
	'response.write sql
	'Set rs = syconn.Execute(sql)
	rs.Open sql,syconn,adOpenStatic,adLockReadonly
	if  Err.Number<>0 then
		'response.write  Err.Description
		rs.close
		rs.Open sql,syconn,adOpenStatic,adLockReadonly
		if  Err.Number<>0 then
			'response.write  Err.Description
			rs.close
			rs.Open sql,syconn,adOpenStatic,adLockReadonly
		end if
	end if
	rs.Open sql,syconn,adOpenStatic,adLockReadonly

	if not rs.EOF then
		tmppassword = trim(rs("usr_pwd"))

		
	'	if tmppassword = pwd then

			session("st_status")=trim(rs("status"))'�ǥͩ�¾�� 
			session("sid")=trim(rs("sid"))
			session("sname")=trim(rs("name"))

			rflag="Y" '���U�L�_
			sql = "select * from boo_profile where enable='Y' and sid='"&loginID&"' "
	'response.write sql

	'response.end
			rs1.Open sql,msconn,adOpenStatic,adLockReadonly
			if not rs1.EOF then
				rflag="N" 
				'if session("sid")= "" or isnull(session("sid")) or isempty(session("sid")) then 
					session("sid")=loginID
				'end if
				'if session("sname")= "" or isnull(session("sname")) or isempty(session("sname")) then 
					session("sname")= trim(rs1("name"))
				'end if
				session("classify")=trim(rs1("classify")) '�v��
				session("dept")=trim(rs1("department"))
				sytle_yn = trim(rs1("sytle_yn")) '�ݨ����ħ_
				strategy_yn = trim(rs1("strategy_yn"))
				if session("classify")="S" then
					'�ǥ�
					if sytle_yn="Y"  and strategy_yn="Y" then
						'�i�����w��
						'response.redirect ""
						session("questionnaire")="Y"
						response.redirect "queryteacher.asp"
						response.end
					else
						'�ݧ����ݨ�
						response.redirect "studentedit.asp?sid=" & loginID
						showmessage="�ݧ����ݨ�"
						response.end
					end if
				elseif session("classify")="A" then
					'�޲z��	
					response.redirect "queryteacher.asp"
					showmessage="�޲z��"
					response.end
				else
					response.redirect "queryteacher.asp"
					response.end
				end if
			end if
			rs1.close
			set rs1 = nothing
			if session("st_status")="02" then
				if rflag="Y" then '
					'�O�ǥͥB�����U��
					session("classify") ="S"
					response.redirect "register.asp?sid="&loginID
					response.end
				end if
			else
				'�@��H���i�H�s��
				'session("classify") ="E"
'showmessage = session("sid") &  session("st_status")		
			'	response.redirect "queryteacher.asp"
				'showmessage="�D�դ��ǥ͵L�k�ϥΡA�Y��������D�Ь��~�E���ߡC"
			end if
	'	else
	'		showmessage="�L�k�n�J�A�ЦA�@���T�{�z���K�X�L�~"
	'	end if
	else
		showmessage="�L�k�n�J�A�ЦA�@���T�{�z���b���L�~"
	end if 
	rs.close
	set rs = nothing
end if
%>
<HTML>
<HEAD>
<TITLE> �iLDCC�^�~�y��O�E�_���ɤ��ߡj</TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="include/js/lib.js" type="text/javascript"></SCRIPT>
<script language="javascript">
function clickblock(id)
{
	obj=document.getElementById(id);
	if (obj!=null)
	{
		if (obj.style.display=="none")
		{
			obj.style.display="block";
		}
		else
		{
			obj.style.display="none";
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
<TR bgcolor="#333333" height="30">
	<TD></TD>
</TR>
<TR>
	<TD align="center"><font color="red"><%=showmessage%></font></TD>
</TR>
<TR>
	<TD>
<!-- ---------------------------------------------------------------------------------------- -->

		<form name="form1" method=post action="login.asp">
		<input type="hidden" value="CheckAccount" name="validate">
		<table border=1 cellpadding=0 cellspacing=2 width="50%" bgcolor="#FFFFF4" bordercolor="#f4c60d" align="center">
		
		<tr valign=top> 
			  <td height=17 width="21%" bgcolor="#f4c60d" align="center"> 
				 <font class="T1">���п�J�z���ӤH�{�ұb���εn�J�K�X�Ӷi�J�t��</font>
			   </td>                         
		  </tr>
		  <tr valign=top> 
			  <td bgcolor="#FFFFF4"> 
				 <TABLE border=0 cellpadding=0 cellspacing=2 height="150" align="center" width="70%" >
				 <TR>
				 	<TD align="right">�b���G</TD>
				 	<TD><input type=text maxlength="10" size=20 name="loginID" class="inputtext"  value="<%=loginID%>"></TD>
				 </TR>
				 <TR>
				 	<TD align="right">�K�X�G</TD>
				 	<TD><input type=password maxlength="20" size=20 name="pwd" class="inputtext" value="<%=pwd%>"></TD>
				 </TR>
				 <TR>
				 	<TD colspan="2" align="center"><input type="submit" class="inputbutton" value="�T�w"><input class="inputbutton" type="reset" value="�M��"></TD>
				 </TR>
				 </TABLE>
			  </td>                         
		  </tr>
		</table>
		</form>
		<TABLE cellSpacing=1 cellPadding=0  width="50%" border=1 bgColor=#FFFFF4 align="center" bordercolor="#f4c60d">
		<TR><TD onclick="javascript:clickblock('remark');" class="showhand"><font color="#CC3300">�n�J�K�X�ϥλ���</font>&nbsp;&nbsp;��</TD></TR>
		<TR style="DISPLAY:block" id=remark>
		<TD><TABLE cellSpacing=0 cellPadding=0  border=0   class="normal" >
			<TR>
			<TD>
			<P>�ǥ͡G<BR>
			�n�J�b�����G�Ǹ�<BR>
			�K�X���G�P�հȨt�άۦP<BR>
			�w�]�K�X���G�����ҫ�|�X<BR>

			<P>�޲z�̡G<BR>
			�n�J�b�����G�����Ҧr��<BR>
			�K�X���G�P�հȨt�άۦP<BR>
			�w�]�K�X���G�����ҫ�|�X<BR>
			</TD>
			</TR>
			</TABLE>
		</TD></TR>
		</TABLE>
<!-- ---------------------------------------------------------------------------------------- -->
	</TD>
</TR>
<TR>
	<TD>
		
		<BR><P><BR>
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
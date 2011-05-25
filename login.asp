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

	'先check帳號密碼 sts_id 為  類別   '02'-->學生     01-->職員
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

			session("st_status")=trim(rs("status"))'學生或職員 
			session("sid")=trim(rs("sid"))
			session("sname")=trim(rs("name"))

			rflag="Y" '註冊過否
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
				session("classify")=trim(rs1("classify")) '權限
				session("dept")=trim(rs1("department"))
				sytle_yn = trim(rs1("sytle_yn")) '問卷有效否
				strategy_yn = trim(rs1("strategy_yn"))
				if session("classify")="S" then
					'學生
					if sytle_yn="Y"  and strategy_yn="Y" then
						'可直接預約
						'response.redirect ""
						session("questionnaire")="Y"
						response.redirect "queryteacher.asp"
						response.end
					else
						'需完成問卷
						response.redirect "studentedit.asp?sid=" & loginID
						showmessage="需完成問卷"
						response.end
					end if
				elseif session("classify")="A" then
					'管理者	
					response.redirect "queryteacher.asp"
					showmessage="管理者"
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
					'是學生且未註冊者
					session("classify") ="S"
					response.redirect "register.asp?sid="&loginID
					response.end
				end if
			else
				'一般人員可以瀏覽
				'session("classify") ="E"
'showmessage = session("sid") &  session("st_status")		
			'	response.redirect "queryteacher.asp"
				'showmessage="非校內學生無法使用，若有任何問題請洽外診中心。"
			end if
	'	else
	'		showmessage="無法登入，請再一次確認您的密碼無誤"
	'	end if
	else
		showmessage="無法登入，請再一次確認您的帳號無誤"
	end if 
	rs.close
	set rs = nothing
end if
%>
<HTML>
<HEAD>
<TITLE> 【LDCC英外語能力診斷輔導中心】</TITLE>
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
				 <font class="T1">◎請輸入您的個人認證帳號及登入密碼來進入系統</font>
			   </td>                         
		  </tr>
		  <tr valign=top> 
			  <td bgcolor="#FFFFF4"> 
				 <TABLE border=0 cellpadding=0 cellspacing=2 height="150" align="center" width="70%" >
				 <TR>
				 	<TD align="right">帳號：</TD>
				 	<TD><input type=text maxlength="10" size=20 name="loginID" class="inputtext"  value="<%=loginID%>"></TD>
				 </TR>
				 <TR>
				 	<TD align="right">密碼：</TD>
				 	<TD><input type=password maxlength="20" size=20 name="pwd" class="inputtext" value="<%=pwd%>"></TD>
				 </TR>
				 <TR>
				 	<TD colspan="2" align="center"><input type="submit" class="inputbutton" value="確定"><input class="inputbutton" type="reset" value="清除"></TD>
				 </TR>
				 </TABLE>
			  </td>                         
		  </tr>
		</table>
		</form>
		<TABLE cellSpacing=1 cellPadding=0  width="50%" border=1 bgColor=#FFFFF4 align="center" bordercolor="#f4c60d">
		<TR><TD onclick="javascript:clickblock('remark');" class="showhand"><font color="#CC3300">登入密碼使用說明</font>&nbsp;&nbsp;▼</TD></TR>
		<TR style="DISPLAY:block" id=remark>
		<TD><TABLE cellSpacing=0 cellPadding=0  border=0   class="normal" >
			<TR>
			<TD>
			<P>學生：<BR>
			登入帳號為：學號<BR>
			密碼為：與校務系統相同<BR>
			預設密碼為：身份證後四碼<BR>

			<P>管理者：<BR>
			登入帳號為：身份證字號<BR>
			密碼為：與校務系統相同<BR>
			預設密碼為：身份證後四碼<BR>
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
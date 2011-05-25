<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<%Server.ScriptTimeOut =600000%>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<%
validate=trim(request("validate"))
showhint=trim(request("showhint"))
score=trim(request("score"))
priority=trim(request("priority"))


sender=ifnull(trim(request("sender")),"orallevellist.asp")

set rs = server.CreateObject("adodb.recordset")
set rs2 = server.CreateObject("adodb.recordset")
if validate="edit" then
	sql = "select * from boo_parameter where id = 'A' "
'	response.end
	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
    if not rs.EOF then
		if showhint<>"" then
            rs("showhint")=showhint
        end if
		if score<>"" then
            rs("score")=score
        end if
		if priority<>"" then
            rs("priority")=priority
        end if
		
		rs("modifydate") = date()
		rs("modifyuid") = session("sid")


		rs.Update
        if Err.Number<>0 then 
		
			showmessage= Err.Description
		end if

	else
		showmessage="找不到該筆資料。"
	end if

	rs.close
elseif  validate="update_date" then
	Set UpdateList = Server.CreateObject("kmuhcom.Simplelist")
	sqlM="select sid,name,department ,grade,slevel,class1  from boo_profile where  enable='Y' and classify='S'  "
	rs.Open sqlM,msconn,adOpenStatic,adLockReadonly
	while not rs.EOF

		sql = "select s30_student.std_id, s30_student.std_name,s30_student.std_birthday ,s30_student.std_sex,s30_student.std_idno "
		sql = sql & " ,s30_student.std_tel,s30_student.std_email,s90_class.cls_id_abr , s90_class.cls_year  , "
		sql = sql & " s90_degree.dgr_name , s90_division.dvs_name , s90_unit.unt_name_abr,s90_class.cls_id "
		sql = sql & " from s30_student , s30_sturec , s90_yms  , s90_class , s90_degree , s90_division , s90_unit "
		sql = sql & " where s90_yms.yms_mark = 'Y' and "
		'sql = sql & " where s90_yms.yms_year = 99 and s90_yms.yms_smester=1 and "
		sql = sql & "  s90_yms.yms_year = s30_sturec.yms_year and  "
		sql = sql & "  s90_yms.yms_smester = s30_sturec.yms_sms and  "
		sql = sql & "  s30_student.std_key = s30_sturec.std_key and  "
		sql = sql & "  s30_sturec.cls_id = s90_class.cls_id  and  "
		sql = sql & "  s90_class.dgr_id = s90_degree.dgr_id and  "
		sql = sql & "  s90_degree.dvs_id = s90_division.dvs_id and  "
		sql = sql & "  s90_class.unt_id = s90_unit.unt_id and "
		sql = sql & " s30_student.std_id = '"&rs("sid")&"'   " 
		'response.write sql & "<br>"
		'response.end
		rs2.Open sql,syconn,adOpenStatic,adLockReadonly
		if not rs2.EOF then
			name = trim(rs2("std_name"))
			department = trim(rs2("unt_name_abr"))
			grade = trim(rs2("cls_year"))
			slevel = trim(rs2("dgr_name"))
			class1 = replace(replace(replace(replace(right(trim(rs2("cls_id")),1),"1","A"),"2","B"),"3","C"),"4","D")

			strupdate = ""
'			if name<>trim(rs("name")) then
'				if strupdate <>"" then strupdate = strupdate & "," end if
'				strupdate = strupdate & "name='"&name&"'"
'			end if
			if department<>trim(rs("department")) then
				if strupdate <>"" then strupdate = strupdate & "," end if
				strupdate = strupdate & "department='"&department&"'"
			end if
			if grade<>trim(rs("grade")) then
				if strupdate <>"" then strupdate = strupdate & "," end if
				strupdate = strupdate & "grade='"&grade&"'"
			end if
			if slevel<>trim(rs("slevel")) then
				if strupdate <>"" then strupdate = strupdate & "," end if
				strupdate = strupdate & "slevel='"&slevel&"'"
			end if
			if class1<>trim(rs("class1"))  then
				if strupdate <>"" then strupdate = strupdate & "," end if
				strupdate = strupdate & "class1='"&class1&"'"
			end if
			if strupdate<>"" then
			sql_update="update boo_profile set  "&strupdate&" where sid='"&rs("sid")&"'"
			'response.write sql_update & "<br>"
			UpdateList.add sql_update
			end if

		else
			'response.write rs("sid") & "非合法學員。<br>"
			sql_update="update boo_profile set  enable='N' where sid='"&rs("sid")&"'"
			'response.write sql_update & "<br>"
			UpdateList.add sql_update
		end if
		rs2.close


		rs.MoveNext
	wend
	rs.close
	if UpdateList.count>0 then
			msconn.BeginTrans
			for each item in UpdateList
				Response.Write item & "<br>"
				msconn.Execute item
			next
			if Err.Number<>0 then
				msconn.RollbackTrans
				showmessage= "失敗：" &  Err.Description
			else
				msconn.CommitTrans
				validate=""
				showmessage="更新成功"
			 end if
		end if
else
	sql = "select * from boo_parameter where id='A' "
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
	if rs.EOF then
		response.redirect sender
	else
		showhint=trim(rs("showhint"))
		score=trim(rs("score"))
		priority=trim(rs("priority"))
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
	
	if (form1.score.value=="")
        errmsg += "英檢成績高於幾分後則不須依處方籤不能為空白\n";
    if (form1.showhint.value=="")
        errmsg += "請指定是否須秀提醒視窗\n";
	
	if (errmsg == "")
	{
		if (form1.score.value!="" && !isint(form1.score.value))
			errmsg += "『英檢成績高於幾分後則不須依處方籤』欄位請填數字\n";
	}
    if (errmsg=="")
        return true;
    else
    {
        alert(errmsg);
        return false;
    }
}

function update_student_data()
{
	if (confirm("確定要學員的學制、年級、班別、系所名稱資料要與教務系統同步嗎？")){
		form1.validate.value="update_date";
		form1.submit();

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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">系統參數設定</TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="parameter.asp" onsubmit="return check_input();">
			<input type="hidden" value="edit" name="validate">
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>英檢成績高於幾分後則不須依處方籤：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=score%>" maxlength="3" size="30" name="score" class="inputtext" >
						</TD>
						
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>秀提醒視窗（預約項目必須依處方籤）：</TD>
						
					</TR>
					<TR>
						
						<TD>
						<select name="showhint" class="inputtext" style="width:150">
						<option value=""> - 請指定 -</option>
						<option value="Y" <%if showhint="Y" then response.write "selected" end if%>>是</option>
						<option value="N" <%if showhint="N" then response.write "selected" end if%>>否</option>
						</select>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>學員資料與校務系統同步：(學期開始前更新)</TD>
					</TR>
					<TR>
						<TD>
						<input  type="button" value="學員資料與校務系統同步"  onclick="update_student_data();"   <%if session("sid")<>"95186" then response.write "disabled" end if%> class="inputbutton" >
						</TD>
						
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>口語練習：</TD>
						
					</TR>
					<TR>
						
						<TD>
						<TABLE cellSpacing=0 cellPadding=0  border=0 >
							<TR>
							<TD><input type="radio" name="priority"  class="inputtext" value="1" <%if priority="1" then response.write "checked" end if%>></TD><TD>以駐診老師為第一順位</TD>
							<TD><input type="radio" name="priority"  class="inputtext" value="2" <%if priority="2" then response.write "checked" end if%>></TD><TD>不受限於是否有駐診老師</TD>
							</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR>
			<TD>
			<BR>
			<input  type="submit" value="更新系統參數" class="inputbutton" >&nbsp;
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
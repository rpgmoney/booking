<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<!-- #INCLUDE virtual="/include/lib/asp/jsCalendar.asp" -->
<%
validate=trim(request("validate"))
sid=trim(request("sid"))
flag = trim(request("flag"))


name=trim(request("name"))
birthday=trim(request("birthday"))
sex=trim(request("sex"))
mail=trim(request("mail"))
cell=trim(request("cell"))
slevel=trim(request("slevel"))
grade=trim(request("grade"))
class1=trim(request("class1"))
department=trim(request("department"))
purpose=trim(request("purpose"))
purpose_remark=trim(request("purpose_remark"))
howknow=trim(request("howknow"))
howknow_remark=trim(request("howknow_remark"))
note=trim(request("note"))
btncontrol=trim(request("btncontrol"))
sender=ifnull(trim(request("sender")),"studentlist.asp" )

set rs = server.CreateObject("adodb.recordset")
if validate="add" then
	sql = "select * from boo_profile where sid='"&sid&"'"
	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
    if rs.EOF then
        rs.AddNew
		if sid<>"" then
            rs("sid")=sid
        end if
		if name<>"" then
            rs("name")=name
        end if
		if birthday<>"" then
            rs("birthday")=birthday
        end if
		if sex<>"" then
            rs("sex")=sex
        end if
		if mail<>"" then
            rs("mail")=mail
        end if
		if cell<>"" then
            rs("cell")=cell
        end if
		if slevel<>"" then
            rs("slevel")=slevel
        end if
		if grade<>"" then
            rs("grade")=grade
        end if
		if class1<>"" then
            rs("class1")=class1
        end if
		if department<>"" then
            rs("department")=department
        end if
		if purpose<>"" then
            rs("purpose")=purpose
        end if
		if purpose_remark<>"" then
            rs("purpose_remark")=purpose_remark
        end if
		if howknow<>"" then
            rs("howknow")=howknow
        end if
		if howknow_remark<>"" then
            rs("howknow_remark")=howknow_remark
        end if
		if note<>"" then
            rs("note")=note
        end if
		rs("initdate") = date()
		rs("sytle_yn") = "N"
		rs("strategy_yn") = "N"
		rs("enable") = "Y"

		rs("classify") = "S"

		rs.Update
        if Err.Number=0 then 
			if flag="1" then
				response.redirect sender
			else
			'	���U�����g�ݨ� 
				response.redirect "studentedit.asp?sid=" & sid 
			end if
          
        else
            showmessage= Err.Description
        end if

	else
		showmessage="�w�g���U�A�Фŭ��s���U�C"
	end if

else

	sql = "select s30_student.std_id, s30_student.std_name,s30_student.std_birthday ,s30_student.std_sex,s30_student.std_idno "
	sql = sql & " ,s30_student.std_tel,s30_student.std_mobile,s30_student.std_email,s90_class.cls_id_abr , s90_class.cls_year  , "
	sql = sql & " s90_degree.dgr_name , s90_division.dvs_name , s90_unit.unt_name_abr,s90_class.cls_id "
	sql = sql & " from s30_student , s30_sturec , s90_yms  , s90_class , s90_degree , s90_division , s90_unit "
	sql = sql & " where s90_yms.yms_mark = 'Y' and "
    sql = sql & "  s90_yms.yms_year = s30_sturec.yms_year and  "
    sql = sql & "  s90_yms.yms_smester = s30_sturec.yms_sms and  "
    sql = sql & "  s30_student.std_key = s30_sturec.std_key and  "
    sql = sql & "  s30_sturec.cls_id = s90_class.cls_id  and  "
    sql = sql & "  s90_class.dgr_id = s90_degree.dgr_id and  "
    sql = sql & "  s90_degree.dvs_id = s90_division.dvs_id and  "
    sql = sql & "  s90_class.unt_id = s90_unit.unt_id and "
    sql = sql & " s30_student.std_id = '"&sid&"'   " 

'response.write sql


	rs.Open sql,syconn,adOpenStatic,adLockReadonly
	if not rs.EOF then
		sid = trim(rs("std_id"))
		name = trim(rs("std_name"))
		birthday = trim(rs("std_birthday"))
		sex = trim(rs("std_sex"))
		idno = trim(rs("std_idno"))
		mail = trim(rs("std_email"))
		cell = ifnull(trim(rs("std_mobile")),trim(rs("std_tel")) )
		classcode = trim(rs("cls_id_abr"))
		
		department = trim(rs("unt_name_abr"))
		grade = trim(rs("cls_year"))
		slevel = trim(rs("dgr_name"))
		class1 = replace(replace(replace(replace(right(trim(rs("cls_id")),1),"1","A"),"2","B"),"3","C"),"4","D")
		

	else
		showmessage = "�D�X�k�ǭ��A�L�k���U�C"
		sid=""
	end if
	rs.close

end if


'�t��
set rsLoad = server.CreateObject("adodb.recordset")
sql ="select * from s90_unit where unt_std='Y' "
rsLoad.Open sql,syconn,adOpenStatic,adLockReadonly
StrDepartment="<option value=''> - �|�����w - </option>"

while not rsLoad.EOF
	if department=rsLoad("unt_name_abr") then
		StrDepartment=StrDepartment&"<option value="""&rsLoad("unt_name_abr")&""" selected>"&  rsLoad("unt_name_abr")&"</option>"
	else
		StrDepartment=StrDepartment&"<option value="""&rsLoad("unt_name_abr")&""" >"&  rsLoad("unt_name_abr")&"</option>"
	end if 
	rsLoad.MoveNext 
wend
rsLoad.close

'�Ǩ�
sql ="select * from s90_degree "
rsLoad.Open sql,syconn,adOpenStatic,adLockReadonly
Strslevel="<option value=''> - �|�����w - </option>"
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
<TITLE> �iLDCC�^�~�y��O�E�_���ɤ��ߡj </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/stm31.js" type="text/javascript"></SCRIPT>
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>
<script language="javascript">

function check_input()
{
    var errmsg=""
	
	if (form1.cell.value=="")
        errmsg += "�q�ܤ��ର�ť�\n";
    if (form1.name.value=="")
        errmsg += "�m�W���ର�ť�\n";
	if (form1.birthday.value=="")
        errmsg += "�ͤ餣�ର�ť�\n";
	if (form1.department.value=="" )
        errmsg += "��t���ର�ť�\n";
    if (form1.grade.value=="" )
        errmsg += "�~�Ť��ର�ť�\n";
	if (form1.mail.value=="")
        errmsg += "E-mail���ର�ť�\n";
	if (form1.purpose1.checked==true)
        errmsg += "�ӳX���ߪ��ت��Opurpose of visiting this center�������\n";
	if (form1.howknow1.checked==true)
        errmsg += "�A�p��o��������How do you know about this center�������\n";
	
	if (errmsg == "")
	{
		var obj=document.getElementById("grade");
		if (obj.value!=""){
			objvalue=obj.value;
			if ( !isint(obj.value))
				errmsg += "�~�ť������b�μƦr\n";  
		}
	}
	
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">�w�����U�@�~ </TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="register.asp" onsubmit="return check_input();">
			<input type="hidden" value="add" name="validate">
			<input type="hidden" id="sender" name="sender" value="<%=sender%>">
			<input type="hidden" id="flag" name="flag" value="<%=flag%>">
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>�Ǹ��G</TD>
						<TD>�m�W�G</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=sid%>" maxlength="25" size="35" name="sid" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
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
						<TD>�q�ܡG</TD>
						<TD>�ͤ�G</TD>
						<TD></TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=cell%>" maxlength="30" size="30"  name="cell" class="inputtext" >
						</TD>
						<TD>
						<input type="text" value="<%=birthday%>" maxlength="25"  name="birthday" class="inputtext" readonly>
						<img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('birthday')" class="showhand">&nbsp;
						</TD>
						<TD>
							<TABLE cellSpacing=0 cellPadding=0  border=0 >
							<TR>
							<TD class="inputlabel">�ʧO�G</TD>
							<TD><input type="radio" name="sex" class="inputtext" value="M" <%if sex="M" then response.write "checked" end if%> ></TD><TD>�k</TD>
							<TD><input type="radio" name="sex" class="inputtext" value="F" <%if sex="F" then response.write "checked" end if%>></TD><TD>�k</TD>
							</TR>
							</TABLE>
						</TD>
						
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>�Ǩ�G</TD>
						<TD>�t�ҡG</TD>
						<TD>�~�šG</TD>
						
						<TD>�Z�šG</TD>
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
						<input type="text" value="<%=grade%>" maxlength="10" size="10"  name="grade" class="inputtext" >
						</TD>
						
						<TD>
						<input type="text" value="<%=class1%>" maxlength="25"  name="class1" class="inputtext" >
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>Email�G(�ж�g�Ǯ�Email�A�H����|�H��C)</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=mail%>" maxlength="100"  size="50" name="mail" class="inputtext" >
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD colspan="2">�ӳX���ߪ��ت��Opurpose of visiting this center�G</TD>
					</TR>
					<TR>
						<TD><input type="radio" name="purpose" id="purpose1" class="inputtext" checked value="" <%if purpose="" then response.write "checked" end if%> ></TD>
						<TD>�����w</TD>
					</TR>
					<TR>
						<TD><input type="radio" name="purpose" class="inputtext" value="dc" <%if purpose="dc" then response.write "checked" end if%> ></TD>
						<TD>�E�_�԰�(�^��ǲߤ�k)Diagnosis and Consultation(dc)</TD>
					</TR>
					<TR>
						<TD><input type="radio" name="purpose" class="inputtext" value="software" <%if purpose="software" then response.write "checked" end if%> ></TD>
						<TD>�ϥέ^�y�۾ǳn��Englisg Learning Software(software)</TD>
					</TR>
					<TR>
						<TD><input type="radio" name="purpose" class="inputtext" value="op" <%if purpose="op" then response.write "checked" end if%> ></TD>
						<TD>�f�y�m��Oral Practice(op)</TD>
					</TR>
					<TR>
						<TD><input type="radio" name="purpose" class="inputtext" value="workshops" <%if purpose="workshops" then response.write "checked" end if%> ></TD>
						<TD>�^�y�ǲߤ�k���yWorkshops</TD>
					</TR>
					<TR>
						<TD><input type="radio" name="purpose" class="inputtext" value="test-prep" <%if purpose="test-prep" then response.write "checked" end if%> ></TD>
						<TD>�y����������Simulation Tests(test-prep)</TD>
					</TR>
					<TR>
						<TD><input type="radio" name="purpose" class="inputtext" value="other" <%if purpose="other" then response.write "checked" end if%> ></TD>
						<TD>��L Other &nbsp;<input type="text" value="<%=purpose_remark%>" maxlength="25" size="50" name="purpose_remark" class="inputtext" ></TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD colspan="2">�A�p��o��������How do you know about this center�G</TD>
					</TR>
					<TR>
						<TD><input type="radio" name="howknow" id="howknow1" class="inputtext" checked value="" <%if howknow="" then response.write "checked" end if%> ></TD>
						<TD>�����w</TD>
					</TR>
					<TR>
						<TD><input type="radio" name="howknow" class="inputtext" value="brochures" <%if howknow="brochures" then response.write "checked" end if%> ></TD>
						<TD>�p�U�l Brochures</TD>
					</TR>
					<TR>
						<TD><input type="radio" name="howknow" class="inputtext" value="Teachers or Classmates" <%if howknow="Teachers or Classmates" then response.write "checked" end if%> ></TD>
						<TD>�Ѯv�ΦP�ǪB�ͧi�� Teachers or Classmates</TD>
					</TR>
					<TR>
						<TD><input type="radio" name="howknow" class="inputtext" value="Website" <%if howknow="Website" then response.write "checked" end if%> ></TD>
						<TD>�q�����W Website</TD>
					</TR>
					
					<TR>
						<TD><input type="radio" name="howknow" class="inputtext" value="other" <%if howknow="other" then response.write "checked" end if%> ></TD>
						<TD>��L Other <input type="text" value="<%=howknow_remark%>" maxlength="25"  name="howknow_remark" size="50" class="inputtext" ></TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>�Ƶ� Note�G</TD>
					</TR>
					<TR>
						<TD>
						<textarea  cols="100" rows="5" name="note" class="inputtext" ><%=note%></textarea>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR>
			<TD>
			<BR>
			<input  type="submit" value="�T�{����U" class="inputbutton" >
			<%if btncontrol="Y" then%>
			<input  type="button" value="��^" class="inputbutton" onclick="window.location='<%=sender%>'">
			<%end if%>
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
<%

if flag<>"1" then

%>
<script language="javascript">
function window.onload()
{
	var ls_parm = 'dialogWidth=650px;'
					+ 'dialogHeight=650px;'
					+ 'center=yes;'
					+ 'border=thin;'
					+ 'help=no;'
					+ 'directories=no;'
					+ 'location=no;'
					+ 'status=no'
	window.open('showrule.asp','�T�����i','fullscreen=1,scrollbars=1');
}
</script>
<%end if%>

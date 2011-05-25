<% SESSION.CODEPAGE="65001"%>
<% Response.CacheControl = "No-cache" %>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	window.close();
}

//-->
</SCRIPT>
<SCRIPT LANGUAGE=javascript FOR=window EVENT=onload>
<!--
 window_onload();
//-->
</SCRIPT>
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->

<!-- INCLUDE END -->
<%
sid=trim(request("sid"))
fname=trim(request("fname"))
name = ""
birthday = ""
sex = ""
idno = ""
mail =""
cell = ""
classcode = ""
department = ""
grade = ""
slevel = ""
class1 = ""

if sid <>"" then
	set rs = server.CreateObject("adodb.recordset")
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
		showmessage = "非合法學員，無法註冊。"
		sid=""
	end if
	rs.close
	set rs=nothing
end if
response.write "name=" & name

response.write "cell=" & cell
%>

<SCRIPT LANGUAGE=javascript>
<!--

		window.parent.document.getElementById("name").value="<%=name%>";
		window.parent.document.getElementById("sid").value="<%=sid%>";
		window.parent.document.getElementById("birthday").value="<%=birthday%>";
		window.parent.document.getElementById("sex").value="<%=sex%>";
		window.parent.document.getElementById("idno").value="<%=idno%>";
		window.parent.document.getElementById("mail").value="<%=mail%>";
		window.parent.document.getElementById("cell").value="<%=cell%>";
		window.parent.document.getElementById("classcode").value="<%=classcode%>";
		window.parent.document.getElementById("department").value="<%=department%>";
		window.parent.document.getElementById("grade").value="<%=grade%>";
		window.parent.document.getElementById("slevel").value="<%=slevel%>";
		window.parent.document.getElementById("class1").value="<%=class1%>";


		<%if fname<>"" then%>
			//	window.parent.execScript("<%=fname%>()");
		<%end if%>
			
//-->
</SCRIPT>
<HTML>
<HEAD>
<TITLE>  </TITLE>
<META http-equiv=Content-Type content="text/html; charset=utf-8">
</HEAD>
</HTML> 
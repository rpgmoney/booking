<%@Language=VBScript LCID=1033%>
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
name=""
email=""
deptname=""
sts_id=""
secondary=""
majorcourse=""
if sid <>"" then
	set rs = server.CreateObject("adodb.recordset")
	set rs2 = server.CreateObject("adodb.recordset")
	'學員資訊
		sql = " select emp_id as sid,emp_name as name,emp_email as email,s90_unit.unt_name_abr  as department ,'' as grade,'教職員' as slevel,'01' as sts_id "
		sql = sql & "  from s10_employee,s90_unit "
		sql = sql & " where (emp_leavedate is null or emp_leavedate='' ) "
		sql = sql &  " and ( emp_id='"&sid&"' or emp_name = '"&sid&"') "
		sql = sql & " and s10_employee.emp_untid = s90_unit.unt_id "
		sql = sql & " union "
		sql = sql & " select s30_student.std_id as sid, s30_student.std_name as name"
		sql = sql & ",s30_student.std_email  as email, s90_unit.unt_name_abr as  department, s90_class.cls_year as grade,s90_degree.dgr_name as slevel,'02' as sts_id "
		sql = sql & " from s30_student , s30_sturec , s90_yms,s90_class,s90_unit ,s90_degree"
		sql = sql & " where s90_yms.yms_mark = 'Y' and  s90_yms.yms_year = s30_sturec.yms_year "
		sql = sql &  " and (s30_student.std_id='"&sid&"' or s30_student.std_name ='"&sid&"') "
		sql = sql & " and  s90_yms.yms_smester = s30_sturec.yms_sms and s30_student.std_key = s30_sturec.std_key  "
		sql = sql & "  and  s30_sturec.cls_id = s90_class.cls_id    "
		sql = sql & "  and  s90_class.unt_id = s90_unit.unt_id  "
		sql = sql & "  and s90_class.dgr_id = s90_degree.dgr_id   "
response.write sql
	
	rs.Open sql,syconn,adOpenStatic,adLockReadonly
	if not rs.EOF then
		sid = trim(rs("sid"))
		name = trim(rs("name"))
		email = trim(rs("email"))
		deptname = trim(rs("department"))
		grade = trim(rs("grade"))
		slevel = trim(rs("slevel"))
		sts_id = trim(rs("sts_id"))
		
	else
		showmessage = " 職號或學號有誤。"
		sid=""
	end if
	rs.close
	set rs=nothing
	set rs2=nothing
end if
response.write "name=" & name
response.write "showmessage=" & showmessage

%>

<SCRIPT LANGUAGE=javascript>
<!--
	

		if (window.parent.document.getElementById("name")!=null)
		{window.parent.document.getElementById("name").value="<%=name%>";}
		if (window.parent.document.getElementById("sid")!=null)
		{window.parent.document.getElementById("sid").value="<%=sid%>";}
		if (window.parent.document.getElementById("deptname")!=null)
		{window.parent.document.getElementById("deptname").value="<%=deptname%>";}


		<%if fname<>"" then%>
			//	window.parent.execScript("<%=fname%>()");
		<%end if%>
//-->
</SCRIPT>

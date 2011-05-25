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

<!-- INCLUDE END -->
<%
sid=trim(request("sid"))
languagecode=trim(request("languagecode"))
name=""
slevel=""
grade=""
class1=""
department=""
score=cdbl(0)
StrSubject=""
StrOralset=""
StrOrallevel=""
StrSubject2=""
StrOralset2=""
StrOrallevel2=""

if sid <>"" then
	set rs = server.CreateObject("adodb.recordset")
	set rs2 = server.CreateObject("adodb.recordset")
	'學員資訊
	sql="select * from boo_profile where sid='"&sid&"' and classify='S' and  enable='Y' and sytle_yn='Y'  and  strategy_yn='Y' "
	'response.write sql
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
	
	if not rs.EOF then
		name=rs("name")
		slevel=rs("slevel")
		grade=rs("grade")
		class1=rs("class1")
		department=rs("department")
	else
		name="輸入學號有誤，或該學號尚未註冊。"
		sid=""
	end if
	rs.close
	'英檢成績
	sql = " SELECT DISTINCT "
	sql = sql & " s30_student.std_id, "
	sql = sql & " s30_student.std_name, "
	sql = sql & " s305_langtest_score.yms_year, "
	sql = sql & " s305_langtest_score.cnt, "
	sql = sql & " s305_langtest_score.ltk_id, "
	sql = sql & " s305_lang_kind.ltk_name , "
	sql = sql & " s305_langtest_score.level_id, "
	sql = sql & " CONVERT(VarChar(10),s305_langtest_score.tot_score) as tot_score, "
	sql = sql & " s305_langtest_score.test_id, "
	sql = sql & " s305_langtest_score.test_date, "
	sql = sql & " s90_class.cls_name_abr , s90_class.cls_id  "
	sql = sql & " FROM s30_student, "   
	sql = sql & " s30_sturec, "   
	sql = sql & " s305_langtest_score  , "
	sql = sql & " s90_class , s305_lang_kind  , s90_unit , s90_yms "
	sql = sql & " WHERE ( s30_student.std_key = s30_sturec.std_key ) and "  
		 sql = sql & " ( s305_langtest_score.ltk_id   = s305_lang_kind.ltk_id ) and "
		 sql = sql & " ( s305_langtest_score.std_key  = s30_sturec.std_key ) and "         
		 sql = sql & " ( s90_unit.unt_id = s90_class.unt_id ) and "
		 sql = sql & " (  s30_sturec.yms_year = s90_yms.yms_year and s30_sturec.yms_sms = s90_yms.yms_smester  ) and "
		 sql = sql & " s90_yms.yms_mark='Y' and  "
		 sql = sql & " ( s30_sturec.cls_id = s90_class.cls_id ) and  "     
		 sql = sql & " (  s30_student.std_id =  '"&sid&"') and "
		 sql = sql & " ( s305_langtest_score.ltk_id = 'E111' ) and "
		 sql = sql & " ( s30_sturec.src_status = '0' ) "
	sql = sql & " order by s305_langtest_score.test_date  desc"
	'response.write sql & "<br>"
	rs.Open sql,syconn,adOpenStatic,adLockReadonly
	
	if not rs.EOF then
		score=cdbl(trim(rs("tot_score")))
		'score=60
	else
		score="0"
		'StrOrallevel2 = "<input type=hidden value=Level1 name=orallevel >"
	end if
	response.write  "score=" &  score & "<br>"
	rs.close

	'口語題目
	if (cdbl(score)>=90) and languagecode="E" then
		StrOrallevel2 = "<select id=orallevel name=orallevel class=inputtext><option value=''>請指定口語級數</option>"
		StrOrallevel2 = StrOrallevel2 & "<option value='Level 1'>Level 1</opton><option value='Level 2'>Level 2</opton>"
		StrOrallevel2 = StrOrallevel2 & "<option value='Level 3'>Level 3</opton><option value='Level 4'>Level 4</opton>"
		StrOrallevel2 = StrOrallevel2 & "</select>"

		StrOralset = "<option value='' selected>請指定口語系列</option>"
		StrOralset = StrOralset &"<option value=""My ET"" > My ET</option>"
		StrOralset = StrOralset &"<option value=""Issues in English I"" > Issues in English I </option>"
		StrOralset = StrOralset & "<option value=""Issues in English II"" > Issues in English II </option>"
		
		StrSubject="<option value="""" selected>- 無 -</option>"
	else

		'sql ="select * from boo_orallevel where rank='1'"
		'rs2.Open sql,msconn,adOpenStatic,adLockReadonly
		'if not rs2.EOF then
		'	StrSubject="<option value="""" selected>- 請指定口語題目 -</option>"
		'else
			StrSubject="<option value="""" selected>- 無 -</option>"
		'end if

		'while not rs2.EOF
		'	StrSubject=StrSubject&"<option value="""&rs2("topic")&""" >"&  rs2("topic")&"</option>"
		'	rs2.MoveNext 
		'wend


		StrOralset = "<option value='' selected>請指定口語系列</option>"
		StrOralset = StrOralset &"<option value=""My ET"" > My ET</option>"
		StrOralset = StrOralset &"<option value=""Conversation Topics"" > Conversation Topics </option>"
		
		StrOrallevel2 = "<input type=hidden value=Level1 name=orallevel >"

		rs2.close
		
	end if			



end if

StrSubject2="<select id=topic name=topic class=inputtext >"& replace(StrSubject,"""","'") & "</select>"
StrOralset2="<select id=oralset name=oralset class=inputtext onchange=changesubject() >"& replace(StrOralset,"""","'") & "</select>"



response.write "sid=" & sid
%>

<SCRIPT LANGUAGE=javascript>
<!--
	if (window.self.opener!=null)
	{
		if (window.self.opener.name!=null)
		{window.self.opener.name.value="<%=name%>";}
		if (window.self.opener.sid!=null)
		{window.self.opener.sid.value="<%=sid%>";}
		if (window.self.opener.slevel!=null)
		{window.self.opener.slevel.value="<%=slevel%>";}
		if (window.self.opener.grade!=null)
		{window.self.opener.grade.value="<%=grade%>";}
		if (window.self.opener.class1!=null)
		{window.self.opener.class1.value="<%=class1%>";}
		if (window.self.opener.department!=null)
		{window.self.opener.department.value="<%=department%>";}
		if (window.self.opener.score!=null)
		{window.self.opener.score.value="<%=score%>";}
		if (window.self.opener.topic!=null)
		{window.self.opener.topic.outerHTML="<%=StrSubject2%>";}
		if (window.self.opener.oralset!=null)
		{window.self.opener.oralset.outerHTML="<%=StrOralset2%>";}
		if (window.self.opener.orallevel!=null)
		{window.self.opener.orallevel.outerHTML="<%=StrOrallevel2%>";}
		
		
	}		
//-->
</SCRIPT>

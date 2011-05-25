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
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->


<!-- INCLUDE END -->
<%
sid=trim(request("sid"))

score=""
response.write "sid=" & sid

if sid <>"" then
	set rs = server.CreateObject("adodb.recordset")

	sql = " SELECT DISTINCT "
	sql = sql & " s30_student.std_id, "
	sql = sql & " s30_student.std_name, "
	sql = sql & " s305_langtest_score.yms_year, "
	sql = sql & " s305_langtest_score.cnt, "
	sql = sql & " s305_langtest_score.ltk_id, "
	sql = sql & " s305_lang_kind.ltk_name , "
	sql = sql & " s305_langtest_score.level_id, "
	sql = sql & " s305_langtest_score.tot_score, "
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
	response.write sql
	rs.Open sql,syconn,adOpenStatic,adLockReadonly
	
	if not rs.EOF then
		score=rs("tot_score")
	else
		score="0"
	end if
	rs.close


end if
response.write "score=" & score

%>

<SCRIPT LANGUAGE=javascript>
<!--
	if (window.self.opener!=null)
	{
		if (window.self.opener.score!=null)
		{window.self.opener.score.value="<%=score%>";}
		
	}		
//-->
</SCRIPT>

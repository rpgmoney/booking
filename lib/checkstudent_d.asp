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
<!-- #INCLUDE file="parameter.inc" -->
<!-- INCLUDE END -->
<%
sid=trim(request("sid"))
item=trim(request("item"))
category=trim(request("category"))
'response.write "item=" & item
name=""
slevel=""
grade=""
class1=""
department=""
score=""
tid=""
if sid <>"" then
	set rs = server.CreateObject("adodb.recordset")
	set rs2 = server.CreateObject("adodb.recordset")
	'學員資訊
	sql="select * from boo_profile where sid='"&sid&"' and classify='S' and  enable='Y' and sytle_yn='Y'  and  strategy_yn='Y'  "
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
	'response.write sql
	rs.Open sql,syconn,adOpenStatic,adLockReadonly
	
	if not rs.EOF then
		score=cdbl(trim(rs("tot_score")))
		'score=60
	else
		score="0"
	end if
	response.write "score=" & score
	rs.close
	if item ="診斷" or item="諮商" then

	else
		
		if  cdbl(score) < cdbl(par_score) then
			'處方籤之回診日期若超過30天則需重新診斷
			tmpdate = datetoNumformat(dateadd("d",par_extinct_day,date()))
			'根據處方秀出可預約項目
			sql = "select b.* from boo_book_T_M a  "
			sql = sql & " inner join boo_diagnosis b on a.id=b.tid  where a.item='診斷'  and  a.signin is not null  and  a.sid='"&sid&"' "
			sql = sql & " and  Cast(b.backdate as int) >='" & tmpdate & "' "
			sql = sql & " order by Cast(a.bdate as int)  desc "
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			if not rs.EOF then
				tid = trim(rs("tid"))'處方單號
				
				if item="口語" then
					optime = ifnull(trim(rs("optime")),0)
					optime_b = ifnull(trim(rs("optime_b")),0)
					'response.write "optime=" & optime & "<br>"
					'response.write "optime_b=" & optime_b & "<br>"
					if  int(optime) <= int(optime_b) then
						name="該學員尚未有該項處方，不得預約。"
						sid=""
					end if
				end if
				if item="詩歌" then
					crkptime = ifnull(trim(rs("crkptime")),0)
					crkptime_b = ifnull(trim(rs("crkptime_b")),0)
					if  int(crkptime) <= int(crkptime_b) then
						name="該學員尚未有該項處方，不得預約。"
						sid=""
					end if
				end if
				if item="簡報" then
					pptime = ifnull(trim(rs("pptime")),0)
					pptime_b = ifnull(trim(rs("pptime_b")),0)
					if  int(pptime) <= int(pptime_b) then
						name="該學員尚未有該項處方，不得預約。"
						sid=""
					end if
				end if
				if item="寫作" then
					writetime = ifnull(trim(rs("writetime")),0)
					writetime_b = ifnull(trim(rs("writetime_b")),0)
					if  int(writetime) <= int(writetime_b) then
						name="該學員尚未有該項處方，不得預約。"
						sid=""
					end if
				end if
				if item="閱讀" then
					readtime = ifnull(trim(rs("readtime")),0)
					readtime_b = ifnull(trim(rs("readtime_b")),0)
					if  int(readtime) <= int(readtime_b) then
						name="該學員尚未有該項處方，不得預約。"
						sid=""
					end if
				end if
			else
				name="該學員尚未有該項處方，不得預約。"
				sid=""

			end if

			rs.Close
		else
			'大專英檢分數超過240分,全部項目都可以預約
			'處方籤之回診日期若超過30天則需重新診斷
			tmpdate = datetoNumformat(dateadd("d",par_extinct_day,date()))
			'根據處方秀出可預約項目
			sql = "select b.* from boo_book_T_M a  "
			sql = sql & " inner join boo_diagnosis b on a.id=b.tid  where a.item='診斷'  and  a.signin is not null  and  a.sid='"&sid&"' "
			sql = sql & " and  Cast(b.backdate as int ) >='" & tmpdate & "' "
			sql = sql & " order by Cast(a.bdate as int) desc "
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			if not rs.EOF then
				tid = trim(rs("tid"))'處方單號
			end if
			rs.Close
		end if 'if  cdbl(score) < cdbl(par_score) then





	end if



end if





%>

<SCRIPT LANGUAGE=javascript>
<!--
//	if (window.self.opener!=null)
//	{
	//	if (window.self.opener.name_d!=null)
//		{window.self.opener.name_d.value="<%=name%>";}
//		if (window.self.opener.sid_d!=null)
//		{window.self.opener.sid_d.value="<%=sid%>";}
//		if (window.self.opener.slevel_d!=null)
//		{window.self.opener.slevel_d.value="<%=slevel%>";}
//		if (window.self.opener.grade_d!=null)
//		{window.self.opener.grade_d.value="<%=grade%>";}
//		if (window.self.opener.class1_d!=null)
//		{window.self.opener.class1_d.value="<%=class1%>";}
//		if (window.self.opener.department_d!=null)
//		{window.self.opener.department_d.value="<%=department%>";}
//		if (window.self.opener.score_d!=null)
//		{window.self.opener.score_d.value="<%=score%>";}
//		if (window.self.opener.tid_d!=null)
//		{window.self.opener.tid_d.value="<%=tid%>";}

		if (window.parent.document.getElementById("name_d")!=null)
		{window.parent.document.getElementById("name_d").value="<%=name%>";}
		if (window.parent.document.getElementById("sid_d")!=null)
		{window.parent.document.getElementById("sid_d").value="<%=sid%>";}
		if (window.parent.document.getElementById("slevel_d")!=null)
		{window.parent.document.getElementById("slevel_d").value="<%=slevel%>";}
		if (window.parent.document.getElementById("grade_d")!=null)
		{window.parent.document.getElementById("grade_d").value="<%=grade%>";}
		if (window.parent.document.getElementById("class1_d")!=null)
		{window.parent.document.getElementById("class1_d").value="<%=class1%>";}
		if (window.parent.document.getElementById("department_d")!=null)
		{window.parent.document.getElementById("department_d").value="<%=department%>";}
		if (window.parent.document.getElementById("score_d")!=null)
		{window.parent.document.getElementById("score_d").value="<%=score%>";}
		if (window.parent.document.getElementById("tid_d")!=null)
		{window.parent.document.getElementById("tid_d").value="<%=tid%>";}
//	}		
//-->
</SCRIPT>

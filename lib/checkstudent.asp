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
languagecode=trim(request("languagecode"))
category=trim(request("category"))
response.write "category=" & category
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
tid=""
opFlag = "none"
ppFlag = "none"
crkpFlag="none"
writeFlag = "none"
readFlag="none"
today=year(date())-1911 & right("0" & month(date()),2) & right("0" & day(date()),2)

if sid <>"" then
	set rs0 = server.CreateObject("adodb.recordset")
	set rs = server.CreateObject("adodb.recordset")
	set rs2 = server.CreateObject("adodb.recordset")
	'學員資訊
	sql="select * from boo_profile where sid='"&sid&"' and classify in ('S','E') and  enable='Y' and sytle_yn='Y'  and  strategy_yn='Y' "
	'response.write sql
	rs0.Open sql,msconn,adOpenStatic,adLockReadonly
	
	if not rs0.EOF then
		name=rs0("name")
		slevel=rs0("slevel")
		grade=rs0("grade")
		class1=rs0("class1")
		department=rs0("department")

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

			'rs2.close
			
		end if
		if  cdbl(score) < cdbl(par_score) then
			'處方籤之回診日期若超過30天則需重新診斷
			tmpdate = datetoNumformat(dateadd("d",par_extinct_day,date()))
			'根據處方秀出可預約項目
			sql = "select b.* from boo_book_T_M a  "
			sql = sql & " inner join boo_diagnosis b on a.id=b.tid  where a.item='診斷'  and  a.signin is not null  and  a.sid='"&sid&"' "
			sql = sql & " and  Cast(b.backdate as int ) >='" & tmpdate & "' "
			sql = sql & " order by Cast( a.bdate as int ) desc "

			response.write sql
			'response.end
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			if not rs.EOF then
				tid = trim(rs("tid"))'處方單號
				optime = trim(rs("optime"))
				optime_b = ifnull(trim(rs("optime_b")),0)
				if  int(optime) > int(optime_b) then
					opFlag = "block"
				else
					opFlag = "none"
				end if
				crkptime = trim(rs("crkptime"))
				crkptime_b = ifnull(trim(rs("crkptime_b")),0)
				if  int(crkptime) > int(crkptime_b) then
					crkpFlag = "block"
				else
					crkpFlag = "none"
				end if
				if category="T" then
					pptime = trim(rs("pptime"))
					pptime_b = ifnull(trim(rs("pptime_b")),0)
					if  int(pptime) > int(pptime_b) then
						ppFlag = "block"
					else
						ppFlag = "none"
					end if
					writetime = trim(rs("writetime"))
					writetime_b = ifnull(trim(rs("writetime_b")),0)
					if  int(writetime) > int(writetime_b) then
						writeFlag = "block"
					else
						writeFlag = "none"
					end if
					readtime = trim(rs("readtime"))
					readtime_b = ifnull(trim(rs("readtime_b")),0)
					if  int(readtime) > int(readtime_b) then
						readFlag = "block"
					else
						readFlag = "none"
					end if
				end if
			else
				opFlag = "none"
				ppFlag = "none"
				crkpFlag="none"
				writeFlag = "none"
				readFlag="none"

			end if

			rs.Close
		else
			'處方籤之回診日期若超過30天則需重新診斷
			tmpdate = datetoNumformat(dateadd("d",par_extinct_day,date()))
			'根據處方秀出可預約項目
			sql = "select b.* from boo_book_T_M a  "
			sql = sql & " inner join boo_diagnosis b on a.id=b.tid  where a.item='診斷'  and  a.signin is not null  and  a.sid='"&sid&"' "
			sql = sql & " and  Cast(b.backdate as int) >='" & tmpdate & "' "
			sql = sql & " order by  Cast(a.bdate as int)  desc "
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			if not rs.EOF then
				tid = trim(rs("tid"))'處方單號
			end if
			rs.Close
			opFlag = "block"
			ppFlag = "block"
			crkpFlag="block"
			writeFlag = "block"
			readFlag="block"
		end if


	else
		name="輸入學號有誤，或該學號尚未註冊。"
		sid=""
	end if
	rs0.close


else

	opFlag = "none"
	ppFlag = "none"
	crkpFlag="none"
	writeFlag = "none"
	readFlag="none"

end if

'若是預約小老師則只開放口語和詩歌
if  category="ST" then
	ppFlag = "none"
	writeFlag = "none"
	readFlag="none"
end if

StrSubject2="<select id=topic name=topic class=inputtext >"& replace(StrSubject,"""","'") & "</select>"
StrOralset2="<select id=oralset name=oralset class=inputtext onchange=changesubject() >"& replace(StrOralset,"""","'") & "</select>"



response.write "sid=" & sid
response.write "tid=" & tid
%>

<SCRIPT LANGUAGE=javascript>
<!--
//	if (window.self.opener!=null)
//	{
		

		if (window.parent.document.getElementById("name")!=null)
		{window.parent.document.getElementById("name").value="<%=name%>";}
		if (window.parent.document.getElementById("sid")!=null)
		{window.parent.document.getElementById("sid").value="<%=sid%>";}
		if (window.parent.document.getElementById("slevel")!=null)
		{window.parent.document.getElementById("slevel").value="<%=slevel%>";}
		if (window.parent.document.getElementById("grade")!=null)
		{window.parent.document.getElementById("grade").value="<%=grade%>";}
		if (window.parent.document.getElementById("department")!=null)
		{window.parent.document.getElementById("department").value="<%=department%>";}
		
		if (window.parent.document.getElementById("class1")!=null)
		{window.parent.document.getElementById("class1").value="<%=class1%>";}
		if (window.parent.document.getElementById("score")!=null)
		{window.parent.document.getElementById("score").value="<%=score%>";}
		if (window.parent.document.getElementById("topic")!=null)
		{window.parent.document.getElementById("topic").outerHTML="<%=StrSubject2%>";}
		if (window.parent.document.getElementById("oralset")!=null)
		{window.parent.document.getElementById("oralset").outerHTML="<%=StrOralset2%>";}
		if (window.parent.document.getElementById("orallevel")!=null)
		{window.parent.document.getElementById("orallevel").outerHTML="<%=StrOrallevel2%>";}

		if (window.parent.document.getElementById("tid")!=null)
		{window.parent.document.getElementById("tid").value="<%=tid%>";}

		var item3_o=window.parent.document.getElementById("item3_option");
		item3_o.style.display="<%=opFlag%>";
		var item3_l=window.parent.document.getElementById("item3_lab");
		item3_l.style.display="<%=opFlag%>";
		var item4_o=window.parent.document.getElementById("item4_option");
		item4_o.style.display="<%=ppFlag%>";
		var item4_l=window.parent.document.getElementById("item4_lab");
		item4_l.style.display="<%=ppFlag%>";

		var item5_o=window.parent.document.getElementById("item5_option");
		item5_o.style.display="<%=crkpFlag%>";
		var item5_l=window.parent.document.getElementById("item5_lab");
		item5_l.style.display="<%=crkpFlag%>";

		var item6_o=window.parent.document.getElementById("item6_option");
		item6_o.style.display="<%=writeFlag%>";
		var item6_l=window.parent.document.getElementById("item6_lab");
		item6_l.style.display="<%=writeFlag%>";

		var item7_o=window.parent.document.getElementById("item7_option");
		item7_o.style.display="<%=readFlag%>";
		var item7_l=window.parent.document.getElementById("item7_lab");
		item7_l.style.display="<%=readFlag%>";

	
//	}		
//-->
</SCRIPT>

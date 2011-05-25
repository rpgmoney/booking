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
StrItem=""
StrItem2=""
tid=""
today=year(date())-1911 & right("0" & month(date()),2) & right("0" & day(date()),2)
times="0"
if sid <>"" then
	set rs0 = server.CreateObject("adodb.recordset")
	set rs = server.CreateObject("adodb.recordset")
	set rs2 = server.CreateObject("adodb.recordset")
	'學員資訊
	sql="select * from boo_profile where sid='"&sid&"' and classify in('S','E') and  enable='Y' and sytle_yn='Y'  and  strategy_yn='Y' "
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
		if  cdbl(score) < cdbl(par_score) then
			'處方籤之回診日期若超過30天則需重新診斷
			tmpdate = datetoNumformat(dateadd("d",par_extinct_day,date()))
			sql = "select b.* from boo_book_T_M a  "
			sql = sql & " inner join boo_diagnosis b on a.id=b.tid  where a.item='診斷'  and  a.signin is not null  and  a.sid='"&sid&"' "
			sql = sql & " and  Cast(b.backdate as int) >='" & tmpdate & "' "
			sql = sql & " order by  Cast(a.bdate as int) desc "

			response.write "sql_1=" & sql
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			if not rs.EOF then
				tid = trim(rs("tid"))'處方單號
			end if
			
			rs.Close
			if category="S" then
				'軟體
				if tid<>"" then
					'根據處方秀出可預約項目
					sql = "select a.id as tid,d.id,d.floor,d.software,c.times,c.times_b,c.times_c,c.times-c.times_b as times1  from boo_book_T_M a  "
					sql = sql & " inner join boo_diagnosis b on a.id=b.tid "
					sql = sql & " inner  join boo_diagnosis_softwore c on a.id=c.tid  and c.category='S' "
					sql = sql & " inner join boo_software d on c.sid=d.id"
					sql = sql & " where a.item='診斷'  and  a.signin is not null  and  a.sid='"&sid&"' "
					sql = sql & " and  a.id ='" &tid & "' and c.times-c.times_b >0"

					rs2.Open sql,msconn,adOpenStatic,adLockReadonly
					if not rs2.EOF then
						StrItem="<option value="""" selected>- 請指定軟體 -</option>"
					else
						StrItem="<option value="""" selected>- 無 -</option>"
					end if

					while not rs2.EOF
						StrItem=StrItem&"<option dname="""&rs2("times1")&"""  value="""&rs2("id")&""" >"&rs2("floor") &"&nbsp;-&nbsp;"&  rs2("software")&"</option>"
						rs2.MoveNext 
					wend
					rs2.close
				else
					StrItem="<option value="""" selected>- 無 -</option>"
				end if
			else
				'補充教材
				if tid<>"" then
					'根據處方秀出可預約項目
					sql = "select a.id as tid,d.id,d.floor,d.software,c.times,c.times_b,c.times_c,c.times-c.times_b as times1  from boo_book_T_M a  "
					sql = sql & " inner join boo_diagnosis b on a.id=b.tid "
					sql = sql & " inner  join boo_diagnosis_softwore c on a.id=c.tid  and c.category='T' "
					sql = sql & " inner join boo_software d on c.sid=d.id"
					sql = sql & " where a.item='診斷'  and  a.signin is not null  and  a.sid='"&sid&"' "
					sql = sql & " and  a.id ='" &tid & "' and c.times-c.times_b >0"
response.write sql
					rs2.Open sql,msconn,adOpenStatic,adLockReadonly
					if not rs2.EOF then
						StrItem="<option value="""" selected>- 請指定補充教材 -</option>"
					else
						StrItem="<option value="""" selected>- 無 -</option>"
					end if

					while not rs2.EOF
						StrItem=StrItem&"<option dname="""&rs2("times1")&"""  value="""&rs2("id")&""" >"&  rs2("software")&"</option>"
						rs2.MoveNext 
					wend
					rs2.close
				else
					StrItem="<option value="""" selected>- 無 -</option>"
				end if

			end if
		else
			'處方籤之回診日期若超過30天則需重新診斷
			tmpdate = datetoNumformat(dateadd("d",par_extinct_day,date()))
			sql = "select b.* from boo_book_T_M a  "
			sql = sql & " inner join boo_diagnosis b on a.id=b.tid  where a.item='診斷'  and  a.signin is not null  and  a.sid='"&sid&"' "
			sql = sql & " and  Cast(b.backdate as int ) >='" & tmpdate & "' "
			sql = sql & " order by Cast(a.bdate as int) desc "
			rs.Open sql,msconn,adOpenStatic,adLockReadonly
			if not rs.EOF then
				tid = trim(rs("tid"))'處方單號
			end if
			rs.Close
			if category="S" then
				sql ="select * from boo_software where yn='Y' and category='S' order by floor,software"
				rs2.Open sql,msconn,adOpenStatic,adLockReadonly
				StrItem="<option value="""" selected>- 請指定軟體 -</option>"
				while not rs2.EOF
					StrItem=StrItem&"<option dname='120' value="""&rs2("id")&""" >"&rs2("floor") &"&nbsp;-&nbsp;"&  rs2("software")&"</option>"
					rs2.MoveNext 
				wend
				rs2.close
				times="120"
			else
				sql ="select * from boo_software where yn='Y' and category='T' order by floor,software"
				rs2.Open sql,msconn,adOpenStatic,adLockReadonly
				StrItem="<option value="""" selected>- 請指定補充教材 -</option>"
				while not rs2.EOF
					StrItem=StrItem&"<option dname='480' value="""&rs2("id")&""" >"&  rs2("software")&"</option>"
					rs2.MoveNext 
				wend
				rs2.close
				times="480"
			end if
			

		end if


	else
		name="輸入學號有誤，或該學號尚未註冊。"
		sid=""
	end if
	rs0.close



end if
StrItem2="<select id=item name=item onchange=itemchange()  class=inputtext >"& replace(StrItem,"""","'") & "</select>"

response.write "tid=" & tid
%>

<SCRIPT LANGUAGE=javascript>
<!--
//	if (window.self.opener!=null)
//	{
//		if (window.self.opener.name!=null)
//		{window.self.opener.name.value="<%=name%>";}
//		if (window.self.opener.sid!=null)
//		{window.self.opener.sid.value="<%=sid%>";}
//		if (window.self.opener.slevel!=null)
//		{window.self.opener.slevel.value="<%=slevel%>";}
//		if (window.self.opener.grade!=null)
//		{window.self.opener.grade.value="<%=grade%>";}
//		if (window.self.opener.class1!=null)
//		{window.self.opener.class1.value="<%=class1%>";}
//		if (window.self.opener.department!=null)
//		{window.self.opener.department.value="<%=department%>";}
//		if (window.self.opener.score!=null)
//		{window.self.opener.score.value="<%=score%>";}
//		if (window.self.opener.item!=null)
//		{window.self.opener.item.outerHTML="<%=StrItem2%>";}
//		if (window.self.opener.times!=null)
//		{window.self.opener.times.value="<%=times%>";}
//		if (window.self.opener.tid!=null)
//		{window.self.opener.tid.value="<%=tid%>";}

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

		if (window.parent.document.getElementById("item")!=null)
		{window.parent.document.getElementById("item").outerHTML="<%=StrItem2%>";}

		if (window.parent.document.getElementById("times")!=null)
		{window.parent.document.getElementById("times").value="<%=times%>";}
		if (window.parent.document.getElementById("tid")!=null)
		{window.parent.document.getElementById("tid").value="<%=tid%>";}
	
//	}		
//-->
</SCRIPT>

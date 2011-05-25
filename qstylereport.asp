<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 50 %>
<% Response.CacheControl = "No-cache" %>

<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<%
sid = left(trim(request("sid")),10)
initdate= trim(request("initdate"))
qsid= trim(request("qsid"))

set rs = server.CreateObject("adodb.recordset")

	'if sid<>"" then
	'sql = "select a.*,b.name from boo_questionnaire_style a left join boo_profile b on a.sid=b.sid where a.sid = '"&sid&"' and a.yn='Y' order by a.initdate desc"
	'else
	'sql = "select a.*,b.name from boo_questionnaire_style a left join boo_profile b on a.sid=b.sid where a.qsid = '"&qsid&"' and a.yn='Y' order by a.initdate desc"
	'end if

	if sid<>"" and initdate<>""  then
		sql = "select a.*,b.name from boo_questionnaire_style  a left join boo_profile b on a.sid=b.sid where a.sid = '"&sid&"' and a.initdate='"&initdate&"' and a.yn='Y' order by a.initdate desc"
	elseif sid<>"" and initdate=""  then
		sql = "select a.*,b.name from boo_questionnaire_style  a left join boo_profile b on a.sid=b.sid where a.sid = '"&sid&"' and a.yn='Y' order by a.initdate desc"
	else
		sql = "select a.*,b.name from boo_questionnaire_style  a left join boo_profile b on a.sid=b.sid where a.qsid = '"&qsid&"' and a.yn='Y' order by a.initdate desc"
	end if
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
	if not  rs.EOF then
		sid=trim(rs("sid"))
		q1 = trim(rs("q1"))
		q2 = trim(rs("q2"))
		q3 = trim(rs("q3"))
		q4 = trim(rs("q4"))
		q5 = trim(rs("q5"))
		q6 = trim(rs("q6"))
		q7 = trim(rs("q7"))
		q8 = trim(rs("q8"))
		q9 = trim(rs("q9"))
		q10 = trim(rs("q10"))
		q11 = trim(rs("q11"))
		q12 = trim(rs("q12"))
		q13 = trim(rs("q13"))
		q14 = trim(rs("q14"))
		q15 = trim(rs("q15"))
		q16 = trim(rs("q16"))
		q17 = trim(rs("q17"))
		q18 = trim(rs("q18"))
		q19 = trim(rs("q19"))
		q20 = trim(rs("q20"))
		q21 = trim(rs("q21"))
		q22 = trim(rs("q22"))
		q23 = trim(rs("q23"))
		q24 = trim(rs("q24"))
		q25 = trim(rs("q25"))
		q26 = trim(rs("q26"))
		q27 = trim(rs("q27"))
		q28 = trim(rs("q28"))
		q29 = trim(rs("q29"))
		q30 = trim(rs("q30"))
		name=trim(rs("name"))
		initdate = trim(rs("initdate"))
	end if
	rs.Close

	sql = "select * from boo_style_analysis where id='1' "
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
	if not rs.EOF then
		Strtactile=trim(rs("tactile")) 
		Strindividual=trim(rs("individual")) 
		Strvisual=trim(rs("visual"))
		Strauditory=trim(rs("auditory"))
		Strkinesthetic=trim(rs("kinesthetic"))
		Strgroup1=trim(rs("group1"))
	end if
	rs.close



%>
<html>
<head>
<title>【LDCC英外語能力診斷輔導中心】</title>


<LINK rel=stylesheet Type="text/css" href="lib\default.css">


</head>
<body bottomMargin=0 leftMargin=0 topMargin=0 rightMargin=0 marginheight="0" marginwidth="0" >

<BR>
<table width="600" align="center"  cellpadding=0 cellspacing=0>
<tr><td  align="center" class="T2">
學習問卷調查 - 學習風格統計分析
</td></tr>
<tr><td>
<BR>
1.	在每一個空格裡，填入該題的分數(SA=5, A=4, U=3, D=2, SD=1)。<BR>
Carefully transfer your score onto each blank.<BR>
2.	把每一欄的成績加起來，把總計寫在各欄的「Total」的格子裡。<BR>
Add up each column.  Put the result on the line marked TOTAL.<BR>
3.	將「Total」的分數乘以2，填入適當的格子。<BR>
Multiply the total score of each column by 2, and put it in the appropriate blank.<BR><BR>

學號：<input type="text" value="&nbsp;&nbsp;<%=sid%>"  size="20"  class="font1"  readonly>
姓名：<input type="text" value="&nbsp;&nbsp;<%=name%>"  size="20"  class="font1"  readonly>
填寫日期：<input type="text" value="&nbsp;&nbsp;<%=initdate%>"  size="20"  class="font1"  readonly>
</td></tr>
<tr><td  valign="top" align="center">
<BR>
	<table width="100%" cellpadding=3 cellspacing=0 border="1" >
	<tr class="inputlabel"><td width="50%">VISUAL</td><td width="50%">AUDITORY</td></tr>
	<tr>
	<td>&nbsp;6.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q6%>"  size="10"  class="font1"  readonly></td>
	<td>&nbsp;1.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q1%>"  size="10"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;10.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q10%>"  size="10"  class="font1"  readonly></td>
	<td>&nbsp;7.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q7%>"  size="10"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;12.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q12%>"  size="10"  class="font1"  readonly></td>
	<td>&nbsp;9.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q9%>"  size="10"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;24.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q24%>"  size="10"  class="font1"  readonly></td>
	<td>&nbsp;17.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q17%>"  size="10"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;29.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q29%>"  size="10"  class="font1"  readonly></td>
	<td>&nbsp;20.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q20%>"  size="10"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;Total：&nbsp;&nbsp; <input type="text" value="&nbsp;&nbsp;<%=cint(q6)+cint(q10)+cint(q12)+cint(q24)+cint(q29)%>"  size="10"  class="font1"  readonly>X 2 = <input type="text" value="&nbsp;&nbsp;<%=(cint(q6)+cint(q10)+cint(q12)+cint(q24)+cint(q29))*2%>"  size="10"  class="font1"  readonly></td>
	<td>&nbsp;Total：&nbsp;&nbsp;<input type="text" value="&nbsp;&nbsp;<%=cint(q1)+cint(q7)+cint(q9)+cint(q17)+cint(q20)%>"  size="10"  class="font1"  readonly> X 2 = <input type="text" value="&nbsp;&nbsp;<%=(cint(q1)+cint(q7)+cint(q9)+cint(q17)+cint(q20))*2%>"  size="10"  class="font1"  readonly></td>
	</tr>
	<tr class="inputlabel"><td width="50%">KINESTHETIC</td><td width="50%">TACTILE</td></tr>
	<tr>
	<td>&nbsp;2.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q2%>"  size="10"  class="font1"  readonly></td>
	<td>&nbsp;11.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q11%>"  size="10"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;8.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q8%>"  size="10"  class="font1"  readonly></td>
	<td>&nbsp;14.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q14%>"  size="10"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;15.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q15%>"  size="10"  class="font1"  readonly></td>
	<td>&nbsp;16.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q16%>"  size="10"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;19.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q19%>"  size="10"  class="font1"  readonly></td>
	<td>&nbsp;22.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q22%>"  size="10"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;26.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q26%>"  size="10"  class="font1"  readonly></td>
	<td>&nbsp;25.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q25%>"  size="10"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;Total：&nbsp;&nbsp;<input type="text" value="&nbsp;&nbsp;<%=cint(q2)+cint(q8)+cint(q15)+cint(q19)+cint(q26)%>"  size="10"  class="font1"  readonly> X 2 =<input type="text" value="&nbsp;&nbsp;<%=(cint(q2)+cint(q8)+cint(q15)+cint(q19)+cint(q26))*2%>"  size="10"  class="font1"  readonly> </td>
	<td>Total：&nbsp;&nbsp;<input type="text" value="&nbsp;&nbsp;<%=cint(q11)+cint(q14)+cint(q16)+cint(q22)+cint(q25)%>"  size="10"  class="font1"  readonly> X 2 = <input type="text" value="&nbsp;&nbsp;<%=(cint(q11)+cint(q14)+cint(q16)+cint(q22)+cint(q25))*2%>"  size="10"  class="font1"  readonly></td>
	</tr>
	<tr class="inputlabel"><td width="50%">GROUP</td><td width="50%">INDIVIDUAL</td></tr>
	<tr>
	<td>&nbsp;3.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q3%>"  size="10"  class="font1"  readonly></td>
	<td>&nbsp;13.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q13%>"  size="10"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;4.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q4%>"  size="10"  class="font1"  readonly></td>
	<td>&nbsp;18.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q18%>"  size="10"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;5.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q5%>"  size="10"  class="font1"  readonly></td>
	<td>&nbsp;27.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q27%>"  size="10"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;21.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q21%>"  size="10"  class="font1"  readonly></td>
	<td>&nbsp;28.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q28%>"  size="10"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;23.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q23%>"  size="10"  class="font1"  readonly></td>
	<td>&nbsp;30.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q30%>"  size="10"  class="font1"  readonly></td>
	</tr>
	<tr><td>&nbsp;Total：&nbsp;&nbsp;<input type="text" value="&nbsp;&nbsp;<%=cint(q3)+cint(q4)+cint(q5)+cint(q21)+cint(q23)%>"  size="10"  class="font1"  readonly> X 2 = <input type="text" value="&nbsp;&nbsp;<%=(cint(q3)+cint(q4)+cint(q5)+cint(q21)+cint(q23))*2%>"  size="10"  class="font1"  readonly></td>
	<td>&nbsp;Total：&nbsp;&nbsp;<input type="text" value="&nbsp;&nbsp;<%=cint(q13)+cint(q18)+cint(q27)+cint(q28)+cint(q30)%>"  size="10"  class="font1"  readonly> X 2 = <input type="text" value="&nbsp;&nbsp;<%=(cint(q13)+cint(q18)+cint(q27)+cint(q28)+cint(q30))*2%>"  size="10"  class="font1"  readonly></td></tr>
	</table>

</td></tr>
<tr><td class="T2">
<BR><BR>
Major learning Style Preference		38-50<BR>
Minor learning Style Preference	    25-37<BR>
Negligible					     0-24<BR>
</td></tr>
</table>
<BR><BR>
<p style="page-break-after:always"></p>
<BR><BR>
<table width="600" align="center"  cellpadding=0 cellspacing=0>
<tr><td >
	<table  width="100%" align="center"   border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
	<tr><td  width="20%">學習風格傾向</td><td   width="80%">你的學習特徵</td></tr>
	<tr><td >Major Learning style preference主要的學習風格傾向(38-50)</td>
	<td>
		<% 
		for i = 50 to 37 step -1
			if Cstr((cint(q11)+cint(q14)+cint(q16)+cint(q22)+cint(q25))*2) = Cstr(i)  then
		%>
			<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
			<tr><td width="5%" rowspan="2"><%=(cint(q11)+cint(q14)+cint(q16)+cint(q22)+cint(q25))*2%></td><td width="95%" class="T3" align="center">Tactile Major Learning Style Preference <br>觸覺學習者</td></tr>
			<tr><td ><%=replace(Strtactile,vbcrlf,"<br>")%></td></tr>
			</table>
		<%
			end if
			if Cstr((cint(q13)+cint(q18)+cint(q27)+cint(q28)+cint(q30))*2) = Cstr(i)  then
		%>
			<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
			<tr><td width="5%" rowspan="2"><%=(cint(q13)+cint(q18)+cint(q27)+cint(q28)+cint(q30))*2%></td><td width="95%" class="T3" align="center">Individual Major Learning Style Preference <br>個人學習者</td></tr>
			<tr><td ><%=replace(Strindividual,vbcrlf,"<br>")%></td></tr>
			</table>
		<%
			end if
			if Cstr((cint(q6)+cint(q10)+cint(q12)+cint(q24)+cint(q29))*2) = Cstr(i)  then
		%>
			<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
			<tr><td width="5%"  rowspan="2"><%=(cint(q6)+cint(q10)+cint(q12)+cint(q24)+cint(q29))*2%></td><td width="95%" class="T3" align="center">Visual Major Learning Style Preference<br>視覺學習者</td></tr>
			<tr><td ><%=replace(Strvisual,vbcrlf,"<br>")%></td></tr>
			</table>
		<%
			end if
			if Cstr((cint(q1)+cint(q7)+cint(q9)+cint(q17)+cint(q20))*2) = Cstr(i)  then
		%>
			<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
			<tr><td width="5%"  rowspan="2"><%=(cint(q1)+cint(q7)+cint(q9)+cint(q17)+cint(q20))*2%></td><td width="95%"  class="T3" align="center">Auditory Major Learning Style Preference<br>聽覺學習者</td></tr>
			<tr><td width="95%"><%=replace(Strauditory,vbcrlf,"<br>")%></td></tr>
			</table>
		<%
			end if
			if Cstr((cint(q2)+cint(q8)+cint(q15)+cint(q19)+cint(q26))*2) = Cstr(i)  then
		%>
			<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
			<tr><td width="5%" rowspan="2"><%=(cint(q2)+cint(q8)+cint(q15)+cint(q19)+cint(q26))*2%></td><td width="95%" class="T3" align="center">Kinesthetic Major Learning Style Preference<br>動覺學習者</td></tr>
			<tr><td width="95%"><%=replace(Strkinesthetic,vbcrlf,"<br>")%></td></tr>
			</table>
		<%
			end if
			if Cstr((cint(q3)+cint(q4)+cint(q5)+cint(q21)+cint(q23))*2) = Cstr(i)  then
		%>
			<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
			<tr><td width="5%" rowspan="2"><%=(cint(q3)+cint(q4)+cint(q5)+cint(q21)+cint(q23))*2%></td><td width="95%"  class="T3" align="center">Croup Major Learning Style Preference<br>團體學習者</td></tr>
			<tr><td width="95%"><%=replace(Strgroup1,vbcrlf,"<br>")%></td></tr>
			</table>
		<%
			end if
		Next 
		%>
		&nbsp;
	</td>
	</tr>
	<tr><td >Minor Learning style preference次要的學習風格傾向(25-37)</td>
	<td>
	<% 
		for i = 37 to 24 step -1
			if Cstr((cint(q11)+cint(q14)+cint(q16)+cint(q22)+cint(q25))*2) = Cstr(i)  then
		%>
			<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
			<tr><td width="5%" rowspan="2"><%=(cint(q11)+cint(q14)+cint(q16)+cint(q22)+cint(q25))*2%></td><td width="95%" class="T3" align="center">Tactile Major Learning Style Preference <br>觸覺學習者</td></tr>
			<tr><td ><%=replace(Strtactile,vbcrlf,"<br>")%></td></tr>
			</table>
		<%
			end if
			if Cstr((cint(q13)+cint(q18)+cint(q27)+cint(q28)+cint(q30))*2) = Cstr(i)  then
		%>
			<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
			<tr><td width="5%" rowspan="2"><%=(cint(q13)+cint(q18)+cint(q27)+cint(q28)+cint(q30))*2%></td><td width="95%" class="T3" align="center">Individual Major Learning Style Preference <br>個人學習者</td></tr>
			<tr><td ><%=replace(Strindividual,vbcrlf,"<br>")%></td></tr>
			</table>
		<%
			end if
			if Cstr((cint(q6)+cint(q10)+cint(q12)+cint(q24)+cint(q29))*2) = Cstr(i)  then
		%>
			<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
			<tr><td width="5%"  rowspan="2"><%=(cint(q6)+cint(q10)+cint(q12)+cint(q24)+cint(q29))*2%></td><td width="95%" class="T3" align="center">Visual Major Learning Style Preference<br>視覺學習者</td></tr>
			<tr><td ><%=replace(Strvisual,vbcrlf,"<br>")%></td></tr>
			</table>
		<%
			end if
			if Cstr((cint(q1)+cint(q7)+cint(q9)+cint(q17)+cint(q20))*2) = Cstr(i)  then
		%>
			<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
			<tr><td width="5%"  rowspan="2"><%=(cint(q1)+cint(q7)+cint(q9)+cint(q17)+cint(q20))*2%></td><td width="95%"  class="T3" align="center">Auditory Major Learning Style Preference<br>聽覺學習者</td></tr>
			<tr><td width="95%"><%=replace(Strauditory,vbcrlf,"<br>")%></td></tr>
			</table>
		<%
			end if
			if Cstr((cint(q2)+cint(q8)+cint(q15)+cint(q19)+cint(q26))*2) = Cstr(i)  then
		%>
			<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
			<tr><td width="5%" rowspan="2"><%=(cint(q2)+cint(q8)+cint(q15)+cint(q19)+cint(q26))*2%></td><td width="95%" class="T3" align="center">Kinesthetic Major Learning Style Preference<br>動覺學習者</td></tr>
			<tr><td width="95%"><%=replace(Strkinesthetic,vbcrlf,"<br>")%></td></tr>
			</table>
		<%
			end if
			if Cstr((cint(q3)+cint(q4)+cint(q5)+cint(q21)+cint(q23))*2) = Cstr(i)  then
		%>
			<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
			<tr><td width="5%" rowspan="2"><%=(cint(q3)+cint(q4)+cint(q5)+cint(q21)+cint(q23))*2%></td><td width="95%"  class="T3" align="center">Croup Major Learning Style Preference<br>團體學習者</td></tr>
			<tr><td width="95%"><%=replace(Strgroup1,vbcrlf,"<br>")%></td></tr>
			</table>
		<%
			end if

		Next 
		%>
		&nbsp;
	</td>
	</tr>
	<tr><td >Negligible Learning style preference比較被忽略的學習風格傾向(0-24)</td>
	<td>
	<% 
		for i = 24 to 0 step -1
			if Cstr((cint(q11)+cint(q14)+cint(q16)+cint(q22)+cint(q25))*2) = Cstr(i)  then
		%>
			<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
			<tr><td width="5%" rowspan="2"><%=(cint(q11)+cint(q14)+cint(q16)+cint(q22)+cint(q25))*2%></td><td width="95%" class="T3" align="center">Tactile Major Learning Style Preference <br>觸覺學習者</td></tr>
			<tr><td ><%=replace(Strtactile,vbcrlf,"<br>")%></td></tr>
			</table>
		<%
			end if
			if Cstr((cint(q13)+cint(q18)+cint(q27)+cint(q28)+cint(q30))*2) = Cstr(i)  then
		%>
			<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
			<tr><td width="5%" rowspan="2"><%=(cint(q13)+cint(q18)+cint(q27)+cint(q28)+cint(q30))*2%></td><td width="95%" class="T3" align="center">Individual Major Learning Style Preference <br>個人學習者</td></tr>
			<tr><td ><%=replace(Strindividual,vbcrlf,"<br>")%></td></tr>
			</table>
		<%
			end if
			if Cstr((cint(q6)+cint(q10)+cint(q12)+cint(q24)+cint(q29))*2) = Cstr(i)  then
		%>
			<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
			<tr><td width="5%"  rowspan="2"><%=(cint(q6)+cint(q10)+cint(q12)+cint(q24)+cint(q29))*2%></td><td width="95%" class="T3" align="center">Visual Major Learning Style Preference<br>視覺學習者</td></tr>
			<tr><td ><%=replace(Strvisual,vbcrlf,"<br>")%></td></tr>
			</table>
		<%
			end if
			if Cstr((cint(q1)+cint(q7)+cint(q9)+cint(q17)+cint(q20))*2) = Cstr(i)  then
		%>
			<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
			<tr><td width="5%"  rowspan="2"><%=(cint(q1)+cint(q7)+cint(q9)+cint(q17)+cint(q20))*2%></td><td width="95%"  class="T3" align="center">Auditory Major Learning Style Preference<br>聽覺學習者</td></tr>
			<tr><td width="95%"><%=replace(Strauditory,vbcrlf,"<br>")%></td></tr>
			</table>
		<%
			end if
			if Cstr((cint(q2)+cint(q8)+cint(q15)+cint(q19)+cint(q26))*2) = Cstr(i)  then
		%>
			<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
			<tr><td width="5%" rowspan="2"><%=(cint(q2)+cint(q8)+cint(q15)+cint(q19)+cint(q26))*2%></td><td width="95%" class="T3" align="center">Kinesthetic Major Learning Style Preference<br>動覺學習者</td></tr>
			<tr><td width="95%"><%=replace(Strkinesthetic,vbcrlf,"<br>")%></td></tr>
			</table>
		<%
			end if
			if Cstr((cint(q3)+cint(q4)+cint(q5)+cint(q21)+cint(q23))*2) = Cstr(i)  then
		%>
			<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
			<tr><td width="5%" rowspan="2"><%=(cint(q3)+cint(q4)+cint(q5)+cint(q21)+cint(q23))*2%></td><td width="95%"  class="T3" align="center">Croup Major Learning Style Preference<br>團體學習者</td></tr>
			<tr><td width="95%"><%=replace(Strgroup1,vbcrlf,"<br>")%></td></tr>
			</table>
		<%
			end if

		Next 
		%>
		&nbsp;
	</td>
	</tr>
	</table>
</td>
</tr>
<tr>
<td class="T5" align="center">
<br><br>
～如欲了解如何運用自己的學習風格、<br>
加強使用學習策略來提升英文能力，<br>
英外語能力診斷輔導中心歡迎您～
</td>
</tr>
</table>
<BR><BR>

</body>
</html>
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->
<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 50 %>
<% Response.CacheControl = "No-cache" %>

<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<%
sid = trim(request("sid"))
qstid= trim(request("qstid"))

set rs = server.CreateObject("adodb.recordset")

	if sid<>"" then
		sql = "select a.*,b.name from boo_questionnaire_strategy  a left join boo_profile b on a.sid=b.sid where a.sid = '"&sid&"' and a.yn='Y' order by a.initdate desc"
	else
		sql = "select a.*,b.name from boo_questionnaire_strategy  a left join boo_profile b on a.sid=b.sid where a.qstid = '"&qstid&"' and a.yn='Y' order by a.initdate desc"
	end if
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
	if rs.EOF then

	else
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
		q31 = trim(rs("q31"))
		q32 = trim(rs("q32"))
		q33 = trim(rs("q33"))
		q34 = trim(rs("q34"))
		q35 = trim(rs("q35"))
		q36 = trim(rs("q36"))
		q37 = trim(rs("q37"))
		q38 = trim(rs("q38"))
		q39 = trim(rs("q39"))
		q40 = trim(rs("q40"))
		q41 = trim(rs("q41"))
		q42 = trim(rs("q42"))
		q43 = trim(rs("q43"))
		q44 = trim(rs("q44"))
		q45 = trim(rs("q45"))
		q46 = trim(rs("q46"))
		q47 = trim(rs("q47"))
		q48 = trim(rs("q48"))
		q49 = trim(rs("q49"))
		q50 = trim(rs("q50"))
		name=trim(rs("name"))
		initdate = trim(rs("initdate"))
	end if
	rs.close	

partA = cint(q1)+cint(q2)+cint(q3)+cint(q4)+cint(q5)+cint(q6)+cint(q7)+cint(q8)+cint(q9)
partB = cint(q10)+cint(q11)+cint(q12)+cint(q13)+cint(q14)+cint(q15)+cint(q16)+cint(q17)+cint(q18)+cint(q19)+cint(q20)+cint(q21)+cint(q22)+cint(q23)
partC = cint(q24)+cint(q25)+cint(q26)+cint(q27)+cint(q28)+cint(q29)
partD = cint(q30)+cint(q31)+cint(q32)+cint(q33)+cint(q34)+cint(q35)+cint(q36)+cint(q37)+cint(q38)
partE = cint(q39)+cint(q40)+cint(q41)+cint(q42)+cint(q43)+cint(q44)
partF = cint(q45)+cint(q46)+cint(q47)+cint(q48)+cint(q49)+cint(q50)
Total = cint(partA) + cint(partB) + cint(partC) + cint(partD) + cint(partE) + cint(partF)

sql = "select * from boo_strategy_analysis where id='1' "
rs.Open sql,msconn,adOpenStatic,adLockReadonly
if not rs.EOF then
	StrA=trim(rs("A")) 
	StrB=trim(rs("B")) 
	StrC=trim(rs("C"))
	StrD=trim(rs("D"))
	StrE=trim(rs("E"))
	StrF=trim(rs("F"))
end if
rs.close

%>


<html>
<head>
<title>【LDCC英外語能力診斷輔導中心】</title>

<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">


</head>
<body bottomMargin=0 leftMargin=0 topMargin=0 rightMargin=0 marginheight="0" marginwidth="0" >

<table width="650" align="center"  cellpadding=0 cellspacing=0>
<tr><td  align="center" class="T2">
Worksheet for Answering and Scoring of SILL<BR>
答案分數工作單

</td></tr>
<tr><td>
<BR><BR>
1.	Write your response to each item (that is, write 1, 2, 3, 4, 5) in each of the blanks.<BR>
2.	Add up each column.  Put the result on the line marked SUM.<BR>
3.	Divide by the number under SUM to get the average for each column.  Round this average off to the nearest tenth, as in 3.4.<BR>
4.	Figure out your overall average.  To do this, add up all the SUMS for the different parts of the SILL.  Then divide by 50.<BR>
5.	Copy your averages (for each part and for the whole SILL) from the Worksheet to the Profile.<BR>
1.	在每一個空格裡，寫下你的回答(1, 2, 3, 4, 5)。<BR>
2.	把每一欄的成績加起來，把總計寫在「加總」的格子裡。<BR>
3.	再把總和除以每一欄底下的數字求得平均數，算至小數點第一位，例如3.4。<BR>
4.	為求總平均數，將每一項總和加起來除以50。<BR>
5.	把自己工作單上面的平均數抄到剖析表裡。<BR><BR><BR>

學號：<input type="text" value="&nbsp;&nbsp;<%=sid%>"  size="20"  class="font1"  readonly>
姓名：<input type="text" value="&nbsp;&nbsp;<%=name%>"  size="20"  class="font1"  readonly>
填寫日期：<input type="text" value="&nbsp;&nbsp;<%=initdate%>"  size="20"  class="font1"  readonly>
<BR><BR>
</td></tr>
<tr><td  valign="top" align="center">
	<table width="100%" cellpadding=2 cellspacing=0 border="1" >
	<tr class="inputlabel"><td>Part A</td><td>Part B</td><td>Part C</td><td>Part D</td><td>Part E</td><td>Part F</td><td>Whole SILL</td></tr>
	<tr>
	<td>&nbsp;1.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q1%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;10.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q10%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;24.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q24%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;30.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q30%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;39.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q39%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;45.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q45%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;SUM Part A&nbsp;<input type="text" value="&nbsp;&nbsp;<%=partA%>"  size="3"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;2.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q2%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;11.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q11%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;25.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q25%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;31.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q31%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;40.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q40%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;46.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q46%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;SUM Part B&nbsp;<input type="text" value="&nbsp;&nbsp;<%=partB%>"  size="3"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;3.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q3%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;12.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q12%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;26.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q26%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;32.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q32%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;41.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q41%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;47.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q47%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;SUM Part C&nbsp;<input type="text" value="&nbsp;&nbsp;<%=partC%>"  size="3"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;4.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q4%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;13.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q13%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;27.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q27%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;33.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q33%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;42.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q42%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;48.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q48%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;SUM Part D&nbsp;<input type="text" value="&nbsp;&nbsp;<%=partD%>"  size="3"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;5.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q5%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;14.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q14%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;28.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q28%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;34.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q34%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;43.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q43%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;49.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q49%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;SUM Part E&nbsp;<input type="text" value="&nbsp;&nbsp;<%=partE%>"  size="3"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;6.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q6%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;15.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q15%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;29.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q29%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;35.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q35%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;44.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q44%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;50.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q50%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;SUM Part F&nbsp;<input type="text" value="&nbsp;&nbsp;<%=partF%>"  size="3"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;7.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q7%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;16.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q16%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;36.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q36%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;</td>
	</tr>
	<tr>
	<td>&nbsp;8.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q8%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;17.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q17%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;37.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q37%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;</td>
	</tr>
	<tr>
	<td>&nbsp;9.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q9%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;18.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q18%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;38.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q38%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;</td>
	</tr>
	<tr>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;19.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q19%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;</td>
	</tr>
	<tr>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;20.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q20%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;</td>
	</tr>
	<tr>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;21.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q21%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;</td>
	</tr>
	<tr>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;22.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q22%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;</td>
	</tr>
	<tr>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;23.&nbsp;<input type="text" value="&nbsp;&nbsp;<%=q23%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;&nbsp;</td>
	<td>&nbsp;</td>
	</tr>
	<tr>
	<td>&nbsp;SUM&nbsp;<input type="text" value="&nbsp;&nbsp;<%=partA%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;SUM&nbsp;<input type="text" value="&nbsp;&nbsp;<%=partB%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;SUM&nbsp;<input type="text" value="&nbsp;&nbsp;<%=partC%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;SUM&nbsp;<input type="text" value="&nbsp;&nbsp;<%=partD%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;SUM&nbsp;<input type="text" value="&nbsp;&nbsp;<%=partE%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;SUM&nbsp;<input type="text" value="&nbsp;&nbsp;<%=partF%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;SUM&nbsp;<input type="text" value="&nbsp;&nbsp;<%=Total%>"  size="5"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td>&nbsp;÷&nbsp;9&nbsp;<input type="text" value="&nbsp;&nbsp;<%=round(cdbl(partA)/cdbl(9),1)%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;÷&nbsp;14&nbsp;<input type="text" value="&nbsp;&nbsp;<%=round(cdbl(partB)/cdbl(14),1)%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;÷&nbsp;6&nbsp;<input type="text" value="&nbsp;&nbsp;<%=round(cdbl(partC)/cdbl(6),1)%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;÷&nbsp;9&nbsp;<input type="text" value="&nbsp;&nbsp;<%=round(cdbl(partD)/cdbl(9),1)%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;÷&nbsp;6&nbsp;<input type="text" value="&nbsp;&nbsp;<%=round(cdbl(partE)/cdbl(6),1)%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;÷&nbsp;6&nbsp;<input type="text" value="&nbsp;&nbsp;<%=round(cdbl(partF)/cdbl(6),1)%>"  size="5"  class="font1"  readonly></td>
	<td>&nbsp;÷&nbsp;50&nbsp;<input type="text" value="&nbsp;&nbsp;<%=round(cdbl(Total)/cdbl(50),1)%>"  size="5"  class="font1"  readonly></td>
	</tr>
	</table>

</td></tr>

</table>
<BR><BR>
<p style="page-break-after:always"></p>
<BR><BR>
<table width="650" align="center"  cellpadding=0 cellspacing=0>
<tr><td  align="center" class="T2">
Profile of Results on the SILL<BR>
結果剖析表
</td></tr>
<tr><td>
<BR><BR>
This Profile will show your SILL results.  These results will tell you the kinds of strategies you use in learning English.  There are no right or wrong answers.  To complete this profile, transfer your averages for each part of the SILL, and your overall average for the whole SILL.<BR>
這張剖析表將顯示你的語言學習策略查核表所得結果。這些結果讓你得知自己在學習英語時所使用的學習策略，而這並沒有對或錯的答案。將你工作單上的各項平均數及總平均數抄寫至剖析表上完成作業。
<BR><BR>
學號：<input type="text" value="&nbsp;&nbsp;<%=sid%>"  size="20"  class="font1"  readonly>
姓名：<input type="text" value="&nbsp;&nbsp;<%=name%>"  size="20"  class="font1"  readonly>
填寫日期：<input type="text" value="&nbsp;&nbsp;<%=initdate%>"  size="20"  class="font1"  readonly>
<BR><BR>
</td></tr>
<tr><td  valign="top" align="center">
	<table width="100%" cellpadding=2 cellspacing=0 border="0" >
	<tr class="inputlabel"><td width="5%">Part</td><td width="60%">What Strategies Are Covered</td><td width="35%">Your Average on This Part</td></tr>
	<tr>
	<td align="center">A</td>
	<td>Remembering more effectively</td>
	<td align="center"><input type="text" value="&nbsp;&nbsp;<%=round(cdbl(partA)/cdbl(9),1)%>"  size="5"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td align="center">B</td>
	<td>Using all your mental processes</td>
	<td align="center"><input type="text" value="&nbsp;&nbsp;<%=round(cdbl(partB)/cdbl(14),1)%>"  size="5"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td align="center">C</td>
	<td>Compensating for missing knowledge</td>
	<td align="center"><input type="text" value="&nbsp;&nbsp;<%=round(cdbl(partC)/cdbl(6),1)%>"  size="5"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td align="center">D</td>
	<td>Organizing and evaluating your learning</td>
	<td align="center"><input type="text" value="&nbsp;&nbsp;<%=round(cdbl(partD)/cdbl(9),1)%>"  size="5"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td align="center">E</td>
	<td>Managing your emotions</td>
	<td align="center"><input type="text" value="&nbsp;&nbsp;<%=round(cdbl(partE)/cdbl(6),1)%>"  size="5"  class="font1"  readonly></td>
	</tr>
	<tr>
	<td align="center">F</td>
	<td>Learning with others</td>
	<td align="center"><input type="text" value="&nbsp;&nbsp;<%=round(cdbl(partF)/cdbl(6),1)%>"  size="5"  class="font1"  readonly></td>
	</tr>
	</table>
</td></tr>
<tr><td  align="center" class="T2">
OVERALL AVERAGE
</td></tr>
<tr><td  align="center" >
	<B>Key to Understanding Your Averages</B>
	<table width="90%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
	<tr><td>Hight</td><td>Always or almost always used <br>Usually used</td><td>4.5 to 5.0<BR>3.5 to 4.4</td></tr>
	<tr><td>Medium</td><td>Sometimes used</td><td>2.5 to 3.4</td></tr>
	<tr><td>Low</td><td>Generally not used<br>Never or almost never used</td><td>1.5 to 2.4<BR>1.0 to 1.4</td></tr>
	</table><BR><BR>
	<B>Graph Your Averages Here</B>
	<BR><BR>
	 <table cellpadding="0" cellspacing="0" border=0 height="186" width="420">
	 <tr height="5" align="center">
	 <td><%=round(cdbl(partA)/cdbl(9),1)%></td>
	 <td><%=round(cdbl(partB)/cdbl(14),1)%></td>
	 <td><%=round(cdbl(partC)/cdbl(6),1)%></td>
	 <td><%=round(cdbl(partD)/cdbl(9),1)%></td>
	 <td><%=round(cdbl(partE)/cdbl(6),1)%></td>
	 <td><%=round(cdbl(partF)/cdbl(6),1)%></td>
	 <td><%=round(cdbl(Total)/cdbl(50),1)%></td>
	 </tr>
	 <tr valign="bottom" align="center"> 
	 <td width="60"><img src="images/attitude_bar.gif" width="20" height="<%=round(cdbl(partA)/cdbl(9),1)*20%>" ></td>
	 <td width="60"><img src="images/attitude_bar.gif" width="20" height="<%=round(cdbl(partB)/cdbl(14),1)*20%>" ></td>
	 <td width="60"><img src="images/attitude_bar.gif" width="20" height="<%=round(cdbl(partC)/cdbl(6),1)*20%>" ></td>
	 <td width="60"><img src="images/attitude_bar.gif" width="20" height="<%=round(cdbl(partD)/cdbl(9),1)*20%>" ></td>
	 <td width="60"><img src="images/attitude_bar.gif" width="20" height="<%=round(cdbl(partE)/cdbl(6),1)*20%>" ></td>
	 <td width="60"><img src="images/attitude_bar.gif" width="20" height="<%=round(cdbl(partF)/cdbl(6),1)*20%>" ></td>
	 <td width="60"><img src="images/attitude_bar.gif" width="20" height="<%=round(cdbl(Total)/cdbl(50),1)*20%>" ></td>
	 </tr>
	 <tr height="5" align="center"><td>A</td><td>B</td><td>C</td><td>D</td><td>E</td><td>F</td><td>overall<br>Average</td></tr>
	 </table>

</td></tr>
</table>
<BR><BR>
<BR><BR>
<p style="page-break-after:always"></p>
<BR><BR>

<table width="650" align="center"  cellpadding=0 cellspacing=0>
<tr><td >
您一定很想知道每一個部分代表的是什麼學習策略，讓我們一起往下看，找出平常使用的英／外語學習策略是什麼吧！！
</td></tr>
<tr><td>

	<table width="98%"  border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
	<tr><td  width="15%">使用頻率</td><td   width="20%">使用頻率</td><td   width="65%">DIRECT STRATEGIES 直接策略</td></tr>
	<tr><td rowspan="2">Hight</td><td>Always or almost always used<br> 隨時使用<br>(4.5 to 5.0)</td>
	<td>
	<% 
		if round(cdbl(partA)/cdbl(9),1) >= 4.5 and round(cdbl(partA)/cdbl(9),1) <= 5.0 then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partA)/cdbl(9),1)%></td><td width="5%">A部份</td><td width="20%">Memory strategies 記憶策略</td><td width="70%"><%=replace(StrA,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
		if round(cdbl(partB)/cdbl(14),1) >= 4.5 and round(cdbl(partB)/cdbl(14),1) <= 5.0 then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partB)/cdbl(14),1) %></td><td width="5%">B部份</td><td width="20%">Cognitive strategies 認知策略</td><td width="70%"><%=replace(StrB,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	
	%>
	&nbsp;
	</td>
	</tr>
	<tr><td>Usually used<br>蠻常使用<br>(3.5 to 4.4)</td>
	<td>
	<% 
		if round(cdbl(partA)/cdbl(9),1) >= 3.5 and round(cdbl(partA)/cdbl(9),1) <= 4.4 then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partA)/cdbl(9),1)%></td><td width="5%">A部份</td><td width="20%">Memory strategies 記憶策略</td><td width="70%"><%=replace(StrA,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
		if round(cdbl(partB)/cdbl(14),1) >= 3.5 and round(cdbl(partB)/cdbl(14),1) <= 4.4 then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partB)/cdbl(14),1) %></td><td width="5%">B部份</td><td width="20%">Cognitive strategies 認知策略</td><td width="70%"><%=replace(StrB,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	
	%>
	&nbsp;
	</td>
	</tr>
	<tr><td>Medium</td><td>Sometimes used<br>偶而使用<br>(2.5 to 3.4)</td>
	<td>
	<% 
		if round(cdbl(partA)/cdbl(9),1) >= 2.5 and round(cdbl(partA)/cdbl(9),1) <= 3.4 then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partA)/cdbl(9),1)%></td><td width="5%">A部份</td><td width="20%">Memory strategies 記憶策略</td><td width="70%"><%=replace(StrA,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
		if round(cdbl(partB)/cdbl(14),1) >= 2.5 and round(cdbl(partB)/cdbl(14),1) <= 3.4 then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partB)/cdbl(14),1) %></td><td width="5%">B部份</td><td width="20%">Cognitive strategies 認知策略</td><td width="70%"><%=replace(StrB,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	
	%>
	&nbsp;
	</td>
	</tr>
	<tr><td  rowspan="2">Low</td><td>Generally not used<br>一般說來不太使用<br>(1.5 to 2.4)</td>
	<td>
	<% 
		if round(cdbl(partA)/cdbl(9),1) = 2.4  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partA)/cdbl(9),1)%></td><td width="5%">A部份</td><td width="20%">Memory strategies 記憶策略</td><td width="70%"><%=replace(StrA,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
		if round(cdbl(partB)/cdbl(14),1) = 2.4 then
	%>
		<table width="98%" height="100%" border=1 cellpadding=1 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partB)/cdbl(14),1) %></td><td width="5%">B部份</td><td width="20%">Cognitive strategies 認知策略</td><td width="70%"><%=replace(StrB,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partC)/cdbl(6),1) = 2.4  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partC)/cdbl(6),1)%></td><td width="5%">C部份</td><td width="20%">Compensati on strategies 補償策略</td><td width="70%"><%=replace(StrC,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partD)/cdbl(9),1) = 2.4  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partD)/cdbl(9),1)%></td><td width="5%">D部份</td><td width="20%">Meta-cognitive Strategies 後設認知策略</td><td width="70%"><%=replace(StrD,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partE)/cdbl(6),1) = 2.4  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partE)/cdbl(6),1)%></td><td width="5%">E部份</td><td width="20%">Affective strategies 情意策略</td><td width="70%"><%=replace(StrE,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partF)/cdbl(6),1) = 2.4  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partF)/cdbl(6),1)%></td><td width="5%">F部份</td><td width="20%">Social strategies 社交策略</td><td width="70%"><%=replace(StrF,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partA)/cdbl(9),1) = 2.3  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partA)/cdbl(9),1)%></td><td width="5%">A部份</td><td width="20%">Memory strategies 記憶策略</td><td width="70%"><%=replace(StrA,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
		if round(cdbl(partB)/cdbl(14),1) = 2.3 then
	%>
		<table width="98%" height="100%" border=1 cellpadding=1 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partB)/cdbl(14),1) %></td><td width="5%">B部份</td><td width="20%">Cognitive strategies 認知策略</td><td width="70%"><%=replace(StrB,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partC)/cdbl(6),1) = 2.3  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partC)/cdbl(6),1)%></td><td width="5%">C部份</td><td width="20%">Compensati on strategies 補償策略</td><td width="70%"><%=replace(StrC,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partD)/cdbl(9),1) = 2.3  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partD)/cdbl(9),1)%></td><td width="5%">D部份</td><td width="20%">Meta-cognitive Strategies 後設認知策略</td><td width="70%"><%=replace(StrD,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partE)/cdbl(6),1) = 2.3  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partE)/cdbl(6),1)%></td><td width="5%">E部份</td><td width="20%">Affective strategies 情意策略</td><td width="70%"><%=replace(StrE,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partF)/cdbl(6),1) = 2.3  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partF)/cdbl(6),1)%></td><td width="5%">F部份</td><td width="20%">Social strategies 社交策略</td><td width="70%"><%=replace(StrF,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partA)/cdbl(9),1) = 2.2  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partA)/cdbl(9),1)%></td><td width="5%">A部份</td><td width="20%">Memory strategies 記憶策略</td><td width="70%"><%=replace(StrA,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
		if round(cdbl(partB)/cdbl(14),1) = 2.2 then
	%>
		<table width="98%" height="100%" border=1 cellpadding=1 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partB)/cdbl(14),1) %></td><td width="5%">B部份</td><td width="20%">Cognitive strategies 認知策略</td><td width="70%"><%=replace(StrB,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partC)/cdbl(6),1) = 2.2  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partC)/cdbl(6),1)%></td><td width="5%">C部份</td><td width="20%">Compensati on strategies 補償策略</td><td width="70%"><%=replace(StrC,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partD)/cdbl(9),1) = 2.2  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partD)/cdbl(9),1)%></td><td width="5%">D部份</td><td width="20%">Meta-cognitive Strategies 後設認知策略</td><td width="70%"><%=replace(StrD,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partE)/cdbl(6),1) = 2.2  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partE)/cdbl(6),1)%></td><td width="5%">E部份</td><td width="20%">Affective strategies 情意策略</td><td width="70%"><%=replace(StrE,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partF)/cdbl(6),1) = 2.2  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partF)/cdbl(6),1)%></td><td width="5%">F部份</td><td width="20%">Social strategies 社交策略</td><td width="70%"><%=replace(StrF,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partA)/cdbl(9),1) = 2.1  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partA)/cdbl(9),1)%></td><td width="5%">A部份</td><td width="20%">Memory strategies 記憶策略</td><td width="70%"><%=replace(StrA,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
		if round(cdbl(partB)/cdbl(14),1) = 2.1 then
	%>
		<table width="98%" height="100%" border=1 cellpadding=1 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partB)/cdbl(14),1) %></td><td width="5%">B部份</td><td width="20%">Cognitive strategies 認知策略</td><td width="70%"><%=replace(StrB,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partC)/cdbl(6),1) = 2.1  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partC)/cdbl(6),1)%></td><td width="5%">C部份</td><td width="20%">Compensati on strategies 補償策略</td><td width="70%"><%=replace(StrC,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partD)/cdbl(9),1) = 2.1  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partD)/cdbl(9),1)%></td><td width="5%">D部份</td><td width="20%">Meta-cognitive Strategies 後設認知策略</td><td width="70%"><%=replace(StrD,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partE)/cdbl(6),1) = 2.1  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partE)/cdbl(6),1)%></td><td width="5%">E部份</td><td width="20%">Affective strategies 情意策略</td><td width="70%"><%=replace(StrE,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partF)/cdbl(6),1) = 2.1  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partF)/cdbl(6),1)%></td><td width="5%">F部份</td><td width="20%">Social strategies 社交策略</td><td width="70%"><%=replace(StrF,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partA)/cdbl(9),1) = 2.0  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partA)/cdbl(9),1)%></td><td width="5%">A部份</td><td width="20%">Memory strategies 記憶策略</td><td width="70%"><%=replace(StrA,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
		if round(cdbl(partB)/cdbl(14),1) = 2.0 then
	%>
		<table width="98%" height="100%" border=1 cellpadding=1 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partB)/cdbl(14),1) %></td><td width="5%">B部份</td><td width="20%">Cognitive strategies 認知策略</td><td width="70%"><%=replace(StrB,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partC)/cdbl(6),1) = 2.0  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partC)/cdbl(6),1)%></td><td width="5%">C部份</td><td width="20%">Compensati on strategies 補償策略</td><td width="70%"><%=replace(StrC,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partD)/cdbl(9),1) = 2.0  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partD)/cdbl(9),1)%></td><td width="5%">D部份</td><td width="20%">Meta-cognitive Strategies 後設認知策略</td><td width="70%"><%=replace(StrD,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partE)/cdbl(6),1) = 2.0  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partE)/cdbl(6),1)%></td><td width="5%">E部份</td><td width="20%">Affective strategies 情意策略</td><td width="70%"><%=replace(StrE,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partF)/cdbl(6),1) = 2.0  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partF)/cdbl(6),1)%></td><td width="5%">F部份</td><td width="20%">Social strategies 社交策略</td><td width="70%"><%=replace(StrF,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partA)/cdbl(9),1) = 1.9  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partA)/cdbl(9),1)%></td><td width="5%">A部份</td><td width="20%">Memory strategies 記憶策略</td><td width="70%"><%=replace(StrA,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
		if round(cdbl(partB)/cdbl(14),1) = 1.9 then
	%>
		<table width="98%" height="100%" border=1 cellpadding=1 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partB)/cdbl(14),1) %></td><td width="5%">B部份</td><td width="20%">Cognitive strategies 認知策略</td><td width="70%"><%=replace(StrB,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partC)/cdbl(6),1) = 1.9  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partC)/cdbl(6),1)%></td><td width="5%">C部份</td><td width="20%">Compensati on strategies 補償策略</td><td width="70%"><%=replace(StrC,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partD)/cdbl(9),1) = 1.9  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partD)/cdbl(9),1)%></td><td width="5%">D部份</td><td width="20%">Meta-cognitive Strategies 後設認知策略</td><td width="70%"><%=replace(StrD,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partE)/cdbl(6),1) = 1.9  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partE)/cdbl(6),1)%></td><td width="5%">E部份</td><td width="20%">Affective strategies 情意策略</td><td width="70%"><%=replace(StrE,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partF)/cdbl(6),1) = 1.9  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partF)/cdbl(6),1)%></td><td width="5%">F部份</td><td width="20%">Social strategies 社交策略</td><td width="70%"><%=replace(StrF,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partA)/cdbl(9),1) = 1.8  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partA)/cdbl(9),1)%></td><td width="5%">A部份</td><td width="20%">Memory strategies 記憶策略</td><td width="70%"><%=replace(StrA,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
		if round(cdbl(partB)/cdbl(14),1) = 1.8 then
	%>
		<table width="98%" height="100%" border=1 cellpadding=1 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partB)/cdbl(14),1) %></td><td width="5%">B部份</td><td width="20%">Cognitive strategies 認知策略</td><td width="70%"><%=replace(StrB,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partC)/cdbl(6),1) = 1.8  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partC)/cdbl(6),1)%></td><td width="5%">C部份</td><td width="20%">Compensati on strategies 補償策略</td><td width="70%"><%=replace(StrC,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partD)/cdbl(9),1) = 1.8  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partD)/cdbl(9),1)%></td><td width="5%">D部份</td><td width="20%">Meta-cognitive Strategies 後設認知策略</td><td width="70%"><%=replace(StrD,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partE)/cdbl(6),1) = 1.8  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partE)/cdbl(6),1)%></td><td width="5%">E部份</td><td width="20%">Affective strategies 情意策略</td><td width="70%"><%=replace(StrE,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partF)/cdbl(6),1) = 1.8  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partF)/cdbl(6),1)%></td><td width="5%">F部份</td><td width="20%">Social strategies 社交策略</td><td width="70%"><%=replace(StrF,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partA)/cdbl(9),1) = 1.7  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partA)/cdbl(9),1)%></td><td width="5%">A部份</td><td width="20%">Memory strategies 記憶策略</td><td width="70%"><%=replace(StrA,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
		if round(cdbl(partB)/cdbl(14),1) = 1.7 then
	%>
		<table width="98%" height="100%" border=1 cellpadding=1 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partB)/cdbl(14),1) %></td><td width="5%">B部份</td><td width="20%">Cognitive strategies 認知策略</td><td width="70%"><%=replace(StrB,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partC)/cdbl(6),1) = 1.7  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partC)/cdbl(6),1)%></td><td width="5%">C部份</td><td width="20%">Compensati on strategies 補償策略</td><td width="70%"><%=replace(StrC,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partA)/cdbl(9),1) = 1.6  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partA)/cdbl(9),1)%></td><td width="5%">A部份</td><td width="20%">Memory strategies 記憶策略</td><td width="70%"><%=replace(StrA,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
		if round(cdbl(partB)/cdbl(14),1) = 1.6 then
	%>
		<table width="98%" height="100%" border=1 cellpadding=1 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partB)/cdbl(14),1) %></td><td width="5%">B部份</td><td width="20%">Cognitive strategies 認知策略</td><td width="70%"><%=replace(StrB,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partC)/cdbl(6),1) = 1.6  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partC)/cdbl(6),1)%></td><td width="5%">C部份</td><td width="20%">Compensati on strategies 補償策略</td><td width="70%"><%=replace(StrC,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partA)/cdbl(9),1) = 1.5  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partA)/cdbl(9),1)%></td><td width="5%">A部份</td><td width="20%">Memory strategies 記憶策略</td><td width="70%"><%=replace(StrA,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
		if round(cdbl(partB)/cdbl(14),1) = 1.5 then
	%>
		<table width="98%" height="100%" border=1 cellpadding=1 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partB)/cdbl(14),1) %></td><td width="5%">B部份</td><td width="20%">Cognitive strategies 認知策略</td><td width="70%"><%=replace(StrB,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	<% 
		if round(cdbl(partC)/cdbl(6),1) = 1.5  then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partC)/cdbl(6),1)%></td><td width="5%">C部份</td><td width="20%">Compensati on strategies 補償策略</td><td width="70%"><%=replace(StrC,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	%>
	</td>
	</tr>
	<tr><td>Never or almost never used<br>從來不使用<BR>(1.0 to 1.4)</td>
	<td>
	<% 
		if round(cdbl(partA)/cdbl(9),1) >= 0.0 and round(cdbl(partA)/cdbl(9),1) <= 1.4 then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partA)/cdbl(9),1)%></td><td width="5%">A部份</td><td width="20%">Memory strategies 記憶策略</td><td width="70%"><%=replace(StrA,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
		if round(cdbl(partB)/cdbl(14),1) >= 0.0 and round(cdbl(partB)/cdbl(14),1) <= 1.4 then
	%>
		<table width="98%" height="100%" border=1 cellpadding=2 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td width="5%"><%=round(cdbl(partB)/cdbl(14),1) %></td><td width="5%">B部份</td><td width="20%">Cognitive strategies 認知策略</td><td width="70%"><%=replace(StrB,vbcrlf,"<br>")%><td></tr>
		</table>
	<%
		end if
	
	%>
	&nbsp;
	</td>
	</tr>
	</table>

</td></tr>
</table>

<BR><BR>
<BR><BR>
</body>
</html>
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->
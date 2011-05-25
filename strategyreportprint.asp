<%@Language=VBScript LCID=1033%>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->

<%

sRegDate=trim(request("sRegDate"))
eRegDate=trim(request("eRegDate"))
sIniDate=trim(request("sIniDate"))
eIniDate=trim(request("eIniDate"))
slevel=trim(request("slevel"))
department=trim(request("department"))
yn=trim(request("yn"))



set rs=server.CreateObject("adodb.recordset")

sql = "select a.*,b.name,b.slevel ,b.department,b.grade,b.class1  from boo_questionnaire_strategy a "
sql = sql & " left join boo_profile b on a.sid=b.sid "
sql = sql & " where   1=1  "
	
if sRegDate<>"" and eRegDate="" then
	sql = sql & " and b.initdate >= '" & NumberToDateFormat(sRegDate) & "'  "
end if
if sRegDate="" and eRegDate<>"" then
	sql = sql & " and b.initdate<= '" & NumberToDateFormat(eRegDate) & "'  "
end if
if sRegDate<>"" and eRegDate<>"" then
	sql = sql & " and (b.initdate>= '" & NumberToDateFormat(sRegDate) & "'   and b.initdate<= '" & NumberToDateFormat(eRegDate) & "' )"
end if

if sIniDate<>"" and eIniDate="" then
	sql = sql & " and a.initdate >= '" & NumberToDateFormat(sIniDate) & "'  "
end if
if sIniDate="" and eIniDate<>"" then
	sql = sql & " and a.initdate<= '" & NumberToDateFormat(eIniDate) & "'  "
end if
if sIniDate<>"" and eIniDate<>"" then
	sql = sql & " and (a.initdate>= '" & NumberToDateFormat(sIniDate) & "'   and b.initdate<= '" & NumberToDateFormat(eIniDate) & "' )"
end if

if yn<>"" then
	sql = sql & " and b.enable='"&yn&"'"
end if
if slevel<>"" then
	sql = sql & " and b.slevel='"&slevel&"'"
end if
if department<>"" then
	sql = sql & " and b.department='"&department&"'"
end if

sql = sql & " order by b.slevel ,b.department,b.grade,b.class1,a.sid,a.initdate desc"
'


'response.write sql
'response.end
rs.Open sql,msconn,adOpenStatic,adLockReadonly


'rs.close
'set rs=nothing
%>
<HTML>
<HEAD>
<TITLE> 【LDCC英外語能力診斷輔導中心】  </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>

</HEAD>
<BODY bottomMargin=0 leftMargin=2 topMargin=0 rightMargin=2 marginheight="0" marginwidth="0"  >

<%
	if not rs.eof then 
%>
<P align="center" class="inputlabel"><font size="4">學習問卷調查 - 學習策略明細表</font></P>

<TABLE cellSpacing=1 cellPadding=2 align="center" border=1 >
<TR class="inputlabel" bgcolor="#E7E7E7">
	<td nowrap rowspan="2">序號</td>
	<td nowrap rowspan="2">學制</td>
	<td nowrap rowspan="2">系所</td>
	<td nowrap rowspan="2">年級</td>
	<td nowrap rowspan="2">班別</td>
	<td nowrap rowspan="2">學號</td>
	<td nowrap rowspan="2">姓名</td>
	<td nowrap rowspan="2">填卷日期</td>
	<td nowrap colspan="9" align="center">Part A</td>
	<td nowrap colspan="14" align="center">Part B</td>
	<td nowrap colspan="6" align="center">Part C</td>
	<td nowrap colspan="9" align="center">Part D</td>
	<td nowrap colspan="6" align="center">Part E</td>
	<td nowrap colspan="6" align="center">Part F</td>
	<td nowrap colspan="7" align="center">average</td>

</TR>
<TR class="inputlabel" bgcolor="#E7E7E7">
	
	<td nowrap>Q1</td>
	<td nowrap>Q2</td>
	<td nowrap>Q3</td>
	<td nowrap>Q4</td>
	<td nowrap>Q5</td>
	<td nowrap>Q6</td>
	<td nowrap>Q7</td>
	<td nowrap>Q8</td>
	<td nowrap>Q9</td>
	<td nowrap>Q10</td>
	<td nowrap>Q11</td>
	<td nowrap>Q12</td>
	<td nowrap>Q13</td>
	<td nowrap>Q14</td>
	<td nowrap>Q15</td>
	<td nowrap>Q16</td>
	<td nowrap>Q17</td>
	<td nowrap>Q18</td>
	<td nowrap>Q19</td>
	<td nowrap>Q20</td>
	<td nowrap>Q21</td>
	<td nowrap>Q22</td>
	<td nowrap>Q23</td>
	<td nowrap>Q24</td>
	<td nowrap>Q25</td>
	<td nowrap>Q26</td>
	<td nowrap>Q27</td>
	<td nowrap>Q28</td>
	<td nowrap>Q29</td>
	<td nowrap>Q30</td>
	<td nowrap>Q31</td>
	<td nowrap>Q32</td>
	<td nowrap>Q33</td>
	<td nowrap>Q34</td>
	<td nowrap>Q35</td>
	<td nowrap>Q36</td>
	<td nowrap>Q37</td>
	<td nowrap>Q38</td>
	<td nowrap>Q39</td>
	<td nowrap>Q40</td>
	<td nowrap>Q41</td>
	<td nowrap>Q42</td>
	<td nowrap>Q43</td>
	<td nowrap>Q44</td>
	<td nowrap>Q45</td>
	<td nowrap>Q46</td>
	<td nowrap>Q47</td>
	<td nowrap>Q48</td>
	<td nowrap>Q49</td>
	<td nowrap>Q50</td>
	<td nowrap>ave&nbsp;A</td>
	<td nowrap>ave&nbsp;B</td>
	<td nowrap>ave&nbsp;C</td>
	<td nowrap>ave&nbsp;D</td>
	<td nowrap>ave&nbsp;E</td>
	<td nowrap>ave&nbsp;F</td>
	<td nowrap>Overall</td>
	
</TR>

<%
	
		while not rs.EOF 
			rc=rc +1
			if rc mod 2 = cint(0) then
				vcolor="#E0F7DD"
			else
				vcolor="#FFFFFF"
			end if
		%>
		<tr bgcolor="<%=vcolor%>">
			<td nowrap><%=rc%></td>
			<td nowrap><%=rs("slevel")%></td>
			<td nowrap><%=rs("department")%></td>
			<td nowrap><%=rs("grade")%></td>
			<td nowrap><%=rs("class1")%></td>
			<td nowrap><%=rs("sid")%></td>
			<td nowrap><%=rs("name")%></td>
			<td nowrap><%=rs("initdate")%></td>
			<td nowrap><%=rs("q1")%></td>
			<td nowrap><%=rs("q2")%></td>
			
			<td nowrap><%=rs("q3")%></td>
			<td nowrap><%=rs("q4")%></td>
			<td nowrap><%=rs("q5")%></td>
			<td nowrap><%=rs("q6")%></td>
			<td nowrap><%=rs("q7")%></td>
			<td nowrap><%=rs("q8")%></td>
			<td nowrap><%=rs("q9")%></td>
			<td nowrap><%=rs("q10")%></td>
			<td nowrap><%=rs("q11")%></td>
			<td nowrap><%=rs("q12")%></td>
			<td nowrap><%=rs("q13")%></td>
			<td nowrap><%=rs("q14")%></td>
			<td nowrap><%=rs("q15")%></td>
			<td nowrap><%=rs("q16")%></td>
			<td nowrap><%=rs("q17")%></td>
			<td nowrap><%=rs("q18")%></td>
			<td nowrap><%=rs("q19")%></td>
			<td nowrap><%=rs("q20")%></td>
			<td nowrap><%=rs("q21")%></td>
			<td nowrap><%=rs("q22")%></td>
			<td nowrap><%=rs("q23")%></td>
			<td nowrap><%=rs("q24")%></td>
			<td nowrap><%=rs("q25")%></td>
			<td nowrap><%=rs("q26")%></td>
			<td nowrap><%=rs("q27")%></td>
			<td nowrap><%=rs("q28")%></td>
			<td nowrap><%=rs("q29")%></td>
			<td nowrap><%=rs("q30")%></td>
			<td nowrap ><%=rs("q31")%></td>
			<td nowrap ><%=rs("q32")%></td>
			<td nowrap ><%=rs("q33")%></td>
			<td nowrap ><%=rs("q34")%></td>
			<td nowrap ><%=rs("q35")%></td>
			<td nowrap ><%=rs("q36")%></td>
			<td nowrap ><%=rs("q37")%></td>
			<td nowrap ><%=rs("q38")%></td>
			<td nowrap><%=rs("q39")%></td>
			<td nowrap ><%=rs("q40")%></td>
			<td nowrap ><%=rs("q41")%></td>
			<td nowrap ><%=rs("q42")%></td>
			<td nowrap ><%=rs("q43")%></td>
			<td nowrap ><%=rs("q44")%></td>
			<td nowrap ><%=rs("q45")%></td>
			<td nowrap ><%=rs("q46")%></td>
			<td nowrap ><%=rs("q47")%></td>
			<td nowrap ><%=rs("q48")%></td>
			<td nowrap ><%=rs("q49")%></td>
			<td nowrap ><%=rs("q50")%></td>
			<%
			partA = cint(ifnull(rs("q1"),0))+cint(ifnull(rs("q2"),0))+cint(ifnull(rs("q3"),0))+cint(ifnull(rs("q4"),0))+cint(ifnull(rs("q5"),0))+cint(ifnull(rs("q6"),0))+cint(ifnull(rs("q7"),0))+cint(ifnull(rs("q8"),0))+cint(ifnull(rs("q9"),0))
			partB = cint(ifnull(rs("q10"),0))+cint(ifnull(rs("q11"),0))+cint(ifnull(rs("q12"),0))+cint(ifnull(rs("q13"),0))+cint(ifnull(rs("q14"),0))+cint(ifnull(rs("q15"),0))+cint(ifnull(rs("q16"),0))+cint(ifnull(rs("q17"),0))+cint(ifnull(rs("q18"),0))+cint(ifnull(rs("q19"),0))+cint(ifnull(rs("q20"),0))+cint(ifnull(rs("q21"),0))+cint(ifnull(rs("q22"),0))+cint(ifnull(rs("q23"),0))
			partC = cint(ifnull(rs("q24"),0))+cint(ifnull(rs("q25"),0))+cint(ifnull(rs("q26"),0))+cint(ifnull(rs("q27"),0))+cint(ifnull(rs("q28"),0))+cint(ifnull(rs("q29"),0))
			partD = cint(ifnull(rs("q30"),0))+cint(ifnull(rs("q31"),0))+cint(ifnull(rs("q32"),0))+cint(ifnull(rs("q33"),0))+cint(ifnull(rs("q34"),0))+cint(ifnull(rs("q35"),0))+cint(ifnull(rs("q36"),0))+cint(ifnull(rs("q37"),0))+cint(ifnull(rs("q38"),0))
			partE = cint(ifnull(rs("q39"),0))+cint(ifnull(rs("q40"),0))+cint(ifnull(rs("q41"),0))+cint(ifnull(rs("q42"),0))+cint(ifnull(rs("q43"),0))+cint(ifnull(rs("q44"),0))
			partF = cint(ifnull(rs("q45"),0))+cint(ifnull(rs("q46"),0))+cint(ifnull(rs("q47"),0))+cint(ifnull(rs("q48"),0))+cint(ifnull(rs("q49"),0))+cint(ifnull(rs("q50"),0))
			Total = cint(partA) + cint(partB) + cint(partC) + cint(partD) + cint(partE) + cint(partF)
			%>
			 <td><%=round(cdbl(partA)/cdbl(9),1)%></td>
			 <td><%=round(cdbl(partB)/cdbl(14),1)%></td>
			 <td><%=round(cdbl(partC)/cdbl(6),1)%></td>
			 <td><%=round(cdbl(partD)/cdbl(9),1)%></td>
			 <td><%=round(cdbl(partE)/cdbl(6),1)%></td>
			 <td><%=round(cdbl(partF)/cdbl(6),1)%></td>
			 <td><%=round(cdbl(Total)/cdbl(50),1)%></td>
		</tr>
			
	<%	
			rs.movenext
		wend
	%>
</table>
	

<%
	else
		Response.Write "<FONT class=normal><FONT color=gray>- 沒有符合條件的資料顯示 -</FONT></FONT>"
	end if
%>



<P><BR>
</BODY>
</HTML>
<!-- #INCLUDE virtual="/include/lib/conn/syconnclose.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->
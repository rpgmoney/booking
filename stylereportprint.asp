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

sql = "select a.*,b.name,b.slevel ,b.department,b.grade,b.class1  from boo_questionnaire_style a "
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
<P align="center" class="inputlabel"><font size="4">學習問卷調查 - 學習風格明細表</font></P>

<TABLE cellSpacing=1 cellPadding=2 align="center" border=1 >
<TR class="inputlabel" bgcolor="#E7E7E7">
	<td nowrap >序號</td>
	<td nowrap>學制</td>
	<td nowrap>系所</td>
	<td nowrap>年級</td>
	<td nowrap>班別</td>
	<td nowrap>學號</td>
	<td nowrap>姓名</td>
	<td nowrap>填寫日期</td>
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
	<td nowrap>VISUAL</td>
	<td nowrap>AUDITORY</td>
	<td nowrap>KINESTHETIC</td>
	<td nowrap>TACTILE</td>
	<td nowrap>GROUP</td>
	<td nowrap>INDIVIDUAL</td>
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
			<td nowrap><%=trim(rs("name"))%></td>
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
			<td nowrap><%=(cint(ifnull(rs("q6"),0))+cint(ifnull(rs("q10"),0))+cint(ifnull(rs("q12"),0))+cint(ifnull(rs("q24"),0))+cint(ifnull(rs("q29"),0)))*2%></td>
			<td nowrap><%=(cint(ifnull(rs("q1"),0))+cint(ifnull(rs("q7"),0))+cint(ifnull(rs("q9"),0))+cint(ifnull(rs("q17"),0))+cint(ifnull(rs("q20"),0)))*2%></td>
			<td nowrap><%=(cint(ifnull(rs("q2"),0))+cint(ifnull(rs("q8"),0))+cint(ifnull(rs("q15"),0))+cint(ifnull(rs("q19"),0))+cint(ifnull(rs("q26"),0)))*2%></td>
			<td nowrap><%=(cint(ifnull(rs("q11"),0))+cint(ifnull(rs("q14"),0))+cint(ifnull(rs("q16"),0))+cint(ifnull(rs("q22"),0))+cint(ifnull(rs("q25"),0)))*2%></td>
			<td nowrap><%=(cint(ifnull(rs("q3"),0))+cint(ifnull(rs("q4"),0))+cint(ifnull(rs("q5"),0))+cint(ifnull(rs("q21"),0))+cint(ifnull(rs("q23"),0)))*2%></td>
			<td nowrap><%=(cint(ifnull(rs("q13"),0))+cint(ifnull(rs("q18"),0))+cint(ifnull(rs("q27"),0))+cint(ifnull(rs("q28"),0))+cint(ifnull(rs("q30"),0)))*2%></td>
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
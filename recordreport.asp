<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<!-- #INCLUDE virtual="/include/lib/asp/jsCalendar.asp" -->
<!-- #INCLUDE file="lib/parameter.inc" -->
<%
sid=trim(request("sid"))
forderid=trim(request("forderid"))
if sid="" or isnull(sid) or isempty(sid) then
	sid=session("sid")
end if



sender=trim(request("sender"))


set rs = server.CreateObject("adodb.recordset")
set rs2 = server.CreateObject("adodb.recordset")
sql = "select * from boo_profile   where sid='"& sid &"' and  classify in ('S','E')"
rs.Open sql,msconn,adOpenStatic,adLockReadonly
if  rs.EOF then
	showmessage ="你不是學員喔，此介面只提供學員瀏覽。"
else
	sid=trim(rs("sid"))
	name=trim(rs("name"))
	slevel=trim(rs("slevel"))
	grade=trim(rs("grade"))
	class1=trim(rs("class1"))
	department=trim(rs("department"))


end if

rs.close

'-----------------學習紀錄------------------------

'診斷
sql = "select content  from boo_book_T_M a  "
sql = sql & " left join boo_diagnosis b on a.id=b.tid  where a.sid='"&sid&"' and  a.item='診斷'  and  a.signin is not null  "

dia_total = 0
dia_Listening=0
dia_Speaking=0
dia_Reading=0
dia_Writing=0
dia_Grammar=0
dia_Pronunciation=0
dia_Testtaking=0
dia_PublicSpeaking=0
dia_PresentationSkills=0
dia_Other=0

rs.Open sql,msconn,adOpenStatic,adLockReadonly
while not rs.EOF
	dia_total = cdbl(dia_total) + 1
	if  InStr(rs("content"),"Listening")<>0 then
		dia_Listening = cdbl(dia_Listening)+1
	end if
	if  InStr(rs("content"),"Speaking")<>0 then
		dia_Speaking = cdbl(dia_Speaking)+1
	end if
	if InStr(rs("content"),"Reading")<>0 then
		dia_Reading = cdbl(dia_Reading)+1
	end if
	if InStr(rs("content"),"Writing")<>0 then
		dia_Writing = cdbl(dia_Writing)+1
	end if
	if InStr(rs("content"),"Grammar")<>0 then
		dia_Grammar = cdbl(dia_Grammar)+1
	end if
	if InStr(rs("content"),"Pronunciation")<>0 then
		dia_Pronunciation = cdbl(dia_Pronunciation)+1
	end if
	if InStr(rs("content"),"Test-taking")<>0 then
		dia_Testtaking = cdbl(dia_Testtaking)+1
	end if
	if InStr(rs("content"),"Public Speaking")<>0 then
		dia_PublicSpeaking = cdbl(dia_PublicSpeaking)+1
	end if
	if InStr(rs("content"),"Presentation Skills")<>0 then
		dia_PresentationSkills = cdbl(dia_PresentationSkills)+1
	end if
	if InStr(rs("content"),"Other")<>0 then
		dia_Other = cdbl(dia_Other)+1

	end if

	rs.MoveNext
wend

rs.Close
'諮商
sql = "select content  from boo_book_T_M a  "
sql = sql & " left join boo_consult b on a.id=b.tid  where a.sid='"&sid&"' and  a.item='諮商'  and  a.signin is not null  "

con_total = 0
con_Listening=0
con_Speaking=0
con_Reading=0
con_Writing=0
con_Grammar=0
con_Pronunciation=0
con_Testtaking=0
con_PublicSpeaking=0
con_PresentationSkills=0
con_Other=0

rs.Open sql,msconn,adOpenStatic,adLockReadonly
while not rs.EOF
	con_total = cdbl(con_total) + 1
	if  InStr(rs("content"),"Listening")<>0 then
		con_Listening = cdbl(con_Listening)+1
	elseif  InStr(rs("content"),"Speaking")<>0 then
		con_Speaking = cdbl(con_Speaking)+1
	elseif InStr(rs("content"),"Reading")<>0 then
		con_Reading = cdbl(con_Reading)+1
	elseif InStr(rs("content"),"Writing")<>0 then
		con_Writing = cdbl(con_Writing)+1
	elseif InStr(rs("content"),"Grammar")<>0 then
		con_Grammar = cdbl(con_Grammar)+1
	elseif InStr(rs("content"),"Pronunciation")<>0 then
		con_Pronunciation = cdbl(con_Pronunciation)+1
	elseif InStr(rs("content"),"Test-taking")<>0 then
		con_Testtaking = cdbl(con_Testtaking)+1
	elseif InStr(rs("content"),"Public Speaking")<>0 then
		con_PublicSpeaking = cdbl(con_PublicSpeaking)+1
	elseif InStr(rs("content"),"Presentation Skills")<>0 then
		con_PresentationSkills = cdbl(con_PresentationSkills)+1
	elseif InStr(rs("content"),"Other")<>0 then
		con_Other = cdbl(con_Other)+1

	end if

	rs.MoveNext
wend

rs.Close

'口語
op_total=0
op_ConversationTopic=0
op_IssuesinEnglish1=0
op_IssuesinEnglish2=0

sql = " select  oralset,count(*) as cnt   from boo_book_T_M where  sid='"&sid&"'  and  item='口語'  and signin is not null  and YN='Y' group by oralset "
rs.Open sql,msconn,adOpenStatic,adLockReadonly

while not rs.eof
	op_total = cdbl(op_total) + cdbl(rs("cnt"))
	if  trim(rs("oralset"))="Conversation Topics" then 
		op_ConversationTopic=rs("cnt")
	elseif  trim(rs("oralset"))="Issues in English I" then 
		op_IssuesinEnglish1=rs("cnt")
	elseif  trim(rs("oralset"))="Issues in English II" then 
		op_IssuesinEnglish2=rs("cnt")
	end if
	rs.Movenext
wend
rs.Close
''詩歌','簡報','寫作','閱讀'
crkp=0
pp=0
write1=0
read1=0
sql = "select  item,count(*) as cnt   from boo_book_T_M where  sid='"&sid&"'  and  item in ('詩歌','簡報','寫作','閱讀')  and signin is not null  and YN='Y' group by item"
rs.Open sql,msconn,adOpenStatic,adLockReadonly
while not rs.eof
	if  trim(rs("item"))="詩歌" then 
		crkp=rs("cnt")
	elseif  trim(rs("item"))="簡報" then 
		pp=rs("cnt")
	elseif  trim(rs("item"))="寫作" then 
		write1=rs("cnt")
	elseif  trim(rs("item"))="閱讀" then 
		read1=rs("cnt")
	end if
	rs.Movenext
wend
rs.Close

%>
<HTML>
<HEAD>
<TITLE> 【LDCC英外語能力診斷輔導中心】 </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/stm31.js" type="text/javascript"></SCRIPT>
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>
<script language="javascript">


</script>
</HEAD>
<BODY bottomMargin=0 leftMargin=0 topMargin=0 rightMargin=0 marginheight="0" marginwidth="0"  bgcolor="#f4c60d">
<TABLE cellSpacing=1 cellPadding=0 width="760"  height="100%" align="center" >
<TR><TD>
<TABLE cellSpacing=0 cellPadding=0 width="760"  height="100%" align="center" bgColor=#ffffff border=0>
<TR height="70"><TD><img src="images\top.jpg" border="0"></TD></TR>
<TR height="25" bgcolor="#333333">
	<TD align="center">
	<TABLE cellSpacing=0 cellPadding=0 border=0 width="100%">
	<TR>
	<TD align="left"><!-- #INCLUDE FILE="lib\link.inc" --></TD>
	<TD align="right"><!-- #INCLUDE FILE="lib\promsg.inc" --></TD>
	</TR>
	</TABLE>
	</TD>
</TR>
<TR valign="top">
	<TD>
<!-- ---------------------------------------------------------------------------------------- -->
	<P><BR>
	<TABLE cellSpacing=3 cellPadding=3 border=0 width="100%">
	<TR >
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">學員個人學習紀錄分析統計</TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>
	<TR >
		<TD></TD>
		<TD>
			<TABLE cellSpacing=1 cellPadding=2  border=0  width="95%"  >
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR >
						<TD class="inputlabel">學號：</TD><TD><%=sid%>&nbsp;</TD>
						<TD class="inputlabel">姓名：</TD><TD><%=name%>&nbsp;</TD>
						<TD class="inputlabel">學制：</TD><TD><%=slevel%>&nbsp;</TD>
						<TD class="inputlabel">系所：</TD><TD><%=department%>&nbsp;</TD>
						<TD class="inputlabel">年級：</TD><TD><%=grade%>&nbsp;</TD>
						<TD class="inputlabel">班級：</TD><TD><%=class1%>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD><TD align="right"><%if sender<>"" then%><input  type="button" value="&nbsp;&nbsp;返回&nbsp;&nbsp;" class="inputbutton" onclick="window.location='<%=sender%>'"><%end if%></TD></TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=2  border=1  width="70%">
			<TR class="inputlabel" align="center" ><TD width="30%">使用項目</TD><TD width="40%">項目</TD><TD width="15%">次數</TD><TD width="15%">總次數</TD></TR>
			<TR><TD rowspan="10" align="center" valign="center">診斷</TD><TD >Listening</TD><TD align="center">&nbsp;<%=dia_Listening%></TD><TD  align="center" rowspan="10">&nbsp;<%=dia_total%></TD></TR>
			<TR><TD >Speaking</TD><TD align="center">&nbsp;<%=dia_Speaking%></TD></TR>
			<TR><TD >Reading</TD><TD align="center">&nbsp;<%=dia_Reading%></TD></TR>
			<TR><TD >Writing</TD><TD align="center">&nbsp;<%=dia_Writing%></TD></TR>
			<TR><TD >Grammar</TD><TD align="center">&nbsp;<%=dia_Grammar%></TD></TR>
			<TR><TD >Pronunciation</TD><TD align="center">&nbsp;<%=dia_Pronunciation%></TD></TR>
			<TR><TD >Test-taking</TD><TD align="center">&nbsp;<%=dia_Testtaking%></TD></TR>
			<TR><TD >Public Speaking</TD><TD align="center">&nbsp;<%=dia_PublicSpeaking%></TD></TR>
			<TR><TD >Presentation Skills</TD><TD align="center">&nbsp;<%=dia_PresentationSkills%></TD></TR>
			<TR><TD >Other</TD><TD align="center">&nbsp;<%=dia_Other%></TD></TR>
			<TR class="inputlabel" align="center" ><TD >使用項目</TD><TD >項目</TD><TD >次數</TD><TD >總次數</TD></TR>
			<TR><TD rowspan="10" align="center" valign="center">諮商</TD><TD >Listening</TD><TD align="center">&nbsp;<%=con_Listening%></TD><TD  align="center" rowspan="10">&nbsp;<%=con_total%></TD></TR>
			<TR><TD >Speaking</TD><TD align="center">&nbsp;<%=con_Speaking%></TD></TR>
			<TR><TD >Reading</TD><TD align="center">&nbsp;<%=con_Reading%></TD></TR>
			<TR><TD >Writing</TD><TD align="center">&nbsp;<%=con_Writing%></TD></TR>
			<TR><TD >Grammar</TD><TD align="center">&nbsp;<%=con_Grammar%></TD></TR>
			<TR><TD >Pronunciation</TD><TD align="center">&nbsp;<%=con_Pronunciation%></TD></TR>
			<TR><TD >Test-taking</TD><TD align="center">&nbsp;<%=con_Testtaking%></TD></TR>
			<TR><TD >Public Speaking</TD><TD align="center">&nbsp;<%=con_PublicSpeaking%></TD></TR>
			<TR><TD >Presentation Skills</TD><TD align="center">&nbsp;<%=con_PresentationSkills%></TD></TR>
			<TR><TD >Other</TD><TD align="center">&nbsp;<%=con_Other%></TD></TR>
			<TR class="inputlabel" align="center" ><TD >使用項目</TD><TD >項目</TD><TD >次數</TD><TD >總次數</TD></TR>
			<TR><TD rowspan="3" align="center" valign="center">口語練習</TD><TD >Conversation Topic</TD><TD align="center">&nbsp;<%=op_ConversationTopic%></TD><TD align="center" rowspan="3">&nbsp;<%=op_total%></TD></TR>
			<TR><TD >Issues in English I</TD><TD align="center">&nbsp;<%=op_IssuesinEnglish1%></TD></TR>
			<TR><TD >Issues in English II</TD><TD align="center">&nbsp;<%=op_IssuesinEnglish2%></TD></TR>
			<TR class="inputlabel" align="center" ><TD >使用項目</TD><TD >&nbsp;</TD><TD >&nbsp;</TD><TD >總次數</TD></TR>
			<TR><TD align="center" >詩歌饒舌</TD><TD >&nbsp;</TD><TD >&nbsp;</TD><TD align="center">&nbsp;<%=crkp%></TD></TR>
			<TR><TD align="center" >簡報練習</TD><TD >&nbsp;</TD><TD >&nbsp;</TD><TD align="center">&nbsp;<%=pp%></TD></TR>
			<TR><TD align="center" >寫作諮商</TD><TD >&nbsp;</TD><TD >&nbsp;</TD><TD align="center">&nbsp;<%=write1%></TD></TR>
			<TR><TD align="center" >閱讀技巧</TD><TD >&nbsp;</TD><TD >&nbsp;</TD><TD align="center">&nbsp;<%=read1%></TD></TR>
			<TR class="inputlabel" align="center" ><TD >使用項目</TD><TD >項目</TD><TD >時數</TD><TD >總次數</TD></TR>
			<%
			'軟體
			set rsLoad = server.CreateObject("adodb.recordset")
			sql = "select a.*,b.summin from "
			sql = sql & " ( "
			sql = sql & " select * from boo_software where yn='Y' and category='S'  "
			sql = sql & " ) a left join "
			sql = sql & " ( "
			sql = sql & " select  item,sum(summin) as summin  from boo_book_software where    sid='"&sid&"'  and  yn='Y' and signin is not null and category='S'   group by item"
			sql = sql & " ) b on  a.id=b.item  order by floor,software"

			'response.write sql
			'response.end
			rsLoad.Open sql,msconn,adOpenStatic,adLockReadonly
			i = 0
			StrSoftware=""
			totalS = 0
			while not rsLoad.EOF
				i= i +1
				totalS = cdbl(totalS) + cdbl(ifnull(rsLoad("summin"),0) )
				StrSoftware = StrSoftware & "<TR> "
				if i=1 then
				StrSoftware = StrSoftware & " <TD rowspan=':tmprowspan' align=center>自學軟體</TD> "
				end if
				StrSoftware = StrSoftware & "<TD>&nbsp;" & rsLoad("floor") & "&nbsp;-&nbsp;" & rsLoad("software")  & "</TD> "
				StrSoftware = StrSoftware & " <TD align=center>" & ifnull(rsLoad("summin"),0) & " </TD> "
				if i=1 then
				StrSoftware = StrSoftware & " <TD  rowspan=':tmprowspan' align='center'> :totalS</TD> "
				end if
				StrSoftware = StrSoftware & "</TR> "
				rsLoad.MoveNext 
			wend
			rsLoad.close
			response.write replace(replace(StrSoftware,":tmprowspan",i),":totalS",totalS)
			%>
			<TR class="inputlabel" align="center" ><TD >使用項目</TD><TD >項目</TD><TD >時數</TD><TD >總次數</TD></TR>
			<%
			'補充教材
			set rsLoad = server.CreateObject("adodb.recordset")
			sql = "select a.*,b.summin from "
			sql = sql & " ( "
			sql = sql & " select * from boo_software where yn='Y' and category='T'  "
			sql = sql & " ) a left join "
			sql = sql & " ( "
			sql = sql & " select  item,sum(summin) as summin  from boo_book_software where    sid='"&sid&"'  and  yn='Y' and signin is not null and category='T'   group by item"
			sql = sql & " ) b on  a.id=b.item  order by floor,software"

			'response.write sql
			'response.end
			rsLoad.Open sql,msconn,adOpenStatic,adLockReadonly
			i = 0
			StrSoftware=""
			totalS = 0
			while not rsLoad.EOF
				i= i +1
				totalS = cdbl(totalS) + cdbl(ifnull(rsLoad("summin"),0) )
				StrSoftware = StrSoftware & "<TR> "
				if i=1 then
				StrSoftware = StrSoftware & " <TD rowspan=':tmprowspan' align=center>補充教材</TD> "
				end if
				StrSoftware = StrSoftware & "<TD>&nbsp;"  & rsLoad("software")  & "</TD> "
				StrSoftware = StrSoftware & " <TD align=center>" & ifnull(rsLoad("summin"),0) & " </TD> "
				if i=1 then
				StrSoftware = StrSoftware & " <TD  rowspan=':tmprowspan' align='center'> :totalS</TD> "
				end if
				StrSoftware = StrSoftware & "</TR> "
				rsLoad.MoveNext 
			wend
			rsLoad.close
			response.write replace(replace(StrSoftware,":tmprowspan",i),":totalS",totalS)
			%>
			</TABLE>
		</TD>
	</TR>
	<TR><TD></TD><TD>
		
	</TD></TR>
	</TABLE>
	<BR><BR>
<!-- ---------------------------------------------------------------------------------------- -->
	</TD>
</TR>
<TR bgcolor="#333333" height="30">
	<TD class="T1">
	<!-- #include file="lib\bottom.inc" -->
	
	</TD>
</TR>
</TABLE>

</TD>
</TR>
</TABLE>
<iframe style="display:none"  name="iframe_query" id="iframe_query"></iframe>
</BODY>
</HTML>

<!-- #INCLUDE virtual="/include/lib/conn/syconnclose.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->


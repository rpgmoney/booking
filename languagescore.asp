<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<!-- #INCLUDE virtual="/include/lib/asp/jsCalendar.asp" -->

<%
sid=trim(request("sid"))
category=trim(request("category"))
if sid="" or isnull(sid) or isempty(sid) then
	sid=session("sid")
end if
if sid="S224955279" then sid="1096200116" end if

sender=trim(request("sender"))


set dic = server.CreateObject("scripting.dictionary")
dic.Add "1","日"
dic.Add "2","一"
dic.Add "3","二"
dic.Add "4","三"
dic.Add "5","四"
dic.Add "6","五"
dic.Add "7","六"


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

%>
<HTML>
<HEAD>
<TITLE> 【LDCC英外語能力診斷輔導中心】 </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/stm31.js" type="text/javascript"></SCRIPT>
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>

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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">個人大專英檢成績紀錄</TD>
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
			<TABLE cellSpacing=0 cellPadding=0  border=0 width="95%" >
			<TR>
				<TD height="1" bgcolor="#000000" colspan="11"></TD>
			</TR>
			<TR class="inputlabel">
				<TD></TD><TD>年度</TD><TD>檢定種類</TD><TD>考試日期</TD><TD>級數</TD><TD>分數</TD>
				
			</TR>
			<TR>
				<TD height="1" bgcolor="#000000" colspan="11"></TD>
			</TR>
			<%
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
				if rs.EOF then
					response.write "<font   class=""norecord"" >無英檢成績紀錄</font>"
				else
					icnt=0
					while not rs.EOF
					icnt=icnt+1
					if icnt mod 2 = cint(0) then
						vcolor="#E7E7E7"
					else
						vcolor="#FFFFFF"
					end if
				%>
				<TR  bgcolor="<%=vcolor%>">
				<TD></TD>
				<TD><%=rs("yms_year")%></TD><TD><%=rs("ltk_name")%></TD>
				<TD><%=rs("test_date")%></TD>
				<TD><%=rs("level_id")%></TD><TD><%=rs("tot_score")%></TD>
				
				</TR>
				<TR>
					<TD height="1" background="/include/lib/images/untitled.bmp" colspan="11"></TD>
				</TR>
				<%
						rs.MoveNext
					wend 
				end if

				rs.Close



			%>
			</TABLE>
		</TD>
	</TR>
	
	</TABLE>
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


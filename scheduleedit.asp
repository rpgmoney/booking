<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<!-- #INCLUDE file="lib/parameter.inc" -->
<%
validate=trim(request("validate"))

scid=trim(request("scid"))
category=trim(request("category")) '老師或小老師
teacher=trim(request("teacher"))
yms=trim(request("yms"))
btime=trim(request("btime"))
bweek=trim(request("bweek"))
yn=trim(request("yn"))
deptgroup=trim(request("deptgroup"))

group1=trim(request("group1"))
skillcode=trim(request("skillcode"))
languagecode=trim(request("languagecode"))

if category="ST" and (group1="" or isnull(group1) or isempty(group1) ) then
	group1="小老師"
end if
sender=ifnull(replace(trim(request("sender")),"@","&"),"schedulelist.asp")


set rs = server.CreateObject("adodb.recordset")
if validate="edit" then
	sql = "select * from boo_schedule where scid='"&scid&"' "
	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
    if not rs.EOF then
		if btime<>"" then
            rs("btime")=btime
        end if
		if yms<>"" then
            rs("yms")=yms
        end if
		if bweek<>"" then
            rs("bweek")=bweek
        end if
		if teacher<>"" then
            rs("teacher")=teacher
        end if
		if yn<>"" then
            rs("yn")=yn
        end if
		if category<>"" then
            rs("category")=category
        end if
		if group1<>"" then
            rs("group1")=group1
        end if
		if deptgroup<>"" then
            rs("deptgroup")=deptgroup
        end if
		if skillcode<>"" then
            rs("skillcode")=skillcode
        end if
		if languagecode<>"" then
            rs("languagecode")=languagecode
        end if
		
		rs("initdate") = date()
		rs("inituid") = session("sid")


		rs.Update
        if Err.Number=0 then 
           '註冊之後填寫問卷 
          ' response.redirect "studentedit.asp?sid=" & sid 
          
        else
            showmessage= Err.Description
        end if

	else
		showmessage="找不到指定檔案。"
	end if

	rs.close
elseif validate="delete" then
	sql = "delete from boo_schedule where scid='"&scid&"'"
	'response.write sql

	msconn.Execute sql
	
else
	sql = "select * from boo_schedule where scid='"&scid&"'"
	'response.write sql
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
	if rs.EOF then
		response.redirect sender
	else
		category=trim(rs("category")) '老師或小老師
		teacher=trim(rs("teacher"))
		btime=trim(rs("btime"))
		bweek=trim(rs("bweek"))
		yn=trim(rs("yn"))
		group1=trim(rs("group1"))
		skillcode=trim(rs("skillcode"))
		languagecode=trim(rs("languagecode"))
		yms=trim(rs("yms"))
		deptgroup=trim(rs("deptgroup"))
	end if
	rs.close

end if


set rsLoad=server.CreateObject("adodb.recordset")
sql="select id,code,name,showcolor from boo_language where yn='Y' "
'response.write sql
rsLoad.Open sql,msconn,adOpenStatic,adLockReadOnly

StrLanguage=""
if rsLoad.state then	
	while not rsLoad.eof
		if languagecode=rsLoad("code") then
			StrLanguage=StrLanguage&"<option selected value="""&rsLoad("code")&""" style='color:"&rsLoad("showcolor")&"'>" & "■&nbsp;" & rsLoad("code") & "&nbsp;-&nbsp;" & rsLoad("name")&"</option>"
		else
			StrLanguage=StrLanguage&"<option value="""&rsLoad("code")&""" style='color:"&rsLoad("showcolor")&"'>" & "■&nbsp;" &rsLoad("code") & "&nbsp;-&nbsp;" &rsLoad("name")&"</option>"
		end if
		rsLoad.movenext
	wend
end if
rsLoad.close

sql="select id,code,name from boo_skill where yn='Y' "
'response.write sql
rsLoad.Open sql,msconn,adOpenStatic,adLockReadOnly

StrSkill=""
if rsLoad.state then	
	while not rsLoad.eof
		if skillcode=rsLoad("code") then
			StrSkill=StrSkill&"<option selected value="""&rsLoad("code")&""" >" & rsLoad("code") & "&nbsp;-&nbsp;" & rsLoad("name")&"</option>"
		else
			StrSkill=StrSkill&"<option value="""&rsLoad("code")&""" >"&rsLoad("code") & "&nbsp;-&nbsp;" &rsLoad("name")&"</option>"
		end if
		rsLoad.movenext
	wend
end if
rsLoad.close



set rsLoad=nothing	

%>
<HTML>
<HEAD>
<TITLE> 【LDCC英外語能力診斷輔導中心】 </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/stm31.js" type="text/javascript"></SCRIPT>
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>
<script language="javascript">

function check_input()
{
    var errmsg=""
	
	if (form1.teacher.value=="")
        errmsg += "老師不能為空白\n";
    if (form1.yms.value=="")
        errmsg += "學年學期不能為空白\n";
	if (form1.bweek.value=="")
        errmsg += "星期不能為空白\n";
	 if (form1.btime.value=="")
        errmsg += "時段不能為空白\n";
    if (form1.languagecode.value=="")
        errmsg += "語言不能為空白\n";
	if (form1.skillcode.value=="")
        errmsg += "專長不能為空白\n";
	if (form1.group1.value=="")
        errmsg += "系別不能為空白\n";
	if (form1.deptgroup.value=="")
        errmsg += "診斷諮商單位不能為空白\n";
    if (errmsg=="")
        return true;
    else
    {
        alert(errmsg);
        return false;
    }
}

function DeleteDc()
{
	var errmsg=""
	
	
	
	if (confirm("確定要刪除該筆資料嗎？")){
		form1.validate.value="delete";
		form1.submit();

	}
	
}
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center"><%if category="T" then response.write "修改駐診教師 " else response.write "修改小老師班表 " end if%></TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="scheduleedit.asp" onsubmit="return check_input();">
			<input type="hidden" value="edit" name="validate">
			<input type="hidden" id="sender" name="sender" value="<%=sender%>">
			<input type="hidden" id="category" name="category" value="<%=category%>">
			<input type="hidden" id="scid" name="scid" value="<%=scid%>">
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>教師：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text"  value="<%=teacher%>" maxlength="50" size="35" name="teacher" class="inputtext"  <%if session("sid")<>"95186" then response.write "readonly" end if%>>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>學年學期：</TD>
						<TD>星期：</TD>
						<TD>時段：</TD>
						<TD>開放否：</TD>
						<TD>診斷諮商單位：</TD>
					</TR>
					<TR>
						<TD>
						<select name="yms" class="inputtext">
						<option value=""> - 請指定 -</option>
						<%=YmsOption(94,Year(dateadd("m",-6,date()))-1911,yms)%>
						</select>
						</TD>
						<TD>
						<select name="bweek" class="inputtext">
						<option value=""> - 請指定 -</option>
						<option value="1" <%if bweek="1" then response.write "selected" end if%>>Monday - 星期一</option>
						<option value="2" <%if bweek="2" then response.write "selected" end if%>>Tuesday - 星期二 </option>
						<option value="3" <%if bweek="3" then response.write "selected" end if%>>Wednesday - 星期三</option>
						<option value="4" <%if bweek="4" then response.write "selected" end if%>>Thursday - 星期四</option>
						<option value="5" <%if bweek="5" then response.write "selected" end if%>>Friday - 星期五</option>
						</select>
						</TD>
						<TD>
							<select name="btime" class="inputtext">
							<option value=""> - 請指定 -</option>
							<optgroup label="上午">
							<%if category="ST" then%>
							<option value="0810" <%if btime="0810" then response.write "selected" end if %>>8:10～9:00</option>
							<option value="0910" <%if btime="0910" then response.write "selected" end if %>>9:10～10:00</option>
							<%end if%>
							<option value="1010" <%if btime="1010" then response.write "selected" end if %>>10:10～11:00</option>
							<option value="1110" <%if btime="1110" then response.write "selected" end if %>>11:10～12:00</option>
							</optgroup>
							<optgroup label="中午">
							<option value="1210" <%if btime="1210" then response.write "selected" end if %>>12:10～13:00</option>
							</optgroup>
							<optgroup label="下午">
							<option value="1310" <%if btime="1310" then response.write "selected" end if %>>13:10～14:00</option>
							<option value="1410" <%if btime="1410" then response.write "selected" end if %>>14:10～15:00</option>
							<option value="1510" <%if btime="1510" then response.write "selected" end if %>>15:10～16:00</option>
							<option value="1610" <%if btime="1610" then response.write "selected" end if %>>16:10～17:00</option>
							<%if category="ST" then%>
							<option value="1710" <%if btime="1710" then response.write "selected" end if %>>17:10～18:00</option>
							<%end if%>
							</optgroup>
							</select>
						</TD>
						<TD>
						<select name="yn" class="inputtext">
						<option value=""> - 請指定 -</option>
						<option value="Y" <%if yn="Y" then response.write "selected" end if%>>開放</option>
						<option value="N" <%if yn="N" then response.write "selected" end if%>>關閉</option>
						</select>
						</TD>
						<TD>
						<select name="deptgroup" class="inputtext">
						<option value=""> - 請指定 -</option>
						<option value="LDCC英外語能力診斷輔導中心"  <%if deptgroup="LDCC英外語能力診斷輔導中心" then response.write "selected" end if%> >LDCC英外語能力診斷輔導中心</option>
						<option value="ELC英語學習中心" <%if deptgroup="ELC英語學習中心" then response.write "selected" end if%>>ELC英語學習中心</option>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD <%if category="ST" then response.write "style='display:none'" end if%>>特殊專長領域：</TD>
						<TD>語言專長：</TD>
						<TD <%if category="ST" then response.write "style='display:none'" end if%>>系別：</TD>
						<TD></TD>
					</TR>
					<TR>
						<TD <%if category="ST" then response.write "style='display:none'" end if%>>
						<select name="skillcode" class="inputtext" style="width:150">
						<%if category="ST" then%>
						<option value="小老師">小老師</option>
						<%else%>
						<option value=""> - 請指定 -</option>
						<%=StrSkill%>
						<%end if%>
						</select>
						</TD>
						<TD>
						<select name="languagecode" class="inputtext" style="width:150">
						<option value=""> - 請指定 -</option>
						<%=StrLanguage%>
						</select>
						</TD>
						<TD <%if category="ST" then response.write "style='display:none'" end if%>>
						<select name="group1" class="inputtext" style="width:150">
						<%if category="ST" then%>
						<option value="小老師" <%if group1="小老師" then response.write "selected" end if%>>小老師</option>
						<%else%>
						<option value=""> - 請指定 -</option>
						<option value="外語教學系" <%if group1="外語教學系" then response.write "selected" end if%>>外語教學系</option>
						<option value="英文系" <%if group1="英文系" then response.write "selected" end if%>>英文系</option>
						<option value="翻譯系" <%if group1="翻譯系" then response.write "selected" end if%>>翻譯系</option>
						<option value="其它" <%if group1="其它" then response.write "selected" end if%>>其它</option>
						<%end if%>
						</select>
						</select>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR>
			<TD>
			<BR>
			<input  type="submit" value="修改" class="inputbutton" >
			<input  type="button" value="刪除" onclick="DeleteDc();" class="inputbutton" <%if session("sid")<>"S224955279" then response.write "disabled" end if%>>
			<input  type="button" value="返回" class="inputbutton" onclick="window.location='<%=sender%>'">
			</TD>
			</TR>
			</form>
			</TABLE>
		</TD>
	</TR>
	<TR>
	<TD></TD><TD valign="top">
		<%
		sql = "select * from boo_schedule a where teacher='"&teacher&"' and category='"&category&"' and yms='"&yms&"' order by bweek,btime"
		rs.Open sql,msconn,adOpenStatic,adLockReadonly
		%>
		<font color="blue">※該人員其它有課時段</font>
		<TABLE cellSpacing=1 cellPadding=0 width="70%"  border=0   >
		<TR>
			<TD height="1" bgcolor="#000000" colspan="11"></TD>
		</TR>
		<TR class="inputlabel">
			<TD></TD><TD>老師</TD><TD>學年學期/星期</TD><TD>時段</TD><TD <%if category="ST" then response.write "style='display:none'" end if%>>特殊專長領域</TD><TD <%if category="ST" then response.write "style='display:none'" end if%>>系別</TD><TD>語言專長</TD><TD>開放否</TD><TD></TD>
		</TR>
		<TR>
			<TD height="1" bgcolor="#000000" colspan="11"></TD>
		</TR>
		<%

		while not rs.EOF
		%>
		<TR>
			<TD><a href="scheduleedit.asp?scid=<%=rs("scid")%>&sender=<%=replace(sender,"&","@")%>"><img border="0" src="/include/lib/images/wri.gif"></a></TD>
			<TD><%=rs("teacher")%></TD><TD><%=rs("yms")%>/<%=rs("bweek")%></TD><TD><%=rs("btime")%></TD>
			<TD align="center" <%if category="ST" then response.write "style='display:none'" end if%>><%=rs("skillcode")%></TD>
			<TD <%if category="ST" then response.write "style='display:none'" end if%>><%=rs("group1")%></TD>
			<TD><%=rs("languagecode")%></TD><TD><%=rs("yn")%></TD><TD></TD><TD></TD>
		</TR>
		<TR>
				<TD height="1" background="/include/lib/images/untitled.bmp" colspan="11"></TD>
			</TR>
		<%
			rs.MoveNext
		wend
		rs.close
		
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

</BODY>
</HTML>
<!-- #INCLUDE virtual="/include/lib/conn/syconnclose.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->
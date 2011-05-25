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

category=trim(request("category")) '老師或小老師
'response.write "category=" & category & "<br>"
teacher=trim(request("teacher"))
yms=trim(request("yms"))
btime=trim(request("btime"))
bweek=trim(request("bweek"))
yn=trim(request("yn"))
deptgroup=trim(request("deptgroup"))
group1=trim(request("group1"))
skillcode=trim(request("skillcode"))
languagecode=trim(request("languagecode"))
if yms="" then
	yms=par_yms
end if
sender=ifnull(trim(request("sender")),"schedulelist.asp")
'response.write "sender = " & sender
set rs = server.CreateObject("adodb.recordset")
if validate="add" then
	sql = "select * from boo_schedule where bweek='"&bweek&"' and btime='"&btime&"' and teacher='"&teacher&"' and yms='"&yms&"'"
	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
    if rs.EOF then
        rs.AddNew
		scid=getguid()
		if scid<>"" then
            rs("scid")=scid
        end if
		if yms<>"" then
            rs("yms")=yms
        end if
		if btime<>"" then
            rs("btime")=btime
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
          ' response.redirect "scheduleadd.asp?category=" & category 
          
        else
            showmessage= Err.Description
        end if

	else
		showmessage="資料重覆。"
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center"><%if category="T" then response.write "新增駐診教師 " else response.write "新增小老師班表 " end if%></TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="scheduleadd.asp" onsubmit="return check_input();">
			<input type="hidden" value="add" name="validate">
			<input type="hidden" id="sender" name="sender" value="<%=sender%>">
			<input type="hidden" id="category" name="category" value="<%=category%>">
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>教師：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=teacher%>" maxlength="50" size="35" name="teacher" class="inputtext" >
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
						<option value="2" <%if bweek="2" then response.write "selected" end if%>>Tuesday - 星期二</option>
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
							<option value="0810">8:10～9:00</option>
							<option value="0910">9:10～10:00</option>
							<%end if%>
							<option value="1010">10:10～11:00</option>
							<option value="1110">11:10～12:00</option>
							</optgroup>
							<optgroup label="中午">
							<option value="1210">12:10～13:00</option>
							</optgroup>
							<optgroup label="下午">
							<option value="1310">13:10～14:00</option>
							<option value="1410">14:10～15:00</option>
							<option value="1510">15:10～16:00</option>
							<option value="1610">16:10～17:00</option>
							<%if category="ST" then%>
							<option value="1710">17:10～18:00</option>
							<%end if%>
							</optgroup>
							</select>
						</TD>
						<TD>
						<select name="yn" class="inputtext">
						<option value=""> - 請指定 -</option>
						<option value="Y" selected>開放</option>
						<option value="N">關閉</option>
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
						<select name="skillcode" class="inputtext" style="width:150" >
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
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR>
			<TD>
			<BR>
			<input  type="submit" value="新增" class="inputbutton" >
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
		sql = "select * from boo_schedule a where teacher='"&teacher&"' and category='"&category&"'"


		rs.Open sql,msconn,adOpenStatic,adLockReadonly
		%>
		<font color="blue">※該人員其它有課時段</font>
		<TABLE cellSpacing=1 cellPadding=0 width="70%"  border=0   >
		<TR>
			<TD height="1" bgcolor="#000000" colspan="11"></TD>
		</TR>
		<TR class="inputlabel">
			<TD></TD><TD>老師</TD><TD>星期</TD><TD>時段</TD><TD <%if category="ST" then response.write "style='display:none'" end if%>>特殊專長領域</TD><TD <%if category="ST" then response.write "style='display:none'" end if%>>系別</TD><TD>語言專長</TD><TD>開放否</TD><TD></TD>
		</TR>
		<TR>
			<TD height="1" bgcolor="#000000" colspan="11"></TD>
		</TR>
		<%

		while not rs.EOF
		%>
		<TR>
			<TD></TD>
			<TD><%=rs("teacher")%></TD><TD><%=rs("bweek")%></TD><TD><%=rs("btime")%></TD>
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
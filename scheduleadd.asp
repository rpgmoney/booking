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

category=trim(request("category")) '�Ѯv�Τp�Ѯv
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
		showmessage="��ƭ��СC"
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
			StrLanguage=StrLanguage&"<option selected value="""&rsLoad("code")&""" style='color:"&rsLoad("showcolor")&"'>" & "��&nbsp;" & rsLoad("code") & "&nbsp;-&nbsp;" & rsLoad("name")&"</option>"
		else
			StrLanguage=StrLanguage&"<option value="""&rsLoad("code")&""" style='color:"&rsLoad("showcolor")&"'>" & "��&nbsp;" &rsLoad("code") & "&nbsp;-&nbsp;" &rsLoad("name")&"</option>"
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
<TITLE> �iLDCC�^�~�y��O�E�_���ɤ��ߡj </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/stm31.js" type="text/javascript"></SCRIPT>
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>
<script language="javascript">

function check_input()
{
    var errmsg=""
	
	if (form1.teacher.value=="")
        errmsg += "�Ѯv���ର�ť�\n";
   if (form1.yms.value=="")
        errmsg += "�Ǧ~�Ǵ����ର�ť�\n";
	
	if (form1.bweek.value=="")
        errmsg += "�P�����ର�ť�\n";
	if (form1.btime.value=="")
        errmsg += "�ɬq���ର�ť�\n";
    if (form1.languagecode.value=="")
        errmsg += "�y�����ର�ť�\n";
	if (form1.skillcode.value=="")
        errmsg += "�M�����ର�ť�\n";
	if (form1.group1.value=="")
        errmsg += "�t�O���ର�ť�\n";
	if (form1.deptgroup.value=="")
        errmsg += "�E�_�԰ӳ�줣�ର�ť�\n";
	
	
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center"><%if category="T" then response.write "�s�W�n�E�Юv " else response.write "�s�W�p�Ѯv�Z�� " end if%></TD>
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
						<TD>�Юv�G</TD>
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
						<TD>�Ǧ~�Ǵ��G</TD>
						<TD>�P���G</TD>
						<TD>�ɬq�G</TD>
						<TD>�}��_�G</TD>
						<TD>�E�_�԰ӳ��G</TD>
					</TR>
					<TR>
						<TD>
						<select name="yms" class="inputtext">
						<option value=""> - �Ы��w -</option>
						<%=YmsOption(94,Year(dateadd("m",-6,date()))-1911,yms)%>
						</select>
						</TD>
						<TD>
						<select name="bweek" class="inputtext">
						<option value=""> - �Ы��w -</option>
						<option value="1" <%if bweek="1" then response.write "selected" end if%>>Monday - �P���@</option>
						<option value="2" <%if bweek="2" then response.write "selected" end if%>>Tuesday - �P���G</option>
						<option value="3" <%if bweek="3" then response.write "selected" end if%>>Wednesday - �P���T</option>
						<option value="4" <%if bweek="4" then response.write "selected" end if%>>Thursday - �P���|</option>
						<option value="5" <%if bweek="5" then response.write "selected" end if%>>Friday - �P����</option>
						</select>
						</TD>
						<TD>
							<select name="btime" class="inputtext">
							<option value=""> - �Ы��w -</option>
							<optgroup label="�W��">
							<%if category="ST" then%>
							<option value="0810">8:10��9:00</option>
							<option value="0910">9:10��10:00</option>
							<%end if%>
							<option value="1010">10:10��11:00</option>
							<option value="1110">11:10��12:00</option>
							</optgroup>
							<optgroup label="����">
							<option value="1210">12:10��13:00</option>
							</optgroup>
							<optgroup label="�U��">
							<option value="1310">13:10��14:00</option>
							<option value="1410">14:10��15:00</option>
							<option value="1510">15:10��16:00</option>
							<option value="1610">16:10��17:00</option>
							<%if category="ST" then%>
							<option value="1710">17:10��18:00</option>
							<%end if%>
							</optgroup>
							</select>
						</TD>
						<TD>
						<select name="yn" class="inputtext">
						<option value=""> - �Ы��w -</option>
						<option value="Y" selected>�}��</option>
						<option value="N">����</option>
						</select>
						</TD>
						<TD>
						<select name="deptgroup" class="inputtext">
						<option value=""> - �Ы��w -</option>
						<option value="LDCC�^�~�y��O�E�_���ɤ���"  <%if deptgroup="LDCC�^�~�y��O�E�_���ɤ���" then response.write "selected" end if%> >LDCC�^�~�y��O�E�_���ɤ���</option>
						<option value="ELC�^�y�ǲߤ���" <%if deptgroup="ELC�^�y�ǲߤ���" then response.write "selected" end if%>>ELC�^�y�ǲߤ���</option>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD <%if category="ST" then response.write "style='display:none'" end if%>>�S��M�����G</TD>
						<TD>�y���M���G</TD>
						<TD <%if category="ST" then response.write "style='display:none'" end if%>>�t�O�G</TD>
						<TD></TD>
					</TR>
					<TR>
						<TD <%if category="ST" then response.write "style='display:none'" end if%>>
						<select name="skillcode" class="inputtext" style="width:150" >
						<%if category="ST" then%>
						<option value="�p�Ѯv">�p�Ѯv</option>
						<%else%>
						<option value=""> - �Ы��w -</option>
						<%=StrSkill%>
						<%end if%>
						</select>
						</TD>
						<TD>
						<select name="languagecode" class="inputtext" style="width:150">
						<option value=""> - �Ы��w -</option>
						<%=StrLanguage%>
						
						</select>
						</TD>
						<TD <%if category="ST" then response.write "style='display:none'" end if%>>
						<select name="group1" class="inputtext" style="width:150">
						
						<%if category="ST" then%>
						<option value="�p�Ѯv" <%if group1="�p�Ѯv" then response.write "selected" end if%>>�p�Ѯv</option>
						<%else%>
						<option value=""> - �Ы��w -</option>
						<option value="�~�y�оǨt" <%if group1="�~�y�оǨt" then response.write "selected" end if%>>�~�y�оǨt</option>
						<option value="�^��t" <%if group1="�^��t" then response.write "selected" end if%>>�^��t</option>
						<option value="½Ķ�t" <%if group1="½Ķ�t" then response.write "selected" end if%>>½Ķ�t</option>
						<option value="�䥦" <%if group1="�䥦" then response.write "selected" end if%>>�䥦</option>
						<%end if%>
						</select>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR>
			<TD>
			<BR>
			<input  type="submit" value="�s�W" class="inputbutton" >
			<input  type="button" value="��^" class="inputbutton" onclick="window.location='<%=sender%>'">
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
		<font color="blue">���ӤH���䥦���Үɬq</font>
		<TABLE cellSpacing=1 cellPadding=0 width="70%"  border=0   >
		<TR>
			<TD height="1" bgcolor="#000000" colspan="11"></TD>
		</TR>
		<TR class="inputlabel">
			<TD></TD><TD>�Ѯv</TD><TD>�P��</TD><TD>�ɬq</TD><TD <%if category="ST" then response.write "style='display:none'" end if%>>�S��M�����</TD><TD <%if category="ST" then response.write "style='display:none'" end if%>>�t�O</TD><TD>�y���M��</TD><TD>�}��_</TD><TD></TD>
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
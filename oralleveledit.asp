<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<%
validate=trim(request("validate"))
nextrec=trim(request("nextrec"))
id=trim(request("id"))
category=trim(request("category"))
topic=trim(request("topic"))
rank=trim(request("rank"))
yn=trim(request("yn"))



sender=ifnull(trim(request("sender")),"orallevellist.asp")

set rs = server.CreateObject("adodb.recordset")
if validate="edit" then
	sql = "select * from boo_orallevel where id = '"&id&"' "
'	response.end
	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
    if not rs.EOF then
		if category<>"" then
            rs("category")=category
        end if
		if topic<>"" then
            rs("topic")=topic
        end if
		if rank<>"" then
            rs("rank")=rank
        end if
		if yn<>"" then
            rs("yn")=yn
        end if
		rs("initdate") = date()
		rs("inituid") = session("sid")


		rs.Update
        if Err.Number=0 then 
			response.redirect "orallevellist.asp"
		else
			showmessage= Err.Description
		end if

	else
		showmessage="找不到該筆資料。"
	end if

	rs.close
else
	sql = "select * from boo_orallevel where id='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
	if rs.EOF then
		response.redirect sender
	else
		id=trim(rs("id"))
		category=trim(rs("category"))
		topic=trim(rs("topic"))
		rank=trim(rs("rank"))
		yn=trim(rs("yn"))


	end if

end if


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
	
	if (form1.topic.value=="")
        errmsg += "口語題目不能為空白\n";
    if (form1.category.value=="")
        errmsg += "類別不能為空白\n";
	// if (form1.rank.value=="")
       // errmsg += "等級不能為空白\n";
    if (form1.yn.value=="")
        errmsg += "開放否不能為空白\n";
	
	
	
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">編輯口語題目資料</TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="oralleveledit.asp" onsubmit="return check_input();">
			<input type="hidden" value="edit" name="validate">
			<input type="hidden" id="id" name="id" value="<%=id%>">
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>口語題目：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=topic%>" maxlength="100" size="55" name="topic" class="inputtext" >
						</TD>
						
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>系列：</TD>
						<!-- <TD>等級：</TD> -->
						<TD>開放否：</TD>
					</TR>
					<TR>
						<TD>
						<select name="category" class="inputtext">
						<option value=""> - 請指定 -</option>
						<option value="My ET" <%if category="My ET" then response.write "selected" end if%>>My ET</option>
						<option value="Conversation Topics" <%if category="Conversation Topics" then response.write "selected" end if%>>Conversation Topics</option>
						<option value="Issues in English I" <%if category="Issues in English I" then response.write "selected" end if%>>Issues in English I</option>
						<option value="Issues in English II" <%if category="Issues in English II" then response.write "selected" end if%>>Issues in English II</option>
						</select>
						</TD>
						<!-- <TD>
						<select name="rank" class="inputtext">
						<option value=""> - 請指定 -</option>
						<option value="1" <%if rank="1" then response.write "selected" end if%>>大專英檢低於90分</option>
						<option value="2" <%if rank="2" then response.write "selected" end if%>>大專英檢高於90分</option>
						</select>
						</TD> -->
						<TD>
						<select name="yn" class="inputtext">
						<option value=""> - 請指定 -</option>
						<option value="Y" <%if yn="Y" then response.write "selected" end if%>>開放</option>
						<option value="N" <%if yn="N" then response.write "selected" end if%>>關閉</option>
						</select>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			
			<TR>
			<TD>
			<BR>
			<input  type="submit" value="儲存" class="inputbutton" >
			<input  type="button" value="返回" class="inputbutton" onclick="window.location='<%=sender%>'">
			</TD>
			</TR>
			</form>
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
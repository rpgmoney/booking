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
code=trim(request("code")) 
name=trim(request("name"))
yn=trim(request("yn"))
showcolor=trim(request("showcolor"))


sender=ifnull(trim(request("sender")),"language.asp")

set rs = server.CreateObject("adodb.recordset")
if validate="add" then
	sql = "select * from boo_language where code='"&code&"' "
	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
    if rs.EOF then
        rs.AddNew
		if code<>"" then
            rs("code")=code
        end if
		if name<>"" then
            rs("name")=name
        end if
		
		if yn<>"" then
            rs("yn")=yn
        end if
		if showcolor<>"" then
            rs("showcolor")=showcolor
        end if

		
		rs("initdate") = date()
		rs("inituid") = session("sid")


		rs.Update
        if Err.Number=0 then 
			if nextrec="Y" then
				validate=""
				nextrec=""
				code=""
				name=""
				showcolor=""
			else
				response.redirect "language.asp"
			end if
		else
			showmessage= Err.Description
		end if

	else
		showmessage="專長代號重覆。"
	end if

	rs.close
	
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
	
	if (form1.code.value=="")
        errmsg += "專長代號不能為空白\n";
    if (form1.name.value=="")
        errmsg += "專長名稱不能為空白\n";
	 if (form1.yn.value=="")
        errmsg += "開放否不能為空白\n";
	 if (form1.showcolor.value=="")
        errmsg += "代表色不能為空白\n";
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">新增語言專長</TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="languageadd.asp" onsubmit="return check_input();">
			<input type="hidden" value="add" name="validate">
			<input type="hidden" id="nextrec" name="nextrec">
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>專長代號：</TD>
						<TD>專長名稱：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=code%>" maxlength="10" size="10" name="code" class="inputtext" >
						</TD>
						<TD>
						<input type="text" value="<%=name%>" maxlength="50" size="55" name="name" class="inputtext" >
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>代表顏色：</TD>
						<TD>開放否：</TD>
					</TR>
					<TR>
						<TD>
						<select name="showcolor" class="inputtext" style="width:150">
						<option value=""> - 請指定 -</option>
						<option value="#CC9900" style="color:#CC9900;font-weight:bold " <%if showcolor="#CC9900" then response.write "selected" end if%>>■ 黃色</option>
						<option value="#3333FF" style="color:#3333FF;font-weight:bold "<%if showcolor="#3333FF" then response.write "selected" end if%>>■ 藍色</option>
						<option value="#FF0000" style="color:#FF0000;font-weight:bold "<%if showcolor="#FF0000" then response.write "selected" end if%>>■ 紅色</option>
						<option value="#339933" style="color:#339933;font-weight:bold "<%if showcolor="#339933" then response.write "selected" end if%>>■ 綠色</option>
						<option value="#9900CC" style="color:#9900CC;font-weight:bold "<%if showcolor="#9900CC" then response.write "selected" end if%>>■ 紫色</option>
						<option value="#FF9900" style="color:#FF9900;font-weight:bold "<%if showcolor="#FF9900" then response.write "selected" end if%>>■ 橘色</option>
						<option value="#660000" style="color:#660000;font-weight:bold "<%if showcolor="#660000" then response.write "selected" end if%>>■ 棕色</option>
						<option value="#000000" style="color:#000000;font-weight:bold "<%if showcolor="#000000" then response.write "selected" end if%>>■ 黑色</option>
						<option value="#999999" style="color:#999999;font-weight:bold "<%if showcolor="#999999" then response.write "selected" end if%>>■ 灰色</option>
						</select>
						</TD>
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
			<input  type="submit" value="新增" class="inputbutton" >
			<input  type="submit" onclick="form1.nextrec.value='Y'" value="新增後繼續新增" class="inputbutton">
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
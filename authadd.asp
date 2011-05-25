<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->

<%
'
today=year(date())-1911 & right("0" & month(date()),2) & right("0" & day(date()),2)


validate=trim(request("validate"))
deptname=trim(request("deptname"))
sid=trim(request("sid"))
name=trim(request("name"))
classify=trim(request("classify"))
enable=trim(request("enable"))

sender=ifnull(trim(request("sender")),"auth.asp")

set rs = server.CreateObject("adodb.recordset")
if validate="add" then
	sql = "select * from boo_profile where sid='"&sid&"' "
	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
    if rs.EOF then
        rs.AddNew
		if sid<>"" then
            rs("sid")=sid
        end if
		if name<>"" then
            rs("name")=name
        end if
		if deptname<>"" then
            rs("department")=deptname
        end if
		if classify<>"" then
            rs("classify")=classify
        end if
		if enable<>"" then
            rs("enable")=enable
        end if
		
		rs("initdate") = date()
		rs("inituid") = session("sid")
		

		rs.Update
        if Err.Number=0 then 
			if nextrec="Y" then
				validate=""
				nextrec=""
				deptname=""
				sid=""
				name=""
				classify=""
				enable=""
				
			else
				response.redirect sender
			end if
        else
            showmessage= Err.Description
        end if

	else
		showmessage="請勿重複新增。"
	end if
	rs.Close
end if

'

%>
<HTML>
<HEAD>
<TITLE> 【LDCC英外語能力診斷輔導中心】  </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/stm31.js" type="text/javascript"></SCRIPT>
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>
<script language="javascript">

function check_input()
{
    var errmsg=""
	
    if (form1.deptname.value=="")
        errmsg += "系所不能為空白\n";
	 if (form1.sid.value=="")
        errmsg += "學號不能為空白\n";

    if (errmsg=="")
        return true;
    else
    {
        alert(errmsg);
        return false;
    }
}

function clickblock(id)
{
	obj=document.getElementById("remark");
	if (obj!=null)
	{
		if (id==1)
		{
			obj.style.display="block";
			//obj1.src="images/icon-rectup.gif"
		}
		else
		{
			obj.style.display="none";
			//obj1.src="images/icon-rectdown.gif"
		}
	}
}

function ChkSid()
{
	vWinCal2 = window.open("lib/checksid.asp?sid="+form1.sid.value  ,"iframe_query", "width=200,height=50,status=no,resizable=no,scrollbars=no");
	vWinCal2.opener = form1;
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">老師/小老師名單新增</TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=2 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="authadd.asp" onsubmit="return check_input();">
			<input type="hidden" value="add" name="validate">
			<input type="hidden" id="nextrec" name="nextrec">
			<input type="hidden" id="category" name="category" value="T">
			
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>職員編號：</TD>
						<TD>姓名：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=sid%>" maxlength="25" size="35" onblur="ChkSid()" name="sid" class="inputtext" >
						</TD>
						<TD>
						<input type="text" value="<%=name%>" maxlength="25" size="35" name="name" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly >
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>單位</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=deptname%>" maxlength="25" size="65"  name="deptname" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly  >
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD >身份別</TD>
					</TR>
					<TR >
						<TD>
						<select name="classify" class="inputtext">
						<option value="T">老師</option>
						</select>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD >是否開啟</TD>
					</TR>
					<TR >
						<TD>
						<select name="enable" class="inputtext">
						<option value="Y">開啟</option>
						<option value="N">關閉</option>
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
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconnclose.asp" -->


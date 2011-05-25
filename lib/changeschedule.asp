<%@Language=VBScript LCID=1033%>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->

<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function window_onload() {
	window.close();
}

//-->
</SCRIPT>
<SCRIPT LANGUAGE=javascript FOR=window EVENT=onload>
<!--
 window_onload();
//-->
</SCRIPT>


<%
yms=trim(request("yms"))
StrSchedule=""
StrSchedule2=""


if yms <> "" then
	set rs2 = server.CreateObject("adodb.recordset")
	sql ="select distinct teacher from boo_schedule where yms='"&yms&"' and category='T' "
	rs2.Open sql,msconn,adOpenStatic,adLockReadonly
	if not rs2.EOF then
		StrSchedule="<option value="""" selected>- 請指定老師 -</option>"
	else
		StrSchedule="<option value="""" selected>- 無 -</option>"
	end if

	while not rs2.EOF
		StrSchedule=StrSchedule&"<option value="""&rs2("teacher")&""" >"&  rs2("teacher")&"</option>"
		rs2.MoveNext 
	wend
	set rs=nothing
else
	StrSchedule="<option value="""" selected>- 無 -</option>"
end if	

StrSchedule2="<select id=teachername name=teachername  class=inputtext >"& replace(StrSchedule,"""","'") & "</select>"

%>

<SCRIPT LANGUAGE=javascript>
<!--
	if (window.self.opener!=null)
	{
		if (window.self.opener.teachername!=null)
		{window.self.opener.teachername.outerHTML="<%=StrSchedule2%>";}
	}
//-->
</SCRIPT>

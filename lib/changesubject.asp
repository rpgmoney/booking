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
oralset=trim(request("oralset"))
StrSubject=""
StrSubject2=""


if oralset <> "" then
	set rs2 = server.CreateObject("adodb.recordset")
	sql ="select * from boo_orallevel where category='"&oralset&"' and yn='Y'"
	rs2.Open sql,msconn,adOpenStatic,adLockReadonly
	if not rs2.EOF then
		StrSubject="<option value="""" selected>- 請指定口語題目 -</option>"
	else
		StrSubject="<option value="""" selected>- 無 -</option>"
	end if

	while not rs2.EOF
		StrSubject=StrSubject&"<option value="""&rs2("topic")&""" >"&  rs2("topic")&"</option>"
		rs2.MoveNext 
	wend
	set rs=nothing
else
	StrSubject="<option value="""" selected>- 無 -</option>"
end if	

StrSubject2="<select id=topic name=topic class=inputtext >"& replace(StrSubject,"""","'") & "</select>"

%>

<SCRIPT LANGUAGE=javascript>
<!--
//	if (window.self.opener!=null)
//	{
//		if (window.self.opener.topic!=null)
//		{window.self.opener.topic.outerHTML="<%=StrSubject2%>";}

		if (window.parent.document.getElementById("topic")!=null)
		{window.parent.document.getElementById("topic").outerHTML="<%=StrSubject2%>";}
//	}
//-->
</SCRIPT>

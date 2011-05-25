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
if sid="S224955279" then sid="1095101007" end if

if forderid="" then
	forderid=1
end if


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
<script language="javascript">

var forderid="1";
function sel_tab(id)
{
	obj=document.getElementById("tab_"+forderid);
	obj1=document.getElementById("tab_"+forderid+"b");
	obj2=document.getElementById("folder_"+forderid);
	if (obj!=null)
	{
		obj.className="tabinactive";
		obj1.bgColor="silver";
		obj2.style.display="none";
	}
	obj=document.getElementById("tab_"+id);
	obj1=document.getElementById("tab_"+id+"b");
	obj2=document.getElementById("folder_"+id);
	if (obj!=null)
	{
		obj.className="tabactive";
		obj1.bgColor="white";
		obj2.style.display="block";
		forderid=id;
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">學員個人教師輔導療程紀錄</TD>
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
	<TR >
		<TD></TD><TD ><!-- #include file="lib\tab.inc" --></TD>
	</TR>
	<TR>
		<TD></TD><TD valign="top">
			
			<!-- #include file="psn_diagnosis.asp" -->
			<!-- #include file="psn_consult.asp" -->
			<!-- #include file="psn_op.asp" -->
			<!-- #include file="psn_write.asp" -->
			<!-- #include file="psn_crkp.asp" -->
			<!-- #include file="psn_pp.asp" -->
			<!-- #include file="psn_read.asp" -->
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


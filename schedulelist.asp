<%@Language=VBScript LCID=1033%>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<!-- #INCLUDE file="lib/parameter.inc" -->

<%
validate=trim(request("validate"))

teacher=trim(request("teacher"))
category=trim(request("category"))
bweek=trim(request("bweek"))
page=trim(request("page"))
YN=trim(request("YN"))
yms=trim(request("yms"))

if yms="" then
	yms=par_yms
end if

if YN="" then 
	YN="Y"
end if
sender=ifnull(trim(request("sender")),"studentlist.asp")



sender=server.urlencode(replace(request.servervariables("PATH_INFO")&"?page="&page& "&teacher=" & teacher& "&category=" & category,"%","*"))

if  validate="CloseSchedule" then 
	sqlm="update boo_schedule set yn='N'   where category='"&category&"' and yn='Y' and yms='"&yms&"'"
	msconn.Execute sqlm

end if 


%>
<HTML>
<HEAD>
<TITLE> �iLDCC�^�~�y��O�E�_���ɤ��ߡj </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/stm31.js" type="text/javascript"></SCRIPT>
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>
<script language="javascript">
function JumpPage1()
{
	var obj;
	obj= document.getElementById("selectPage");
	var index=obj.value;
	var  frmlistform = document.getElementById("news_form");
	frmlistform.page.value=index;
	frmlistform.submit();
}
function changepage(v)
{
	
	var  frmlistform = document.getElementById("news_form");
	frmlistform.page.value=v;
	frmlistform.submit();
}

function CloseSchedule()
{
	var errmsg=""
	
	
	
	if (confirm("�T�w�n�����Ҧ��ƯZ�ɬq�ܡH")){
		news_form.validate.value="CloseSchedule";
		news_form.submit();

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
<TR>
	<TD align="center"><font color="red"><%=showmessage%></font></TD>
</TR>

<TR valign="top">
	<TD>
<!-- ---------------------------------------------------------------------------------------- -->
	<P><BR>
	<TABLE cellSpacing=3 cellPadding=3 border=0 width="100%">
	<TR >
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%if category="T" then response.write "�n�E�Юv�ƯZ���@" else response.write "�p�Ѯv�Z����ƺ��@" end if%> </TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>
	
	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=0 cellPadding=0 border=0>
			<form id="news_form" name="news_form" method="post" action="schedulelist.asp" >
			<input type="hidden" name="page" value="">
			<input type="hidden" value="" name="validate">
			<input type="hidden" name="category" value="<%=category%>">
			<TR class="inputlabel"><TD width="20"></TD>
				<TD><%if category="T" then response.write "�Юv" else response.write "�p�Ѯv" end if%></TD>
				<TD>�P��</TD>
				<TD>�}��_</TD>
				<TD>�Ǧ~�Ǵ�</TD>
				<TD></TD>
			</TR>
			<TR><TD></TD>
				
				<td>
					<input type="text" name="teacher" value="<%=teacher%>"  maxlength="25" class="inputtext" >
				</td>
				<td>
					<select name="bweek" class="inputtext">
					<option value=""> - ���� -</option>
					<option value="1" <%if bweek="1" then response.write "selected" end if%>>Monday - �P���@</option>
					<option value="2" <%if bweek="2" then response.write "selected" end if%>>Tuesday - �P���G</option>
					<option value="3" <%if bweek="3" then response.write "selected" end if%>>Wednesday - �P���T</option>
					<option value="4" <%if bweek="4" then response.write "selected" end if%>>Thursday - �P���|</option>
					<option value="5" <%if bweek="5" then response.write "selected" end if%>>Friday - �P����</option>
					</select>
				</td>
				<TD>
				<select name="YN" class="inputtext">
				<option value="all"> - ���� -</option>
				<option value="Y" <%if yn="Y" then response.write "selected" end if%>>�}��</option>
				<option value="N" <%if yn="N" then response.write "selected" end if%>>����</option>
				</select>
				</TD>
				<TD>
				<select name="yms" class="inputtext">
				<option value="all"> - ���� -</option>
				<%=YmsOption(94,Year(dateadd("m",-6,date()))-1911,yms)%>
				</select>
				</TD>
				<TD>
				<input  type="submit" value="�d��" class="inputbutton">
				<input  type="button"  onclick="window.location='scheduleadd.asp?category=<%=category%>&sender=<%=sender%>'" value="�s�W" class="inputbutton">
				<input  type="button"   value="�����Ҧ��ɬq"  onclick="CloseSchedule();"   <%if  session("sid")<>"S224955279"   then response.write "disabled" end if%>  class="inputbutton">
				</TD>
			</TR>
			</form>
			</TABLE>
		</TD>
	</TR>
	<%
		set rs = server.CreateObject("adodb.recordset")
		sql = "select a.*,b.showcolor,b.name languagename from boo_schedule a left join boo_language b on a.languagecode=b.code where 1=1  "
		
		if teacher<>"" then
			sql = sql & " and a.teacher like'%"&teacher&"%' "
		end if
		if category<>"" then
			sql = sql & " and a.category='"&category&"' "
		end if
		if bweek<>"" then
			sql = sql & " and a.bweek='"&bweek&"' "
		end if
		if YN<>"all" then 
			sql = sql & " and a.YN='"&YN&"' "
		end if
		if yms<>"all" then
			sql = sql & " and a.yms='"&yms&"' "
		end if
		
		sql = sql & " order by yms desc,bweek,btime,teacher "
		
			
		'response.write sql
		rs.Open sql,msconn,adOpenStatic,adLockReadonly
		if not rs.EOF then
			rscount=rs.RecordCount
			lcount=30   '�]�w�C����ܪ�����
			m_page=request("page")
			if m_page="" then
				m_page=1
			else
				m_page=cint(m_page)   
			end if
			point=(m_page-1)*lcount+1   'Record Point
			if m_page>1 then
			  rs.move point-1
			end if

			'�p��@�X��
			pagecount=int(rscount/lcount)
			if rscount mod lcount >0 then
			  pagecount=pagecount+1
			end if   
			ln=point
		end if
	%>
	
	<TR>
		<TD></TD><TD valign="top">
		<!--�W�@�� , �U�@��  -->
			<TABLE cellSpacing=0 cellPadding=0 align="center" border=0 width="95%">
			<TR>
			<TD align="left">
			<TD>
				<TABLE cellSpacing=0 cellPadding=0 border=0 align="right">
				<TR>
				<%if m_page<=1 then  %>
				<TD><img src="/include/lib/images/arrow_left_no1.gif"></TD>
				<TD>&nbsp;<font color="#CCCCCC">�W�@��</font>&nbsp;</TD>
				<%else%>
				<TD><img src="/include/lib/images/arrow_left1.gif"></TD>
				<TD class="showhand" onclick="changepage(<%=m_page-1%>)">&nbsp;�W�@��&nbsp;</TD>
				<%end if%>
				<TD>�U</TD>
				<%if m_page>=pagecount then %>
				<TD>&nbsp;<font color="#CCCCCC">�U�@��&nbsp;</font></TD>
				<TD><img src="/include/lib/images/arrow_right_no1.gif"></TD>
				<%else%>
				<TD class="showhand" onclick="changepage(<%=m_page+1%>)">&nbsp;�U�@��&nbsp;</TD>
				<TD><img src="/include/lib/images/arrow_right1.gif"></TD>
				<%end if%>
				</TR>
				</TABLE>
			</TD></TR>
			</TABLE>
		<!--  -->
		</TD>
	</TR>
	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=0 cellPadding=0 align="center" border=0 width="95%">
			<TR>
				<TD height="1" bgcolor="#000000" colspan="11"></TD>
			</TR>
			<TR class="inputlabel">
				<TD></TD><TD><%if category="T" then response.write "�Юv" else response.write "�p�Ѯv" end if%></TD><TD>�Ǵ�/�P��</TD><TD>�ɬq</TD><TD align="center" <%if category="ST" then response.write "style='display:none'" end if%>>�S���M�����</TD><TD <%if category="ST" then response.write "style='display:none'" end if%>>�t�O</TD><TD>�y���M��</TD><TD align="center">�}��_</TD><TD></TD><TD></TD>
			</TR>
			<TR>
				<TD height="1" bgcolor="#000000" colspan="11"></TD>
			</TR>
			<% 
			icnt=0
			if rs.EOF then
				response.write "<TR><TD class=""norecord"" colspan=""11"">�S���ŦX���󪺸�����</TD></TR>"
			else
				do while not rs.eof and ln<=(point+lcount)-1 
				icnt=icnt+1
				if icnt mod 2 = cint(0) then
					if category = "T" then
						vcolor="#F8D6D1"
					else
						vcolor="#E7E7E7"
					end if
				else
					vcolor="#FFFFFF"
				end if
			%>
			<TR bgcolor="<%=vcolor%>">
				

				<TD><a href="scheduleedit.asp?scid=<%=rs("scid")%>&sender=<%=sender%>"><img border="0" src="/include/lib/images/wri.gif"></a></TD>
				<TD><%=rs("teacher")%></TD><TD><%=rs("yms")%>/<%=replace(replace(replace(replace(replace(rs("bweek"),"1","�P���@"),"2","�P���G"),"3","�P���T"),"4","�P���|"),"5","�P����")%></TD><TD><%=rs("btime")%></TD>
				<TD align="center" <%if category="ST" then response.write "style='display:none'" end if%>><%=rs("skillcode")%></TD><TD <%if category="ST" then response.write "style='display:none'" end if%>><%=rs("group1")%></TD>
				<TD><font color="<%=rs("showcolor")%>">��</font><%=rs("languagecode")%>&nbsp;-&nbsp;<%=rs("languagename")%></TD>
				<TD align="center"><%=rs("yn")%></TD><TD></TD><TD></TD>
			</TR>
			<TR>
				<TD height="1" background="/include/lib/images/untitled.bmp" colspan="11"></TD>
			</TR>
			<%
				rs.MoveNext
				ln=ln+1
				Loop
			end if
			

			%>
			</TABLE>
			
		</TD>
	</TR>
	<%if rscount>0 then %>
	<TR valign="top"><TD></TD>
	<TD >
			<table cellSpacing=1 cellPadding=2 border=0 align="right">
			<tr><td>
			<%
				response.write "��" & m_page & "��/�@" &pagecount &"��</td>"
				Response.Write "<td>&nbsp;��&nbsp;</td><td><select name=selectPage id=selectPage onchange=JumpPage1() class=inputtext style=width:50>"
				for i=1 to pagecount
					if (i<>m_page)  then
						Response.Write "<option value="&i&">"&i&"</option>"
					else
						Response.Write "<option value="&i&" selected>"&i&"</option>"
					end if
				Next
				Response.Write "</select><td>&nbsp;��</td></td>"
			%>
			<td width="20">&nbsp;</td></tr>
			</table>
	</TD></TR>
	<%end if%>
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
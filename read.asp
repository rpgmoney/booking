<%@Language=VBScript LCID=1033%>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<!-- #INCLUDE virtual="/include/lib/asp/jsCalendar.asp" -->

<%
validate=trim(request("validate"))
yn=trim(request("yn"))
sid=trim(request("sid"))
sdate = trim(request("sdate"))
edate = trim(request("edate"))
teachername= trim(request("teachername"))
if session("classify")="T" or session("classify")="E" then
	teachername=session("sname")
end if
page=trim(request("page"))
sender=ifnull(trim(request("sender")),"studentlist.asp")



sender=server.urlencode(replace(request.servervariables("PATH_INFO")&"?page="&page& "&sid=" & sid& "&yn=" & yn& "&sdate=" & sdate& "&edate=" & edate,"%","*"))

today=year(date())-1911 & right("0" & month(date()),2) & right("0" & day(date()),2)
if sdate="" or isempty(sdate) or isnull(sdate) then
	'sdate = today
end if

if yn="" or isempty(yn) or isnull(yn) then
	yn = "N"
end if

set dic = server.CreateObject("scripting.dictionary")
dic.Add "1","��"
dic.Add "2","�@"
dic.Add "3","�G"
dic.Add "4","�T"
dic.Add "5","�|"
dic.Add "6","��"
dic.Add "7","��"

if yn="" or isnull(yn) or isempty(yn) then
	yn="Y"
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
function signin_record_item(vid)
{
   
        news_form.validate.value="signin_item";
		news_form.signin_item.value=vid;
        news_form.submit();

}
function err_record_item(vid)
{
   
        news_form.validate.value="err_item";
		news_form.signin_item.value=vid;
        news_form.submit();

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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�ǭ��\Ū�������@</TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>
	
	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=0 cellPadding=0 border=0>
			<form id="news_form" name="news_form" method="post" action="read.asp" >
			<input type="hidden" name="page" value="<%=page%>">
			<input type="hidden" value="<%=category%>"  name="category" >
			<input type="hidden" value="<%=sender%>"  name="sender" >
			<input type="hidden" name="signin_item">
			<input type="hidden" name="validate">
			<TR class="inputlabel"><TD width="20"></TD>
				<TD>�w������_���G</TD>
				<TD>�Ǹ��Ωm�W�G</TD>
				<TD>�Ѯv�m�W�G</TD>
				<TD>���A</TD>
			</TR>
			<TR><TD></TD>
				<td>
					<TABLE cellSpacing=0 cellPadding=0 border=0 >
					<TD><input type="text" id="sdate" name="sdate" value="<%=sdate%>" size="12" maxlength="6" class="inputtext"><img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('sdate')" class="showhand"></TD>
					<TD class="inputlabel">&nbsp;&nbsp;~&nbsp;&nbsp;</TD>
					<TD><input type="text" id="edate" name="edate" value="<%=edate%>" size="12" maxlength="6" class="inputtext"><img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('edate')" class="showhand"></TD>
					</table>
				</td>
				<TD>
				<input type="text" value="<%=sid%>" maxlength="25" size="20"  name="sid" class="inputtext" >
				</TD>
				<TD>
				<input type="text" value="<%=teachername%>" maxlength="25" size="20"  name="teachername" class="inputtext" >
				</TD>
				<TD>
				<select name="yn" class="inputtext">
				<option value="all" <%if yn="all" then response.write "selected" end if%>> - ���� - </option>
				<option value="Y" <%if yn="Y" then response.write "selected" end if%>>�w�n��</option>
				<option value="N" <%if yn="N" then response.write "selected" end if%>>���n��</option>
				</select>
				</TD>
				<TD><input  type="submit" value="�d��" class="inputbutton"></TD>
			</TR>
			</form>
			</TABLE>
		</TD>
	</TR>
	<%
		set rs = server.CreateObject("adodb.recordset")
		sql = "select a.id as bookid,a.sid,a.name,a.department,a.grade,a.class1,a.bdate,a.btime,a.timeflag,a.score,a.teachername,b.* from boo_book_T_M a  "
		sql = sql & " left join boo_read b on a.id=b.tid  where a.item='�\Ū'  and  a.signin is not null  "
		if teachername <>"" then
			sql = sql & " and a.teachername like'"&teachername&"%' "
		end if
		if yn="Y" then
			sql = sql & " and   b.id is not null "
		elseif yn="N" then
			sql = sql & " and   b.id is  null "
		end if
		if sid<>"" then
			sql = sql & " and (a.sid='"&sid&"' or name like '%"&sid&"%' ) "
		end if
		if sdate<>"" and edate="" then
			sql = sql & " and a.bdate >= " & sdate
		end if
		if sdate="" and edate<>"" then
			sql = sql & " and a.bdate<=" & edate
		end if
		if sdate<>"" and edate<>"" then
			sql = sql & " and a.bdate>=" & sdate & " and a.bdate<=" & edate
		end if
		
		
		sql = sql & " order by bdate,btime "
		
			
		'response.write sql
		rs.Open sql,msconn,adOpenStatic,adLockReadonly
		if not rs.EOF then
			rscount=rs.RecordCount
			lcount=10   '�]�w�C����ܪ�����
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
				<TD></TD><TD width="32%">�w���H</TD><TD width="13%">���</TD><TD width="7%">�ɬq</TD><TD width="17%">�`��</TD><TD width="8%">�j�M�^��</TD><TD width="15%">�Юv</TD><TD width="8%" align="center"></TD>
			</TR>
			<TR class="inputlabel">
				<TD></TD><TD colspan="2">�\Ū�D�D</TD><TD colspan="5">�\Ū���D</TD>
			</TR>
			<TR class="inputlabel">
				<TD></TD><TD colspan="6">�Ѯv�^�X</TD><TD></TD>
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
					vcolor="#E7E7E7"
				else
					vcolor="#FFFFFF"
				end if
			%>
			<TR bgcolor="<%=vcolor%>">
				<TD>
				<a href="readedit.asp?id=<%=rs("bookid")%>&sender=<%=sender%>"><img border="0" src="/include/lib/images/wri.gif"></a>
				</TD>
				<TD><%=rs("sid")%> - <%=rs("name")%>(<%=rs("department")%>�A<%=rs("grade")%>)</TD>
				<TD><%=rs("bdate")& "(&nbsp;"&dic.Item(cstr(cint(weekday(NumberToDateFormat(rs("bdate"))))))&"&nbsp;)"%></TD><TD><%=rs("btime")%></TD>
				<TD><%=replace(replace(replace(rs("timeflag"),"U","�W�@�`(25��)"),"B","�U�@�`(25��)"),"A","�W�U�G�`(50��)")%></TD>
				<TD><%=rs("score")%></TD>
				<TD><%=rs("teachername")%></TD>
				<TD ></TD>
			</TR>
			<TR bgcolor="<%=vcolor%>">
				<TD></TD><TD colspan="2"><%=rs("subject")%>&nbsp;</TD>
				<TD colspan="5"><%=rs("content")%></TD>
			</TR>
			<TR bgcolor="<%=vcolor%>">
				<TD></TD><TD colspan="6"><%=rs("feedback")%>&nbsp;</TD><TD align="right"><a href="recordprofile.asp?sid=<%=rs("sid")%>&forderid=7&sender=<%=sender%>"><img border="0"  src="images/record.gif"></a></TD>
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
<%@Language=VBScript LCID=1033%>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<!-- #INCLUDE virtual="/include/lib/asp/jsCalendar.asp" -->

<%

validate=trim(request("validate"))
sid=trim(request("sid"))
teachername=trim(request("teachername"))
category=trim(request("category"))
sdate = trim(request("sdate"))
edate = trim(request("edate"))
yn=trim(request("yn"))
page=trim(request("page"))

'response.write "category=" &  category

item = trim(request("item"))
tid = trim(request("tid"))
summin= trim(request("summin"))


sender=ifnull(trim(request("sender")),"signinsoftware.asp")



sender=server.urlencode(replace(request.servervariables("PATH_INFO")&"?page="&page& "&sid=" & sid& "&yn=" & yn& "&sdate=" & sdate& "&edate=" & edate&"&category=" & category,"%","*"))

today=year(date())-1911 & right("0" & month(date()),2) & right("0" & day(date()),2)


if sdate="" or isempty(sdate) or isnull(sdate) then
	sdate = today
end if

if validate="signin_item" then
	
	signin_item = trim(request("signin_item"))
	sqld="update boo_book_software set  signin= Getdate()   where id='"&signin_item&"'"
	msconn.Execute sqld
	if Err.number=0 then
		'��s�B���Ҫ��n��w�����ɶ�
		if tid<>"" then
			updatesql = "update  boo_diagnosis_softwore set  times_c=times_c+"&cdbl(ifnull(summin,0))&"  where tid='"&tid&"' and sid='"&item&"'"
			msconn.Execute updatesql
		end if
		signin_item=""
    else
        showmessage= Err.Description
    end if
elseif  validate="err_item" then
	signin_item = trim(request("signin_item"))
	sqld="update boo_book_software set  signin= NULL   where id='"&signin_item&"'"
	msconn.Execute sqld
	if Err.number=0 then
		if category<>"F" then
			updatesql = "update  boo_diagnosis_softwore set  times_c=times_c-"&cdbl(ifnull(summin,0))&"  where tid='"&tid&"' and sid='"&item&"'"
			msconn.Execute updatesql
			signin_item=""
		end if

    else
        showmessage= Err.Description
    end if

elseif  validate="absence" then
	signin_item = trim(request("signin_item"))
	sqld="update boo_book_software set  yn='A'   where id='"&signin_item&"'"
	msconn.Execute sqld
	if Err.number=0 then
		if category<>"F" then
			updatesql = "update  boo_diagnosis_softwore set  times_b=times_b-"&cdbl(ifnull(summin,0))&"  where tid='"&tid&"' and sid='"&item&"'"
			msconn.Execute updatesql
			signin_item=""
		end if
    else
        showmessage= Err.Description
    end if
elseif  validate="err_absence" then
	signin_item = trim(request("signin_item"))
	sqld="update boo_book_software set  yn='Y'   where id='"&signin_item&"'"
	msconn.Execute sqld
	if Err.number=0 then
		if category<>"F" then
			updatesql = "update  boo_diagnosis_softwore set  times_b=times_b+"&cdbl(ifnull(summin,0))&"  where tid='"&tid&"' and sid='"&item&"'"
			msconn.Execute updatesql
			signin_item=""
		end if
    else
        showmessage= Err.Description
    end if
end if

set dic = server.CreateObject("scripting.dictionary")
dic.Add "1","��"
dic.Add "2","�@"
dic.Add "3","�G"
dic.Add "4","�T"
dic.Add "5","�|"
dic.Add "6","��"
dic.Add "7","��"


if sdate="" or isempty(sdate) or isnull(sdate) then
	sdate = today
end if
if edate="" or isempty(edate) or isnull(edate) then
	edate = today
end if
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
function signin_record_item(vid,vtid,vitem,vsummin)
{
		
        news_form.validate.value="signin_item";
		news_form.signin_item.value=vid;
		news_form.tid.value=vtid;
		news_form.item.value=vitem;
		news_form.summin.value=vsummin;
        news_form.submit();

}
function err_record_item(vid,vtid,vitem,vsummin)
{
   
        news_form.validate.value="err_item";
		news_form.signin_item.value=vid;
		news_form.tid.value=vtid;
		news_form.item.value=vitem;
		news_form.summin.value=vsummin;
        news_form.submit();

}
function absence_record_item(vid,vtid,vitem,vsummin)
{
   
        news_form.validate.value="absence";
		news_form.signin_item.value=vid;
		news_form.tid.value=vtid;
		news_form.item.value=vitem;
		news_form.summin.value=vsummin;
        news_form.submit();

}
function err_absence_item(vid,vtid,vitem,vsummin)
{
   
        news_form.validate.value="err_absence";
		news_form.signin_item.value=vid;
		news_form.tid.value=vtid;
		news_form.item.value=vitem;
		news_form.summin.value=vsummin;
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%if category="F" then response.write "�۾����{" else response.write  "�۾ǳn��/�ɥR�Ч�"%>&nbsp;ñ��@�~</TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>
	
	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=0 cellPadding=0 border=0>
			<form id="news_form" name="news_form" method="post" action="signinsoftware.asp" >
			<input type="hidden" name="page" value="<%=page%>">
			<input type="hidden" value="<%=category%>"  name="category" >
			<input type="hidden" name="signin_item">
			<input type="hidden" name="validate">
			<input type="hidden" name="tid">
			<input type="hidden" name="item">
			<input type="hidden" name="summin">
			<TR class="inputlabel"><TD width="20"></TD>
				<TD>�w������_���G</TD>
				<TD>�w���H�Ǹ��Ωm�W�G</TD>
				<TD>���A</TD>
				<TD></TD>
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
				<select name="yn" class="inputtext">
				<option value="all" <%if yn="all" then response.write "selected" end if%>> - ���� - </option>
				<option value="A" <%if yn="A" then response.write "selected" end if%>>�w������</option>
				<option value="Y" <%if yn="Y" then response.write "selected" end if%>>�w�w��</option>
				<option value="N" <%if yn="N" then response.write "selected" end if%>>�w����</option>
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
		sql = "select a.*,b.itemname from boo_book_software a  "
		sql = sql  & "left join ( "
		sql = sql &						"select floor+ '( ' +software  + ' )' as itemname ,id from boo_software where category='S' "
		sql =sql &						" union "
		sql = sql &						" select  software  as itemname ,id  from boo_software where category='T'  "
		sql =sql &						" union "
		sql = sql &						" select  item  as itemname ,id  from boo_self_item where yn='Y'  "
		sql = sql &				  " ) b  on  a.item=b.id  "
		sql = sql & "  where 1=1  "
		if yn<>"all" then
			sql = sql & " and  a.yn='"&yn&"'   "
		end if
		if category <>"" then
			sql = sql & " and a.category='"&category&"' "
		else
			sql = sql & " and a.category in ('S','T') "
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
		
		
		sql = sql & " order by bdate,stime,sid "
		
			
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
				<TD></TD><TD>���</TD><TD>�}�l�ɶ�</TD><TD>�����ɶ�</TD><TD>����</TD><TD>����</TD><TD>�w���H</TD><TD>�w�����A</TD>
			</TR>
			<TR class="inputlabel">
			<TD></TD>
			<TD colspan="8">����ɶ�</TD>
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
				<%if session("classify")="A" or session("sid")=rs("sid") then %>
				
				<a href="<% if category="F" then response.write "bookselfedit.asp" else  response.write "booksoftwareedit.asp" end if%>?id=<%=rs("id")%>&category=<%=rs("category")%>&sender=<%=sender%>"><img border="0" src="/include/lib/images/wri.gif"></a>
				<%end if%>
				</TD>
				<TD><%=rs("bdate")& "(&nbsp;"&dic.Item(cstr(cint(weekday(NumberToDateFormat(rs("bdate"))))))&"&nbsp;)"%></TD>
				<TD><%=rs("stime")%></TD><TD><%=rs("etime")%></TD>
				<TD><%=rs("summin")%></TD>
				<TD><%=rs("itemname")%></TD><TD><%=rs("sid")%> - <%=rs("name")%>(<%=rs("department")%>�A<%=rs("grade")%>)</TD>
				<TD>&nbsp;&nbsp;<%=replace(replace(replace(rs("yn"),"Y","<font color=""blue"">�w�w��</font>"),"N","<font color=""red"">����</font>"),"A","<font color=""#339900"">�w������</font>")%></TD>
	
			</TR>
			<TR bgcolor="<%=vcolor%>">
				<TD></TD>
				<TD colspan="5"><font color="red"><%=rs("signin")%></font> </TD>
				<TD align="right" colspan="2">
				<%if rs("yn")="A" then%>
					<input  type="button" value="�����w������" <%if rs("yn")="N" then response.write "disabled" end if%>  onclick="err_absence_item('<%=rs("id")%>','<%=rs("tid")%>','<%=rs("item")%>','<%=rs("summin")%>')"   class="inputbutton"  style='color:#3300CC;font-size=9pt' onMouseOver="this.style.color='#3399FF'" onMouseOut="this.style.color='#3300CC'">
				<%else%>
					<%if  rs("signin")<>"" then %>
					<input  type="button" value="��������" <%if rs("yn")="N" then response.write "disabled" end if%>  onclick="err_record_item('<%=rs("id")%>','<%=rs("tid")%>','<%=rs("item")%>','<%=rs("summin")%>')"   class="inputbutton"  style='color:FF6600;font-size=9pt' onMouseOver="this.style.color='#FF0000'" onMouseOut="this.style.color='#FF6600'">
					<%else%>
					<input  type="button" value="�w������" <%if rs("yn")="N" then response.write "disabled" end if%>  onclick="absence_record_item('<%=rs("id")%>','<%=rs("tid")%>','<%=rs("item")%>','<%=rs("summin")%>')"   class="inputbutton"  style='color:#3300CC;font-size=9pt' onMouseOver="this.style.color='#3399FF'" onMouseOut="this.style.color='#3300CC'">
					<input  type="button" value="����n�J" <%if rs("yn")="N" then response.write "disabled" end if%> onclick="signin_record_item('<%=rs("id")%>','<%=rs("tid")%>','<%=rs("item")%>','<%=rs("summin")%>')" class="inputbutton"  style='color:009900;font-size=9pt' onMouseOver="this.style.color='#00cc00'" onMouseOut="this.style.color='#009900'">
					<%end if%>
				<%end if%>
				</TD>
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
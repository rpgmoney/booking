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

if sid="" or isnull(sid) or isempty(sid) then
	sid=session("sid")
end if
if sid="95186" then sid="1095101007" end if

page=trim(request("page"))
sender=ifnull(trim(request("sender")),"studentlist.asp")

sender=server.urlencode(replace(request.servervariables("PATH_INFO")&"?page="&page& "&sid=" & sid& "&teachername=" & teachername& "&sdate=" & sdate& "&edate=" & edate&"&category=" & category,"%","*"))

today=year(date())-1911 & right("0" & month(date()),2) & right("0" & day(date()),2)
if sdate="" or isempty(sdate) or isnull(sdate) then
	sdate = today
end if


if validate="delete_item" then
delete_item = trim(request("delete_item"))
sqlm="update boo_book_lecture set yn='N',canceldate= Convert(varchar(10),Getdate(),111) ,canceluid='"&session("sid")&"' where id='"&delete_item&"'"
	msconn.Execute sqlm
	
	if Err.number=0 then
		yn="N"
    else
        showmessage= Err.Description
    end if

end if



set dic = server.CreateObject("scripting.dictionary")
dic.Add "1","日"
dic.Add "2","一"
dic.Add "3","二"
dic.Add "4","三"
dic.Add "5","四"
dic.Add "6","五"
dic.Add "7","六"

%>
<HTML>
<HEAD>
<TITLE> 【LDCC英外語能力診斷輔導中心】 </TITLE>
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
function delete_record_item(vid)
{
    if (confirm("您確定要取消該次預約資料嗎?"))
    {
        news_form.validate.value="delete_item";
		news_form.delete_item.value=vid;
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%if category="L" then response.write "取消已報名外語學習講座"  else response.write "取消已報名處方課程"  end if%> </TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>
	
	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=0 cellPadding=0 border=0>
			<form id="news_form" name="news_form" method="post" action="cancellecture.asp" >
			<input type="hidden" name="page" value="">
			<input type="hidden" value="<%=category%>"  name="category" >
			<input type="hidden" name="delete_item">
			<input type="hidden" name="validate">
			<TR class="inputlabel"><TD width="20"></TD>
				<TD></TD>
				<TD></TD>
			</TR>
			<TR><TD></TD>
				<td>
					<TABLE cellSpacing=0 cellPadding=0 border=0 >
					<TD><input type="hidden" id="sdate" name="sdate" value="<%=sdate%>" size="12" maxlength="6" class="inputtext"></TD>
					<TD class="inputlabel"></TD>
					<TD><input type="hidden" id="edate" name="edate" value="<%=edate%>" size="12" maxlength="6" class="inputtext"></TD>
					</table>
				</td>
				<TD>
				<input type="hidden" value="<%=sid%>" maxlength="25" size="20"  name="sid" class="inputtext" <%if session("classify")<>"A" then response.write "readonly" end if%>>
				</TD>
				<TD></TD>
			</TR>
			</form>
			</TABLE>
		</TD>
	</TR>
	<%
		set rs = server.CreateObject("adodb.recordset")
		sql = "select a.*,b.subject,b.date1,stime,etime  from boo_book_lecture a  "
		sql = sql  & "left join ( "
		sql = sql &						"select  id,subject,date1,stime,etime  from boo_lecture  where category='"&category&"'"
		sql = sql &				  " ) b  on  a.lid=b.id  "
		sql = sql & " where Cast(b.date1 as int)  > '" & today & "' "
		sql = sql & " and   a.sid='"&sid &"'   "
		if category <>"" then
			sql = sql & " and a.category='"&category&"' "
		end if
	
			
		'response.write sql
		rs.Open sql,msconn,adOpenStatic,adLockReadonly
		if not rs.EOF then
			rscount=rs.RecordCount
			lcount=10   '設定每頁顯示的筆數
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

			'計算共幾頁
			pagecount=int(rscount/lcount)
			if rscount mod lcount >0 then
			  pagecount=pagecount+1
			end if   
			ln=point
		end if
	%>
	
	<TR>
		<TD></TD><TD valign="top">
		<!--上一頁 , 下一頁  -->
			<TABLE cellSpacing=0 cellPadding=0 align="center" border=0 width="95%">
			<TR>
			<TD align="left">
			<TD>
				<TABLE cellSpacing=0 cellPadding=0 border=0 align="right">
				<TR>
				<%if m_page<=1 then  %>
				<TD><img src="/include/lib/images/arrow_left_no1.gif"></TD>
				<TD>&nbsp;<font color="#CCCCCC">上一頁</font>&nbsp;</TD>
				<%else%>
				<TD><img src="/include/lib/images/arrow_left1.gif"></TD>
				<TD class="showhand" onclick="changepage(<%=m_page-1%>)">&nbsp;上一頁&nbsp;</TD>
				<%end if%>
				<TD>｜</TD>
				<%if m_page>=pagecount then %>
				<TD>&nbsp;<font color="#CCCCCC">下一頁&nbsp;</font></TD>
				<TD><img src="/include/lib/images/arrow_right_no1.gif"></TD>
				<%else%>
				<TD class="showhand" onclick="changepage(<%=m_page+1%>)">&nbsp;下一頁&nbsp;</TD>
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
				<TD></TD><TD>日期</TD><TD><%if category="L" then response.write "外語學習講座"  else response.write "處方課程"  end if%></TD><TD>開始時間</TD><TD>結束時間</TD><TD></TD><TD>預約人</TD><TD>預約狀態</TD><TD></TD>
			</TR>
			<TR>
				<TD height="1" bgcolor="#000000" colspan="11"></TD>
			</TR>
			<% 
			icnt=0
			if rs.EOF then
				response.write "<TR><TD class=""norecord"" colspan=""11"">沒有符合條件的資料顯示</TD></TR>"
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
				</TD>
				<TD><%=rs("date1")& "(&nbsp;"&dic.Item(cstr(cint(weekday(NumberToDateFormat(rs("date1"))))))&"&nbsp;)"%></TD>
				<TD><%=rs("subject")%></TD><TD><%=rs("stime")%></TD>
				<TD><%=rs("etime")%></TD>
				<TD></TD><TD><%=rs("sid")%> - <%=rs("name")%>(<%=rs("department")%>，<%=rs("grade")%>)</TD>
				<TD>&nbsp;&nbsp;<%=replace(replace(rs("yn"),"Y","<font color=""blue"">已報名</font>"),"N","<font color=""red"">取消</font>")%></TD>
				<TD><input  type="button" value="取消" <%if rs("yn")="N" then response.write "disabled" end if%> onclick="delete_record_item('<%=rs("id")%>')" class="inputbutton" style='color:bc546f;font-size:9pt' onMouseOver="this.style.color='#ff6666'" onMouseOut="this.style.color='#bc546f'"></TD>
			</TR>
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
				response.write "第" & m_page & "頁/共" &pagecount &"頁</td>"
				Response.Write "<td>&nbsp;第&nbsp;</td><td><select name=selectPage id=selectPage onchange=JumpPage1() class=inputtext style=width:50>"
				for i=1 to pagecount
					if (i<>m_page)  then
						Response.Write "<option value="&i&">"&i&"</option>"
					else
						Response.Write "<option value="&i&" selected>"&i&"</option>"
					end if
				Next
				Response.Write "</select><td>&nbsp;頁</td></td>"
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
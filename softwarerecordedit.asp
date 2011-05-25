<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<!-- #INCLUDE virtual="/include/lib/asp/jsCalendar.asp" -->

<%
validate=trim(request("validate"))
id = trim(request("id"))
bdate=trim(request("bdate"))
sid=trim(request("sid"))
name=trim(request("name"))
slevel=trim(request("slevel"))
grade=trim(request("grade"))
class1=trim(request("class1"))
department=trim(request("department"))
category=trim(request("category"))
stimeh=trim(request("stime"))
stimem=trim(request("stime"))
etimeh=trim(request("etime"))
etimem=trim(request("etime"))

level=trim(request("level"))
topic=trim(request("topic"))
usageItem=trim(request("usageItem"))







today=year(date())-1911 & right("0" & month(date()),2) & right("0" & day(date()),2)

sender=ifnull(trim(request("sender")),"consult.asp" )

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

if validate="add" then
		sql = "select * from boo_software_record where tid='"&id&"' "
		rs.Open sql,msconn,adOpenStatic,adLockOptimistic
		if rs.EOF then
			rs.AddNew
			if id<>"" then
				rs("tid")=id
			end if
			if level<>"" then
				rs("level")=level
			end if
			if topic<>"" then
				rs("topic")=topic
			end if
			if usageItem<>"" then
				rs("usageItem")=usageItem
			end if
			if category<>"" then
				rs("category")=category
			end if
			rs("modifydate") = date()
			rs("modifyuid") = session("sid")
			rs("initdate") = date()
			rs("inituid") = session("sid")

			rs.Update
			if Err.Number=0 then 
				response.redirect sender
	
			else
				showmessage= Err.Description
			end if

		else
			showmessage="資料重覆。"
		end if

		rs.close	

elseif validate="edit" then
	sql = "select * from boo_software_record  where tid='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
	if  not rs.EOF then
		if level<>"" then
			rs("level")=level
		else
			rs("level")=null
		end if
		if topic<>"" then
			rs("topic")=topic
		else
			rs("topic")=null
		end if
		if usageItem<>"" then
			rs("usageItem")=usageItem
		else
			rs("usageItem")=null
		end if

		rs("modifydate") = date()
		rs("modifyuid") = session("sid")


		rs.Update
		if Err.Number=0 then 
			response.redirect sender

		else
			showmessage= Err.Description
		end if

	else
		showmessage="找不到該筆資料。"
	end if

	rs.close	
else
	'預約資訊
	sql = "select a.*,b.itemname  from boo_book_software a "
	sql = sql  & "left join ( "
	sql = sql &						"select floor+ '( ' +software  + ' )' as itemname ,id from boo_software where category='S' "
	sql =sql &						" union "
	sql = sql &						" select  software  as itemname ,id  from boo_software where category='T'  "
	sql =sql &						" union "
	sql = sql &						" select  item  as itemname ,id  from boo_self_item where yn='Y'  "
	sql = sql &				  " ) b  on  a.item=b.id  "
	sql = sql & " where a.id='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
	if rs.eof then
		response.redirect sender
	else
		bdate=trim(rs("bdate"))
		stimeh=left(trim(rs("stime")),2)
		stimem=right(trim(rs("stime")),2)
		etimeh=left(trim(rs("etime")),2)
		etimem=right(trim(rs("etime")),2)
		item=trim(rs("itemname"))
		sid=trim(rs("sid"))
		name=trim(rs("name"))
		slevel=trim(rs("slevel"))
		grade=trim(rs("grade"))
		class1=trim(rs("class1"))
		department=trim(rs("department"))
		summin=trim(rs("summin"))
		yn=trim(rs("yn"))
		canceldate=trim(rs("canceldate"))
		canceluid=trim(rs("canceluid"))
		signin=trim(rs("signin"))
	end if
	rs.close
	'諮商內容資訊
	sql = "select * from boo_software_record   where tid='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
    if rs.EOF then
		validate="add"
	else
		validate="edit"
		did=trim(rs("id"))

		level=trim(rs("level"))
		topic=trim(rs("topic"))
		
		usageItem=trim(rs("usageItem"))
		category=trim(rs("category"))

	end if
	rs.close
end if

if teacher="" or isnull(teacher) or isempty(teacher) then
	teacher=teachername
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

    if (errmsg=="")
        return true;
    else
    {
        alert(errmsg);
        return false;
    }
}

function check_all(title,num)
{
	//alert('title=' + title + '\n' + 'num=' + num);
	tmpobj = document.getElementById(title+'_c');

	if (tmpobj .checked==false){
		for (i=1;i<=num;i++){
			tmpobj1 = document.getElementById(title+i);
			tmpobj1.checked=false;
		}
	}
	else
	{
		for (i=1;i<=num;i++){
			tmpobj2 = document.getElementById(title+i);
			tmpobj2.checked=true;
		}
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center"><%if category="T" then response.write "學員補充教材紀錄編輯" else response.write "學員自學軟體紀錄編輯" end if%></TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=2 width="90%"  border=0    bgcolor="#F3EFE6" style="border: 1px solid #FF66CC; padding: 0">
			
			
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR >
						<TD class="inputlabel">學號：</TD><TD><%=sid%></TD>
						<TD class="inputlabel">姓名：</TD><TD><%=name%></TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR>
						<TD class="inputlabel">學制：</TD><TD><%=slevel%></TD>
						<TD class="inputlabel">系所：</TD><TD><%=department%></TD>
						<TD class="inputlabel">年級：</TD><TD><%=grade%></TD>
						<TD class="inputlabel">班級：</TD><TD><%=class1%></TD>
						<TD></TD>
					</TR>
				</TABLE>
			</TD></TR>
			
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR>
						<TD  class="inputlabel">預約項目：&nbsp;&nbsp;</TD><TD><%=item%></TD>
						<TD  class="inputlabel">預約日期：</TD><TD><%=bdate%></TD>
						<TD  class="inputlabel">星期：</TD><TD><%=dic.Item(cstr(cint(weekday(NumberToDateFormat(bdate)))))%></TD>
						<TD  class="inputlabel">預約時段：</TD><TD><%=btime%></TD>
						<TD><% response.write stimeh & "：" & stimem & "～" & etimeh &  "：" & etimem %></TD>
					</TR>
		
				</TABLE>
			</TD></TR>
		
			</form>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD></TD><TD>
		<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0 >
		<TR><TD class="errmsg"><%=showmessage1%></TD></TR>
		<TR>
		<TD width="100%" valign="Top">
			<form id="form1" name="form1" method="post" action="softwarerecordedit.asp"   onsubmit="return check_input();">
			<input type="hidden" value="<%=validate%>" name="validate">
			<input type="hidden" value="<%=ptime%>" name="ptime">
			<input type="hidden" value="<%=id%>" name="id">
			<input type="hidden" value="<%=category%>" name="category">
			<input type="hidden" value="<%=sender%>" name="sender">
			<input type="hidden" value="<%=sid%>" name="sid">
			<input type="hidden" value="<%=name%>" name="name">
			<input type="hidden" value="<%=bdate%>" name="bdate">
			<input type="hidden" value="<%=slevel%>" name="slevel">
			<input type="hidden" value="<%=grade%>" name="grade">
			<input type="hidden" value="<%=class1%>" name="class1">
			<input type="hidden" value="<%=department%>" name="department">

			<input type="hidden" value="<%=stimeh%>" name="stimeh">
			<input type="hidden" value="<%=stimem%>" name="stimem">
			<input type="hidden" value="<%=etimeh%>" name="etimeh">
			<input type="hidden" value="<%=etimem%>" name="etimem">
		
			<TABLE class=normal cellSpacing=0 cellPadding=0 width="100%"  border=0>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD><%if category="S" then response.write "層級" else if category="T" then response.write "使用項目" else if category="F" then response.write "學習內容/紀錄"  end if%>：</TD>
						<TD><%if category="S" then response.write "主題："  end if%></TD>
					</TR>
					<TR>
					<%if  category="S" then %>
						<TD>
						<select name="level" class="inputtext">
						<option value=""> - 未指定層級 -</option>
						<option value="Level 1" <%if level="Level 1" then response.write "selected" end if%>>Level 1</option>
						<option value="Level 2" <%if level="Level 2" then response.write "selected" end if%>>Level 2</option>
						<option value="Level 3" <%if level="Level 3" then response.write "selected" end if%>>Level 3</option>
						<option value="Level 4" <%if level="Level 4" then response.write "selected" end if%>>Level 4</option>
						</select>
						</TD>
						<TD>
						<input type="text" value="<%=topic%>" maxlength="100" size="80" name="topic" class="inputtext">
						</TD>
					<%elseif  category="T" then%>
						<TD>
						<input type="text" value="<%=usageItem%>" maxlength="100" size="80" name="usageItem" class="inputtext">
						</TD>
					<%elseif  category="F" then%>
						<TD>
						<textarea name="usageItem"  class="inputtext"  cols="100" rows="5"><%=usageItem%></textarea>
						</TD>
					<%end if%>
					</TR>
				</TABLE>
			</TD></TR>
			
			<TR><TD><BR><input  type="submit" value="儲存" class="inputbutton" ><input  type="button" value="返回" class="inputbutton" onclick="window.location='<%=sender%>'"></TD></TR>
			
			</TABLE>	
			</form>
		</TD>
		</TR>
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

<!-- #INCLUDE virtual="/include/lib/conn/syconnclose.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->


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
id=trim(request("id"))
category=trim(request("category")) '老師或小老師
bdate=trim(request("bdate"))
stimeh=trim(request("stimeh"))
stimem=trim(request("stimem"))
etimeh=trim(request("etimeh"))
etimem=trim(request("etimem"))
item=trim(request("item"))
sid=trim(request("sid"))
name=trim(request("name"))
slevel=trim(request("slevel"))
grade=trim(request("grade"))
class1=trim(request("class1"))
department=trim(request("department"))
score=trim(request("score"))
summin=trim(request("summin"))
yn=trim(request("yn"))
tid=trim(request("tid"))
software=trim(request("software"))
floor=trim(request("floor"))
'response.write "summin" & summin
'response.end
set rs = server.CreateObject("adodb.recordset")
set rs2 = server.CreateObject("adodb.recordset")

btnstatus=""

'response.write "btime=" & btime

sender=ifnull(trim(request("sender")),"booksoftwarelist.asp")
today=year(date())-1911 & right("0" & month(date()),2) & right("0" & day(date()),2)

if validate="edit" then
	if category="S" then
		checkflag = "true"
		'檢查同天是否有預約其它時間,一天 不得超過二個小時
		sql = "select   isnull(sum(summin),0) as summin  from  boo_book_software where  bdate='"&bdate&"' and sid='"&sid&"' and id<>'"&id&"'  and yn='Y' "
		'response.write sql
		'response.end
		rs.Open sql,msconn,adOpenStatic,adLockReadonly
		if not  rs.EOF then
			tmpsum = rs("summin")
			if  (cdbl(tmpsum) + cdbl(summin)  >120) then
				showmessage = "你在同一天有預約其它時段喔，同一天預約不得超過二個小時"
				checkflag = "false"
			end if
		end if
		rs.close

		if checkflag="true" then

			sql = "select * from boo_book_software where id='"&id&"' "
			rs.Open sql,msconn,adOpenStatic,adLockOptimistic
			if not rs.EOF then
				if bdate<>"" then
					rs("bdate")=bdate
				end if
				if stimeh<>"" and stimem<>""  then
					rs("stime")=right("0" & stimeh,2) & right("0" & stimem,2)
				end if
				if etimeh<>"" and etimem<>""  then
					rs("etime")=right("0" & etimeh,2) & right("0" & etimem,2)
				end if
				if summin<>"" then
					rs("summin")=summin
				end if
				if item<>"" then
					rs("item")=item
				end if
				if name<>"" then
					rs("name")=name
				end if
				if slevel<>"" then
					rs("slevel")=slevel
				end if
				if grade<>"" then
					rs("grade")=grade
				end if
				if class1<>"" then
					rs("class1")=class1
				end if
				if department<>"" then
					rs("department")=department
				end if
				rs("initdate") = date()
				rs("inituid") = session("sid")
		
		
				rs.Update
				if Err.Number=0 then 
					response.redirect "booksoftwarelist.asp?category=S"
		
				else
					showmessage= Err.Description
				end if

			else
				showmessage="找不到該筆資料，請洽資教中心。"
			end if

			rs.close
		end if
	else
		'預約補充教材
		sql = "select * from boo_book_software where id='"&id&"' "
		rs.Open sql,msconn,adOpenStatic,adLockOptimistic
		if not rs.EOF then
			if bdate<>"" then
				rs("bdate")=bdate
			end if
			if stimeh<>"" and stimem<>""  then
				rs("stime")=right("0" & stimeh,2) & right("0" & stimem,2)
			end if
			if etimeh<>"" and etimem<>""  then
				rs("etime")=right("0" & etimeh,2) & right("0" & etimem,2)
			end if
			if summin<>"" then
				rs("summin")=summin
			end if
			if item<>"" then
				rs("item")=item
			end if
			if name<>"" then
				rs("name")=name
			end if
			if slevel<>"" then
				rs("slevel")=slevel
			end if
			if grade<>"" then
				rs("grade")=grade
			end if
			if class1<>"" then
				rs("class1")=class1
			end if
			if department<>"" then
				rs("department")=department
			end if
			rs("initdate") = date()
			rs("inituid") = session("sid")
	
	
			rs.Update
			if Err.Number=0 then 
				response.redirect "booksoftwarelist.asp?category=T"
	
			else
				showmessage= Err.Description
			end if
		end if

	end if
elseif validate="delete" then

	sqlm="update boo_book_software set yn='N',canceldate= Convert(varchar(10),Getdate(),111) ,canceluid='"&session("sid")&"' where id='"&id&"'"
	msconn.Execute sqlm
	'更新處方籤的軟體已預約時間
	updatesql = "update  boo_diagnosis_softwore set  times_b=times_b-"&cdbl(summin)&"  where tid='"&tid&"' and sid='"&item&"'"
	msconn.Execute updatesql
	
	if Err.number=0 then
		yn="N"
    else
        showmessage= Err.Description
    end if
else
		sql = "select a.*,b.floor,b.software  from boo_book_software a  left join boo_software b on a.item=b.id where a.id='"&id&"' "
		rs.Open sql,msconn,adOpenStatic,adLockReadonly
		if rs.eof then
			response.redirect sender
		else
			bdate=trim(rs("bdate"))
			stimeh=left(trim(rs("stime")),2)
			stimem=right(trim(rs("stime")),2)
			etimeh=left(trim(rs("etime")),2)
			etimem=right(trim(rs("etime")),2)
			item=trim(rs("item"))
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
			software=trim(rs("software"))
			floor=trim(rs("floor"))
			tid=trim(rs("tid"))
		end if
		rs.close

end if

'取消資料不能變更
if  yn="N" or yn="A" or  signin<>""  then
	btnstatus="disabled"
end if
'過期資料除非管理者,否則無法變更
if session("classify")="S" and  cdbl(bdate) =< cdbl(today) then

	btnstatus="disabled"

end if
'response.write "classify=" & session("classify") & "<br>"
'response.write  "today=" & today & "<br>"
'response.write  "btnstatus=" & btnstatus
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
	var stime=0;
	var etime=0;
	var ttime=0;
	if (form1.sid.value=="")
        errmsg += "學號不能為空白\n";
    if (form1.stimeh.value=="" || form1.stimem.value=="" || form1.etimeh.value=="" || form1.etimem.value=="")
        errmsg += "請指定時間起迄\n";
	if (form1.bdate.value=="")
        errmsg += "日期不能為空白\n";
  
	<%if category="S" then%>
	if (errmsg=="")
	{
		
		stime = parseInt(form1.stimeh.value)*60 + parseInt(form1.stimem.value);
		etime = parseInt(form1.etimeh.value)*60 + parseInt(form1.etimem.value);
		ttime =parseInt(etime) - parseInt(stime) ;
		if (stime >= etime)
			errmsg += "開始時間必須小於結束時間\n";
		
		if ( ttime >120)
			errmsg += "每天使用時間不得大於二個小時\n";
	
	}
	<%else%>
		stime = parseInt(form1.stimeh.value)*60 + parseInt(form1.stimem.value);
		etime = parseInt(form1.etimeh.value)*60 + parseInt(form1.etimem.value);
		ttime =parseInt(etime) - parseInt(stime) ;
		if (stime >= etime)
			errmsg += "開始時間必須小於結束時間\n";
		if (form1.item.value=="")
			errmsg += "補充教材不能為空白\n";

	<%end if%>
	
    if (errmsg=="")
	{
        form1.summin.value = ttime;
		return true;
	}
    else
    {
        alert(errmsg);
        return false;
    }
}
function replaceString(string, from, to)
{
	var i = string.indexOf(from);
	if (i == -1)
		return string;  //base case
	else
		return(string.substring(0,i-1) + to + replaceString(string.substring(i+1,string.length-1),from,to));
}
function ChkStudent()
{
	vWinCal2 = window.open("lib/checkstudent.asp?sid="+form1.sid.value+"&languagecode=E","iframe_query", "width=200,height=50,status=no,resizable=no,scrollbars=no");
	vWinCal2.opener = form1;
}
function delete_record(vid)
{
    if (confirm("您確定要取消該學嗎預約資料嗎?"))
    {
        form1.validate.value="delete";
        form1.submit();
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center"><%if category="T" then response.write "預約補充教材" else response.write "預約自學軟體療程" end if%></TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="booksoftwareedit.asp" onsubmit="return check_input();">
			<input type="hidden" value="edit" name="validate">
			<input type="hidden" id="id" name="id"  value="<%=id%>">
			<input type="hidden" value="<%=item%>"  name="item" >
			<input type="hidden" value="<%=software%>"  name="software" >
			<input type="hidden" value="<%=floor%>"  name="floor" >
			<input type="hidden" value="<%=tid%>"  name="tid" >
			<input type="hidden" value="<%=category%>"  name="category" >
			<input type="hidden" value="<%=sender%>"  name="sender" >
			<input type="hidden" value="<%=stimeh%>"  name="stimeh" >
			<input type="hidden" value="<%=stimem%>"  name="stimem" >
			<input type="hidden" value="<%=etimeh%>"  name="etimeh" >
			<input type="hidden" value="<%=etimem%>"  name="etimem" >
			
			<input type="hidden"  name="summin" value="<%=summin%>">
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>學號：</TD>
						<TD>姓名：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=sid%>" maxlength="25" size="35" onblur="ChkStudent()"  name="sid" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly >
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
						<TD>學制：</TD>
						<TD>系所：</TD>
						<TD>年級：</TD>
						<TD>班級：</TD>
						<TD></TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=slevel%>" maxlength="10"   name="slevel" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						<TD>
						<input type="text" value="<%=department%>" maxlength="10"  name="department" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						<TD>
						<input type="text" value="<%=grade%>" maxlength="10" size="10"  name="grade" class="inputtext"  style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						
						<TD>
						<input type="text" value="<%=class1%>" maxlength="25"  name="class1" class="inputtext"  style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
		
			
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>預約日期：</TD>
						<TD>預約時間起迄：</TD>
					</TR>
					<TR>
						
						<TD valign="top">
						<input type="text" value="<%=bdate%>" maxlength="25" size="15"  name="bdate" class="inputtext" readonly>
						<!-- <img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('bdate')" class="showhand">&nbsp; -->
						</TD>
						<TD valign="top">
						<TABLE cellSpacing=1  cellPadding=2  border=0 >
						<TR><TD><%=stimeh%>
							<!-- <select name="stimeh" class="inputtext" >
							<option value="">時</option>
							<option value="8" <%if stimeh="08" or stimeh="8" then response.write "selected" end if %>>8</option>
							<option value="9" <%if stimeh="09"  or stimeh="9" then response.write "selected" end if %>>9</option>
							<option value="10" <%if stimeh="10" then response.write "selected" end if%>>10</option>
							<option value="11" <%if stimeh="11" then response.write "selected" end if%>>11</option>
							<option value="12" <%if stimeh="12" then response.write "selected" end if%>>12</option>
							<option value="13" <%if stimeh="13" then response.write "selected" end if%>>13</option>
							<option value="14" <%if stimeh="14" then response.write "selected" end if%>>14</option>
							<option value="15" <%if stimeh="15" then response.write "selected" end if%>>15</option>
							<option value="16" <%if stimeh="16" then response.write "selected" end if%>>16</option>
							<option value="17" <%if stimeh="17" then response.write "selected" end if %>>17</option>
							</select> -->
						</TD><TD class="inputlabel">&nbsp;:&nbsp;</TD>
						<TD><%=stimem%>
							<!-- <select name="stimem" class="inputtext" >
							<option value="">分</option>
							<option value="0" <%if stimem="00" or stimem="0" then response.write "selected" end if%>>00</option>
							<option value="10" <%if stimem="10" then response.write "selected" end if %>>10</option>
							<option value="20" <%if stimem="20" then response.write "selected" end if %>>20</option>
							<option value="30" <%if stimem="30" then response.write "selected" end if%>>30</option>
							<option value="40" <%if stimem="40" then response.write "selected" end if%>>40</option>
							<option value="50" <%if stimem="50" then response.write "selected" end if%>>50</option>
							
							</select> -->
						</TD>
						<TD class="inputlabel">&nbsp;~&nbsp;</TD>
						<TD><%=etimeh%>
							<!-- <select name="etimeh" class="inputtext" >
							<option value="">時</option>
							<option value="8" <%if etimeh="08" or etimeh="8"  then response.write "selected" end if %>>8</option>
							<option value="9" <%if etimeh="09" or etimeh="9"  then response.write "selected" end if %>>9</option>
							<option value="10" <%if etimeh="10" then response.write "selected" end if%>>10</option>
							<option value="11" <%if etimeh="11" then response.write "selected" end if%>>11</option>
							<option value="12" <%if etimeh="12" then response.write "selected" end if%>>12</option>
							<option value="13" <%if etimeh="13" then response.write "selected" end if%>>13</option>
							<option value="14" <%if etimeh="14" then response.write "selected" end if%>>14</option>
							<option value="15" <%if etimeh="15" then response.write "selected" end if%>>15</option>
							<option value="16" <%if etimeh="16" then response.write "selected" end if%>>16</option>
							<option value="17" <%if etimeh="17" then response.write "selected" end if %>>17</option>
							</select> -->
						</TD><TD class="inputlabel">&nbsp;:&nbsp;</TD>
						<TD><%=etimem%>
							<!-- <select name="etimem" class="inputtext" >
							<option value="">分</option>
							<option value="0" <%if etimem="00"  or etimem="0"   then response.write "selected" end if%>>00</option>
							<option value="10" <%if etimem="10" then response.write "selected" end if %>>10</option>
							<option value="20" <%if etimem="20" then response.write "selected" end if %>>20</option>
							<option value="30" <%if etimem="30" then response.write "selected" end if%>>30</option>
							<option value="40" <%if etimem="40" then response.write "selected" end if%>>40</option>
							<option value="50" <%if etimem="50" then response.write "selected" end if%>>50</option>
							
							</select> -->
						</TD>
						
						</TR>
						</TABLE>
						</TD>
						
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>預約狀態：</TD>
						<TD>&nbsp;&nbsp;取消日期：</TD>
						<TD>&nbsp;&nbsp;取消人員：</TD>
					</TR>
					<TR>
					<input type="hidden" value="<%=yn%>"  name="yn" >
						<TD>	&nbsp;&nbsp;<%=replace(replace(yn,"Y","<font color=""blue"">已預約</font>"),"N","<font color=""red"">取消</font>")%></TD>
						<TD>&nbsp;&nbsp;<%=canceldate%></TD>
						<TD>&nbsp;&nbsp;<%=canceluid%></TD>
						
					</TR>
				</TABLE>
			</TD></TR>
			<%if category="S" then%>
			<TR ><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>軟體：</TD>
					</TR>
				
					<TR>
						<TD><%=floor%>&nbsp;-&nbsp;<%=software%>	</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<%
			else
					
			
			%>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>補充教材：</TD>
					</TR>
					<TR>
						<TD>
							<%=software%>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<%end if%>
			
			<TR>
			<TD>
			<BR>
			<input  type="submit" value="修改後儲存"  class="inputbutton"  disabled>
			<input  type="button" value="取消預約" <%=btnstatus%>   class="inputbutton" onclick="delete_record();">
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

<!-- #INCLUDE virtual="/include/lib/conn/syconnclose.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->
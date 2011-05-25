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
nextrec=trim(request("nextrec"))
category=trim(request("category")) '老師或小老師
bdate=trim(request("bdate"))
stimeh=trim(request("stimeh"))
stimem=trim(request("stimem"))
etimeh=trim(request("etimeh"))
etimem=trim(request("etimem"))
itemc=trim(request("itemc"))
sid=trim(request("sid"))
name=trim(request("name"))
slevel=trim(request("slevel"))
grade=trim(request("grade"))
class1=trim(request("class1"))
department=trim(request("department"))
score=trim(request("score"))
summin=trim(request("summin"))

'response.write "itemc" & itemc
'response.end
set rs = server.CreateObject("adodb.recordset")
set rs2 = server.CreateObject("adodb.recordset")

btnstatus=""


'response.write "btime=" & btime
if category = "T" then
	sender=ifnull(trim(request("sender")),"booksoftwarelist.asp?category=T")
else
	sender=ifnull(trim(request("sender")),"booksoftwarelist.asp?category=S")
end if
if validate="add" then
	checkflag = "true"
	tmpweek = weekday(NumberToDateFormat(bdate)) 
	if  tmpweek="1" or tmpweek="7" then
		showmessage="星期六和星期日沒有開放服務，無法預約。"
		checkflag = "false"
	end if
	if  (cdbl(bdate) > cdbl(datetoNumformat(dateadd("d",14,date())))  or  cdbl(bdate) <= cdbl(datetoNumformat(date())) ) and session("classify")="S" then
		showmessage="只開放14天內之預約，且不能預約當天。"
		checkflag = "false"
	end if

	if checkflag ="true" then
		if category="S" then
			'預約軟體
			
			'檢查同天是否有預約其它時間,一天 不得超過二個小時
			sql = "select   isnull(sum(summin),0) as summin  from  boo_book_software where  bdate='"&bdate&"' and sid='"&sid&"'  and  yn='Y' "
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

				sql = "select * from boo_book_software where 1=0 "
				rs.Open sql,msconn,adOpenStatic,adLockOptimistic
				if rs.EOF then
					rs.AddNew
				'	id= getguid()
				'	if id<>"" then
				'		rs("id")=id
			'		end if
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
					if itemc<>"" then
						rs("item")=itemc
					end if
					if sid<>"" then
						rs("sid")=sid
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
					if category<>"" then
						rs("category")=category
					end if		
					rs("yn") ="Y"
					rs("initdate") = date()
					rs("inituid") = session("sid")
			
			
					rs.Update
					if Err.Number=0 then 
						'response.redirect "booksoftwarelist.asp?category=" & category
						showmessage="新增成功。"
						bdate=""
						stimeh=""
						stimem=""
						etimeh=""
						etimem=""
						itemc=""
						sid=""
						name=""
						slevel=""
						grade=""
						class1=""
						department=""
						score=""
						summin=""
					else
						showmessage= Err.Description
					end if

				else
					showmessage="此時段已有人預約。"
				end if

				rs.close
			end if
		else
			'預約補充教材
			sql = "select * from boo_book_software where 1=0 "
				rs.Open sql,msconn,adOpenStatic,adLockOptimistic
				if rs.EOF then
					rs.AddNew
				'	id= getguid()
				'	if id<>"" then
				'		rs("id")=id
			'		end if
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
					if itemc<>"" then
						rs("item")=itemc
					end if
					if sid<>"" then
						rs("sid")=sid
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
					if category<>"" then
						rs("category")=category
					end if		
					rs("yn") ="Y"
					rs("initdate") = date()
					rs("inituid") = session("sid")
			
			
					rs.Update
					if Err.Number=0 then 
						'response.redirect "booksoftwarelist.asp?category=" & category
						showmessage="新增成功。"
						bdate=""
						stimeh=""
						stimem=""
						etimeh=""
						etimem=""
						itemc=""
						sid=""
						name=""
						slevel=""
						grade=""
						class1=""
						department=""
						score=""
						summin=""
			
					else
						showmessage= Err.Description
					end if

				else
					showmessage="此時段已有人預約。"
				end if

				rs.close




		end if
	end if 'if checkflag="true" then
end if

if session("classify")="S" then
	sid = session("sid")
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
	var stime=0;
	var etime=0;
	var ttime=0;
	if (form1.sid.value=="")
        errmsg += "學號不能為空白\n";
	if (form1.name.value=="")
        errmsg += "姓名不能為空白\n";
	
    if (form1.stimeh.value=="" || form1.stimem.value=="" || form1.etimeh.value=="" || form1.etimem.value=="")
        errmsg += "請指定時間起迄\n";
	if (form1.bdate.value=="")
        errmsg += "日期不能為空白\n";

	<%if category="S" then%>
	if (form1.item0.checked==true)
        errmsg += "請指定預約項目\n";
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

		if (form1.itemc.value=="")
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

function ChkStudent()
{
	vWinCal2 = window.open("lib/checkstudent_h.asp?sid="+form1.sid.value+"&languagecode=E","iframe_query", "width=200,height=50,status=no,resizable=no,scrollbars=no");
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center"><%if category="T" then response.write "預約補充教材" else response.write "加入學員歷史預約紀錄（自學軟體）" end if%></TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="historysoftware.asp" onsubmit="return check_input();">
			<input type="hidden" value="add" name="validate">
			<input type="hidden" id="nextrec" name="nextrec">
			<input type="hidden" value="<%=category%>"  name="category" >
			<input type="hidden" value=""  name="summin" >
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>學號：</TD>
						<TD>姓名：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=sid%>" maxlength="25" size="35" onblur="ChkStudent()" name="sid" class="inputtext" <%if session("classify")<>"A" then response.write "readonly" end if%>>
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
						<img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('bdate')" class="showhand">&nbsp;
						</TD>
						<TD valign="top">
						<TABLE cellSpacing=0 cellPadding=0  border=0 >
						<TR><TD>
							<select name="stimeh" class="inputtext" >
							<option value="">時</option>
							<option value="8" <%if stimeh="8" then response.write "selected" end if %>>8</option>
							<option value="9" <%if stimeh="9" then response.write "selected" end if %>>9</option>
							<option value="10" <%if stimeh="10" then response.write "selected" end if%>>10</option>
							<option value="11" <%if stimeh="11" then response.write "selected" end if%>>11</option>
							<option value="12" <%if stimeh="12" then response.write "selected" end if%>>12</option>
							<option value="13" <%if stimeh="13" then response.write "selected" end if%>>13</option>
							<option value="14" <%if stimeh="14" then response.write "selected" end if%>>14</option>
							<option value="15" <%if stimeh="15" then response.write "selected" end if%>>15</option>
							<option value="16" <%if stimeh="16" then response.write "selected" end if%>>16</option>
							<option value="17" <%if stimeh="17" then response.write "selected" end if %>>17</option>
							</select>
						</TD><TD class="inputlabel">&nbsp;:&nbsp;</TD>
						<TD>
							<select name="stimem" class="inputtext" >
							<option value="">分</option>
							<option value="0" <%if stimem="0" then response.write "selected" end if%>>00</option>
							<option value="10" <%if stimem="10" then response.write "selected" end if %>>10</option>
							<option value="20" <%if stimem="20" then response.write "selected" end if %>>20</option>
							<option value="30" <%if stimem="30" then response.write "selected" end if%>>30</option>
							<option value="40" <%if stimem="40" then response.write "selected" end if%>>40</option>
							<option value="50" <%if stimem="50" then response.write "selected" end if%>>50</option>
							
							</select>
						</TD>
						<TD class="inputlabel">&nbsp;~&nbsp;</TD>
						<TD>
							<select name="etimeh" class="inputtext" >
							<option value="">時</option>
							<option value="8" <%if etimeh="8" then response.write "selected" end if %>>8</option>
							<option value="9" <%if etimeh="9" then response.write "selected" end if %>>9</option>
							<option value="10" <%if etimeh="10" then response.write "selected" end if%>>10</option>
							<option value="11" <%if etimeh="11" then response.write "selected" end if%>>11</option>
							<option value="12" <%if etimeh="12" then response.write "selected" end if%>>12</option>
							<option value="13" <%if etimeh="13" then response.write "selected" end if%>>13</option>
							<option value="14" <%if etimeh="14" then response.write "selected" end if%>>14</option>
							<option value="15" <%if etimeh="15" then response.write "selected" end if%>>15</option>
							<option value="16" <%if etimeh="16" then response.write "selected" end if%>>16</option>
							<option value="17" <%if etimeh="17" then response.write "selected" end if %>>17</option>
							</select>
						</TD><TD class="inputlabel">&nbsp;:&nbsp;</TD>
						<TD>
							<select name="etimem" class="inputtext" >
							<option value="">分</option>
							<option value="0" <%if etimem="0" then response.write "selected" end if%>>00</option>
							<option value="10" <%if etimem="10" then response.write "selected" end if %>>10</option>
							<option value="20" <%if etimem="20" then response.write "selected" end if %>>20</option>
							<option value="30" <%if etimem="30" then response.write "selected" end if%>>30</option>
							<option value="40" <%if etimem="40" then response.write "selected" end if%>>40</option>
							<option value="50" <%if etimem="50" then response.write "selected" end if%>>50</option>
							</select>
						</TD>
						</TR>
						</TABLE>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<%if category="S" then%>
			<TR ><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>軟體：<font color="#FF0000">※預約軟體不受處方籤控制</font></TD>
					</TR>
					<%
					'軟體
					set rsLoad = server.CreateObject("adodb.recordset")
					sql ="select * from boo_software where yn='Y' and floor='2F' order by floor"
					rsLoad.Open sql,msconn,adOpenStatic,adLockReadonly
					Str=""
					i = 1 
					while not rsLoad.EOF
						if i=0 then
							Str=Str& "<TR>"
						end if
						if i >=3 then
							Str=Str& "</TR>"
							i = 0
						end if
						i = i + 1 
						if itemc=rsLoad("id") then
							Str=Str&"<TD>&nbsp;<input type=radio name='itemc' value="""&rsLoad("id")&""" checked></TD><TD>&nbsp;"&  rsLoad("software")&"</TD>"
						else
							Str=Str&"<TD>&nbsp;<input type=radio name='itemc' value="""&rsLoad("id")&""" ></TD><TD>&nbsp;"&  rsLoad("software") &"</TD>"
						end if 
						rsLoad.MoveNext 
					wend
					rsLoad.close
					sql ="select * from boo_software where yn='Y' and floor='3F' order by floor"
					rsLoad.Open sql,msconn,adOpenStatic,adLockReadonly
					Str3=""
					i = 0 
					while not rsLoad.EOF
						if i=0 then
							Str3=Str3& "<TR>"
						end if
						if i >=3 then
							Str3=Str3& "</TR>"
							i = 0
						end if
						i = i + 1 
						if itemc=rsLoad("id") then
							Str3=Str3&"<TD>&nbsp;<input type=radio name='itemc' value="""&rsLoad("id")&""" checked></TD><TD>&nbsp;"&  rsLoad("software") &"</TD>"
						else
							Str3=Str3&"<TD>&nbsp;<input type=radio name='itemc' value="""&rsLoad("id")&""" ></TD><TD>&nbsp;"&  rsLoad("software") &"</TD>"
						end if 
						rsLoad.MoveNext 
					wend
					rsLoad.close
					set rsLoad=nothing
					
					%>
					<TR>
						<TD>
						2F 區域
						</TD>
					</TR>
					<TR>
						<TD>
						<TABLE cellSpacing=0 cellPadding=0  border=0 >
						<TR>
						<TD>&nbsp;<input type=radio id="item0" name='itemc' value="" checked></TD><TD>未指定</TD>
						<%=Str%>
						
						</TABLE>
						</TD>
					</TR>
					<TR>
						<TD>
						3F 區域
						</TD>
					</TR>
					<TR>
						<TD>
						<TABLE cellSpacing=0 cellPadding=0  border=0 >
						<%=Str3%>
						</TABLE>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<%
			else
					set rsLoad = server.CreateObject("adodb.recordset")
					sql ="select * from boo_software where yn='Y' and category='T' order by software"
					rsLoad.Open sql,msconn,adOpenStatic,adLockReadonly
					StrLearn=""
					if rsLoad.state then	
						while not rsLoad.eof
							if itemc=rsLoad("id") then
								StrLearn=StrLearn&"<option selected value='"&rsLoad("id")&"' >"  & rsLoad("software")&"</option>"
							else
								StrLearn=StrLearn&"<option value='"&rsLoad("id")&"' >"  &rsLoad("software")&"</option>"
							end if
							rsLoad.movenext
						wend
					end if
					rsLoad.close
			
			%>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>補充教材：</TD>
					</TR>
					<TR>
						<TD>
							<select name="itemc" class="inputtext" >
							<option value=""> - 請指定 -</option>
							<%=StrLearn%>
							</select>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<%end if%>
			<TR>
			<TD>
			<BR>
			<input  type="submit" value="加入歷史預約" <%=btnstatus%> class="inputbutton" >
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
<%
if 1=1 then
%>
<SCRIPT LANGUAGE=javascript>
<!--
	if (form1.sid.value!="")
		ChkStudent();
//-->
</SCRIPT>
<%
end if
%>


<!-- #INCLUDE virtual="/include/lib/conn/syconnclose.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->
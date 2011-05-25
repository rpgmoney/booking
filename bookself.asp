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

itemtimes = trim(request("itemtimes"))'該項目可預約時間

'response.write "itemtimes=" & itemtimes
'response.end

category=trim(request("category")) '老師或小老師
bdate=trim(request("bdate"))
stimeh=trim(request("stimeh"))
stimem=trim(request("stimem"))
etimeh=trim(request("etimeh"))
etimem=trim(request("etimem"))
item1=trim(request("item1"))
sid=trim(request("sid"))
name=trim(request("name"))
slevel=trim(request("slevel"))
grade=trim(request("grade"))
class1=trim(request("class1"))
department=trim(request("department"))
score=trim(request("score"))
summin=trim(request("summin"))
tid=trim(request("tid"))
'response.write "summin" & summin
'response.end
set rs = server.CreateObject("adodb.recordset")
set rs2 = server.CreateObject("adodb.recordset")

btnstatus=""


'response.write "btime=" & btime
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
			
		'檢查同天是否有預約其它時間,一天 不得超過二個小時
		sql = "select   isnull(sum(summin),0) as summin  from  boo_book_software where  bdate='"&bdate&"' and sid='"&sid&"'  and  yn='Y' and category='F' and item='"&item1&"' "
		'response.write sql
		'response.end
		
		rs.Open sql,msconn,adOpenStatic,adLockReadonly
		if not  rs.EOF then
			tmpsum = rs("summin")
			'response.write "summin" & summin & "<br>"  
			'response.write "tmpsum" & tmpsum & "<br>"
			'response.write "itemtimes" & itemtimes & "<br>"
			if  ((cdbl(tmpsum) + cdbl(summin))  > (cdbl(itemtimes) )) then
				showmessage = "你在同一天有預約其它時段喔，同一天預約不得超過&nbsp;"&itemtimes/60 & "&nbsp;小時"
				checkflag = "false"
			end if
		end if
		rs.close
		'checkflag="false"
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
				if item1<>"" then
					rs("item")=item1
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
				rs("category")="F"
				
				rs("yn") ="Y"
				rs("initdate") = date()
				rs("inituid") = session("sid")
		
		
				rs.Update
				if Err.Number=0 then 
						response.redirect "bookselflist.asp?category=F"
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

if session("classify")="S" or session("classify")="E" then
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

	if (form1.item1.value=="")
        errmsg += "請指定預約項目\n";
	if (errmsg=="")
	{
		
		stime = parseInt(form1.stimeh.value)*60 + parseInt(form1.stimem.value);
		etime = parseInt(form1.etimeh.value)*60 + parseInt(form1.etimem.value);
		time1 =parseInt(form1.itemtimes.value);//可項約時間
		ttime =parseInt(etime) - parseInt(stime) ;
		if (stime >= etime)
			errmsg += "開始時間必須小於結束時間\n";
		//alert(form1.itemtimes.value);
		//alert(time1);
		//alert(ttime);
		if ( ttime >time1)
			errmsg += "不得大於可預約 "+time1+"分鐘\n";
		
	
	}
	
	
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
	vWinCal2 = window.open("lib/checkstudentfors.asp?sid="+form1.sid.value,"iframe_query", "width=200,height=50,status=no,resizable=no,scrollbars=no");
	vWinCal2.opener = form1;
}
function itemchange()
{
	form1.itemtimes.value=form1.item1.options[form1.item1.selectedIndex].dname;
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">預約自學療程</TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="bookself.asp" onsubmit="return check_input();">
			<input type="hidden" value="add" name="validate">
			<input type="hidden" id="nextrec" name="nextrec">
			<input type="hidden" value="<%=itemtimes%>"  name="itemtimes" >
			<input type="hidden" value="<%=score%>"  name="score" >
			<input type="hidden" value="<%=tid%>"  id="tid"  name="tid" >
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
							<select name="stimeh" class="inputtext"  >
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
							<select name="etimeh" class="inputtext"  >
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
			<TR ><TD>
			<%
					set rsLoad = server.CreateObject("adodb.recordset")
					sql ="select * from boo_self_item where yn='Y'  order by item"
					rsLoad.Open sql,msconn,adOpenStatic,adLockReadonly
					StrLearn=""
					if rsLoad.state then	
						while not rsLoad.eof
							if item1=rsLoad("id") then
								StrLearn=StrLearn&"<option selected value='"&rsLoad("id")&"' dname='"&(rsLoad("hours")*60)&"'>"  & rsLoad("item")&"</option>"
							else
								StrLearn=StrLearn&"<option value='"&rsLoad("id")&"' dname='"&(rsLoad("hours")*60)&"'>"  &rsLoad("item")&"</option>"
							end if
							rsLoad.movenext
						wend
					end if
					rsLoad.close
			
			%>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR class="inputlabel">
						<TD>項目：</TD>
					</TR>

					<TR>
						<TD>
						<select name="item1" class="inputtext" onchange="itemchange();">
						<option value=""> - 請指定項目 - </option>
						<%=StrLearn%>
						</select>
						</TD>
					</TR>
				
				</TABLE>
			</TD></TR>

			<TR>
			<TD>
			<BR>
			<input  type="submit" value="預約" <%=btnstatus%> class="inputbutton" >
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


<%

showflag = trim(request("showflag"))
if showflag="1" then
	set rsLoad = server.CreateObject("adodb.recordset")
	sql = "select * from boo_parameter where ID = 'A'"
	rsLoad.Open sql,msconn,adOpenStatic,adLockReadOnly
	if not rsLoad.EOF then
		showhint = trim(rsLoad("showhint"))
		
	end if
	rsLoad.close
	if showhint="Y"  then
	%>
	<script language="javascript">
	function window.onload()
	{
		var ls_parm = 'dialogWidth=650px;'
						+ 'dialogHeight=650px;'
						+ 'center=yes;'
						+ 'border=thin;'
						+ 'help=no;'
						+ 'directories=no;'
						+ 'location=no;'
						+ 'status=no';
		window.open('showrule.asp','預約規則','fullscreen=1,scrollbars=1');
	}
	</script>
	<%end if%>
<%end if%>
<!-- #INCLUDE virtual="/include/lib/conn/syconnclose.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->
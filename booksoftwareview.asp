<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
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

'response.write "summin" & summin
'response.end
set rs = server.CreateObject("adodb.recordset")
set rs2 = server.CreateObject("adodb.recordset")

btnstatus=""

'
'response.write "btime=" & btime

sender=ifnull(trim(request("sender")),"booksoftwarelist.asp")


sql = "select * from boo_book_software where id='"&id&"' "
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


end if
rs.close






%>
<HTML>
<HEAD>
<TITLE> 【LDCC英外語能力診斷輔導中心】 </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>

</HEAD>
<BODY bottomMargin=0 leftMargin=0 topMargin=0 rightMargin=0 marginheight="0" marginwidth="0"  bgcolor="#f4c60d">
<TABLE cellSpacing=1 cellPadding=0 width="100%"  height="100%" align="center"  border=0>
<TR><TD>
<TABLE cellSpacing=0 cellPadding=0 width="100%"  height="100%" align="center" bgColor=#ffffff border=0>
<TR height="15" bgcolor="#333333">
	<TD align="center">
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
						<input type="text" value="<%=sid%>" maxlength="25" size="35"  name="sid" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly >
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
							</select>
						</TD><TD class="inputlabel">&nbsp;:&nbsp;</TD>
						<TD>
							<select name="stimem" class="inputtext" >
							<option value="">分</option>
							<option value="0" <%if stimem="00" or stimem="0" then response.write "selected" end if%>>00</option>
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
							</select>
						</TD><TD class="inputlabel">&nbsp;:&nbsp;</TD>
						<TD>
							<select name="etimem" class="inputtext" >
							<option value="">分</option>
							<option value="0" <%if etimem="00"  or etimem="0"   then response.write "selected" end if%>>00</option>
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
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>預約狀態：</TD>
						<TD>取消日期：</TD>
						<TD>取消人員：</TD>
					</TR>
					<TR>
						<TD>	&nbsp;&nbsp;<%=replace(replace(yn,"Y","<font color=""blue"">已預約</font>"),"N","<font color=""red"">取消</font>")%></TD>
						<TD><%=canceldate%></TD>
						<TD><%=canceluid%></TD>
					</TR>
				</TABLE>
			</TD></TR>
			<%if category="S" then%>
			<TR ><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>軟體：</TD>
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
						if item=rsLoad("id") then
							Str=Str&"<TD>&nbsp;<input type=radio name='item' value="""&rsLoad("id")&""" checked></TD><TD>&nbsp;"&  rsLoad("software") &"</TD>"
						else
							Str=Str&"<TD>&nbsp;<input type=radio name='item' value="""&rsLoad("id")&""" ></TD><TD>&nbsp;"&  rsLoad("software") & "</TD>"
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
						if item=rsLoad("id") then
							Str3=Str3&"<TD>&nbsp;<input type=radio name='item' value="""&rsLoad("id")&""" checked></TD><TD>&nbsp;"&  rsLoad("software") & "</TD>"
						else
							Str3=Str3&"<TD>&nbsp;<input type=radio name='item' value="""&rsLoad("id")&""" ></TD><TD>&nbsp;"&  rsLoad("software") & "</TD>"
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
						<TD>&nbsp;<input type=radio id="item0" name='item' value="" checked></TD><TD>未指定</TD>
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
							if item=rsLoad("id") then
								StrLearn=StrLearn&"<option selected value="""&rsLoad("id")&""" >"  & rsLoad("software")&"</option>"
							else
								StrLearn=StrLearn&"<option value="""&rsLoad("id")&""" >"  &rsLoad("software")&"</option>"
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
							<select name="item" class="inputtext" >
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
<TR bgcolor="#333333" height="15">
	<TD class="T1">
	
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
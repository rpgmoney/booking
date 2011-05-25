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
id=trim(request("id"))
category=trim(request("category")) 

subject=trim(request("subject")) 
date1=trim(request("date1")) 
stimeh=trim(request("stimeh")) 
etimeh=trim(request("etimeh")) 
stimem=trim(request("stimem")) 
etimem=trim(request("etimem")) 
place=trim(request("place")) 
speaker=trim(request("speaker")) 
content=trim(request("content")) 
sdate=trim(request("sdate")) 
edate=trim(request("edate")) 
class1=trim(request("class1")) 


sender=ifnull(trim(request("sender")),"lecture.asp?category=" & category)

set rs = server.CreateObject("adodb.recordset")
if validate="edit" then
	sql = "select * from boo_lecture where id='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
    if  not rs.EOF then
		if subject<>"" then
            rs("subject")=subject
        end if
		if date1<>"" then
            rs("date1")=date1
        end if
		if stimeh<>"" and stimem<>""  then
            rs("stime")=stimeh+stimem
        end if
		if etimeh<>"" and etimem<>""  then
            rs("etime")=etimeh+etimem
        end if
		if place<>"" then
            rs("place")=place
        end if
		if speaker<>"" then
            rs("speaker")=speaker
        end if
		if content<>"" then
            rs("content")=content
        end if
		if sdate<>"" then
            rs("sdate")=sdate
        end if
		if edate<>"" then
            rs("edate")=edate
        end if
		if category<>"" then
            rs("category")=category
        end if
		if class1<>"" then
            rs("class1")=class1
        else
			rs("class1")=null
		end if
		rs("initdate") = date()
		rs("inituid") = session("sid")

		rs.Update
        if Err.Number=0 then 
			if nextrec="Y" then
				validate=""
				nextrec=""
				subject=""
				date1=""
				stime=""
				etime=""
				place=""
				speaker=""
				content=""
				sdate=""
				edate=""
				category=""

			else
				response.redirect sender
			end if
		else
			showmessage= Err.Description
		end if

	else
		showmessage="找不到該筆資料。"
	end if

	rs.close
else

	sql = "select * from boo_lecture where id='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
	if rs.EOF then
		response.redirect sender
	else
		id=trim(rs("id"))
		subject=trim(rs("subject")) 
		date1=trim(rs("date1")) 
		stimeh=left(trim(rs("stime")) ,2)
		etimeh=left(trim(rs("etime")) ,2)
		stimem=right(trim(rs("stime")) ,2)
		etimem=right(trim(rs("etime")) ,2)
		place=trim(rs("place")) 
		speaker=trim(rs("speaker")) 
		content=trim(rs("content")) 
		sdate=trim(rs("sdate")) 
		edate=trim(rs("edate")) 
		class1=trim(rs("class1")) 
	end if
	rs.close
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
	
    if (form1.subject.value=="")
        errmsg += "名稱不能為空白\n";
	if (form1.date1.value=="")
        errmsg += "日期不能為空白\n";
	if (form1.stimeh.value=="" || form1.stimem.value=="" || form1.etimeh.value=="" || form1.etimem.value=="")
        errmsg += "起迄時間不能為空白\n";
	if (form1.place.value=="")
        errmsg += "地點不能為空白\n";
	if (form1.speaker.value=="")
        errmsg += "主講人/老師不能為空白\n";
	if (form1.content.value=="")
        errmsg += "講座內容/課程內容不能為空白\n";
	if (form1.sdate.value=="")
        errmsg += "開始報名日期不能為空白\n";
		if (form1.edate.value=="")
        errmsg += "結束報名日期不能為空白\n";

    if (errmsg=="")
        return true;
    else
    {
        alert(errmsg);
        return false;
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center"><%if category="L" then response.write "編輯外語學習講座資料"  else response.write "編輯處方課程"  end if%></TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="lectureedit.asp" onsubmit="return check_input();">
			<input type="hidden" value="edit" name="validate">
			<input type="hidden" id="category" name="category" value="<%=category%>">
			<input type="hidden" id="id" name="id" value="<%=id%>">
			<input type="hidden" id="sender" name="sender" value="<%=sender%>">
			
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD><%if category="L" then response.write "外語學習講座名稱"  else response.write "處方課程名稱"  end if%>：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=subject%>" maxlength="50" size="65" name="subject" class="inputtext" >
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>日期：</TD>
						<TD>時間起迄：</TD>
						<TD></TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=date1%>" maxlength="25"  name="date1" class="inputtext" readonly>
						<img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('date1')" class="showhand">
						&nbsp;
						</TD>
						<TD>
							<TABLE cellSpacing=0 cellPadding=0  border=0 >
							<TR>
							<TD class="inputlabel">
								<select name="stimeh" class="inputtext">
								<option value=""> 時 </option>
								<optgroup label="上午">
								<option value="08" <%if stimeh="08" then response.write "selected" end if%>>08</option>
								<option value="09" <%if stimeh="09" then response.write "selected" end if%>>09</option>
								<option value="10" <%if stimeh="10" then response.write "selected" end if%>>10</option>
								<option value="11" <%if stimeh="11" then response.write "selected" end if%>>11</option>
								<option value="12" <%if stimeh="12" then response.write "selected" end if%>>12</option>
								</optgroup>
								<optgroup label="下午">
								<option value="13" <%if stimeh="13" then response.write "selected" end if%>>13</option>
								<option value="14" <%if stimeh="14" then response.write "selected" end if%>>14</option>
								<option value="15" <%if stimeh="15" then response.write "selected" end if%>>15</option>
								<option value="16" <%if stimeh="16" then response.write "selected" end if%>>16</option>
								<option value="17" <%if stimeh="17" then response.write "selected" end if%>>17</option>
								</optgroup>
								</select>
								<select name="stimem" class="inputtext">
								<option value=""> 分 </option>
								<option value="00"  <%if stimem="00" then response.write "selected" end if%>>00</option>
								<option value="10"  <%if stimem="10" then response.write "selected" end if%>>10</option>
								<option value="20"  <%if stimem="20" then response.write "selected" end if%>>20</option>
								<option value="30"  <%if stimem="30" then response.write "selected" end if%>>30</option>
								<option value="40"  <%if stimem="40" then response.write "selected" end if%>>40</option>
								<option value="50"  <%if stimem="50" then response.write "selected" end if%>>50</option>
								</select>
							</TD>
							<TD class="inputlabel">&nbsp;~&nbsp;</TD>
							<TD>
								<select name="etimeh" class="inputtext">
								<option value=""> 時 </option>
								<optgroup label="上午">
								<option value="08"  <%if etimeh="08" then response.write "selected" end if%>>08</option>
								<option value="09"  <%if etimeh="09" then response.write "selected" end if%>>09</option>
								<option value="10"  <%if etimeh="10" then response.write "selected" end if%>>10</option>
								<option value="11"  <%if etimeh="11" then response.write "selected" end if%>>11</option>
								<option value="12"  <%if etimeh="12" then response.write "selected" end if%>>12</option>
								</optgroup>
								<optgroup label="下午">
								<option value="13" <%if etimeh="13" then response.write "selected" end if%>>13</option>
								<option value="14" <%if etimeh="14" then response.write "selected" end if%>>14</option>
								<option value="15" <%if etimeh="15" then response.write "selected" end if%>>15</option>
								<option value="16" <%if etimeh="16" then response.write "selected" end if%>>16</option>
								<option value="17" <%if etimeh="17" then response.write "selected" end if%>>17</option>
								</optgroup>
								</select>
								<select name="etimem" class="inputtext">
								<option value=""> 分 </option>
								<option value="00" <%if etimem="00" then response.write "selected" end if%>>00</option>
								<option value="10" <%if etimem="10" then response.write "selected" end if%>>10</option>
								<option value="20" <%if etimem="20" then response.write "selected" end if%>>20</option>
								<option value="30" <%if etimem="30" then response.write "selected" end if%>>30</option>
								<option value="40" <%if etimem="40" then response.write "selected" end if%>>40</option>
								<option value="50" <%if etimem="50" then response.write "selected" end if%>>50</option>
								</select>
							</TD>
							</TR>
							</TABLE>
						</TD>
						<TD>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>地點：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=place%>" maxlength="100" size="55" name="place" class="inputtext" >
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD><%if category="L" then response.write "主講人"  else response.write "老師"  end if%>：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=speaker%>" maxlength="50" size="55" name="speaker" class="inputtext" >
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<%if category="C" then%>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>班別：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=class1%>" maxlength="50" size="55" name="class1" class="inputtext" >
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<%end if%>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD><%if category="L" then response.write "講座內容"  else response.write "課程內容"  end if%>：</TD>
					</TR>
					<TR>
						<TD>
						<textarea name="content" cols="80" rows="6" class="inputtext" ><%=content%></textarea>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						
						<TD>開始報名日期：</TD>
						<TD>結束報名日期：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=sdate%>" maxlength="25"  name="sdate" class="inputtext" readonly>
						<img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('sdate')" class="showhand">
						&nbsp;
						</TD>
						<TD>
						<input type="text" value="<%=edate%>" maxlength="25"  name="edate" class="inputtext" readonly>
						<img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('edate')" class="showhand">
						&nbsp;
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			
			<TR>
			<TD>
			<BR>
			<input  type="submit" value="儲存" class="inputbutton" >
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

</BODY>
</HTML>
<!-- #INCLUDE virtual="/include/lib/conn/syconnclose.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->
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
if validate="add" then
	sql = "select * from boo_lecture where category='"&category&"' and subject='"&subject&"' "
	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
    if rs.EOF then
        rs.AddNew
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
		showmessage="��ƭ��СC"
	end if

	rs.close
	
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

function check_input()
{
    var errmsg=""
	
    if (form1.subject.value=="")
        errmsg += "�W�٤��ର�ť�\n";
	if (form1.date1.value=="")
        errmsg += "������ର�ť�\n";
	if (form1.stimeh.value=="" || form1.stimem.value=="" || form1.etimeh.value=="" || form1.etimem.value=="")
        errmsg += "�_���ɶ����ର�ť�\n";
	if (form1.place.value=="")
        errmsg += "�a�I���ର�ť�\n";
	if (form1.speaker.value=="")
        errmsg += "�D���H/�Ѯv���ର�ť�\n";
	if (form1.content.value=="")
        errmsg += "���y���e/�ҵ{���e���ର�ť�\n";
	if (form1.sdate.value=="")
        errmsg += "�}�l���W������ର�ť�\n";
		if (form1.edate.value=="")
        errmsg += "�������W������ର�ť�\n";

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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center"><%if category="L" then response.write "�s�W�~�y�ǲ����y���"  else response.write "�s�W�B��ҵ{"  end if%></TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="lectureadd.asp" onsubmit="return check_input();">
			<input type="hidden" value="add" name="validate">
			<input type="hidden" id="nextrec" name="nextrec">
			<input type="hidden" id="category" name="category" value="<%=category%>">
			
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD><%if category="L" then response.write "�~�y�ǲ����y�W��"  else response.write "�B��ҵ{�W��"  end if%>�G</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=subject%>" maxlength="50" size="55" name="subject" class="inputtext" >
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>����G</TD>
						<TD>�ɶ��_���G</TD>
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
								<option value=""> �� </option>
								<optgroup label="�W��">
								<option value="08">08</option>
								<option value="09">09</option>
								<option value="10">10</option>
								<option value="11">11</option>
								<option value="11">12</option>
								</optgroup>
								<optgroup label="�U��">
								<option value="13">13</option>
								<option value="14">14</option>
								<option value="15">15</option>
								<option value="16">16</option>
								<option value="17">17</option>
								</optgroup>
								</select>
								<select name="stimem" class="inputtext">
								<option value=""> �� </option>
								<option value="00">00</option>
								<option value="10">10</option>
								<option value="20">20</option>
								<option value="30">30</option>
								<option value="40">40</option>
								<option value="50">50</option>
								</select>
							</TD>
							<TD class="inputlabel">&nbsp;~&nbsp;</TD>
							<TD>
								<select name="etimeh" class="inputtext">
								<option value=""> �� </option>
								<optgroup label="�W��">
								<option value="08">08</option>
								<option value="09">09</option>
								<option value="10">10</option>
								<option value="11">11</option>
								<option value="11">12</option>
								</optgroup>
								<optgroup label="�U��">
								<option value="13">13</option>
								<option value="14">14</option>
								<option value="15">15</option>
								<option value="16">16</option>
								<option value="17">17</option>
								</optgroup>
								</select>
								<select name="etimem" class="inputtext">
								<option value=""> �� </option>
								<option value="00">00</option>
								<option value="10">10</option>
								<option value="20">20</option>
								<option value="30">30</option>
								<option value="40">40</option>
								<option value="50">50</option>
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
						<TD>�a�I�G</TD>
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
						<TD><%if category="L" then response.write "�D���H"  else response.write "�Ѯv"  end if%>�G</TD>
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
						<TD>�Z�O�G</TD>
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
						<TD><%if category="L" then response.write "���y���e"  else response.write "�ҵ{���e"  end if%>�G</TD>
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
						
						<TD>�}�l���W����G</TD>
						<TD>�������W����G</TD>
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
			<input  type="submit" value="�s�W" class="inputbutton" >
			<input  type="submit" onclick="form1.nextrec.value='Y'" value="�s�W���~��s�W" class="inputbutton">
			<input  type="button" value="��^" class="inputbutton" onclick="window.location='<%=sender%>'">
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
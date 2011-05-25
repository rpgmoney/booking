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
sid=trim(request("sid"))
name=trim(request("name"))
slevel=trim(request("slevel"))
grade=trim(request("grade"))
class1=trim(request("class1"))
department=trim(request("department"))
yn=trim(request("yn"))

subject=trim(request("subject")) 
date1=trim(request("date1")) 
stimeh=ltrim(request("stimeh")) 
etimeh=trim(request("etimeh")) 
stimem=trim(request("stimem"))
etimem=trim(request("etimem"))
place=trim(request("place")) 
speaker=trim(request("speaker")) 
content=trim(request("content")) 
sdate=trim(request("sdate")) 
edate=trim(request("edate")) 

pretest=trim(request("pretest"))
posttest=trim(request("posttest"))

sender=ifnull(trim(request("sender")),"booklecturelist.asp?category=" & category&"&lid=" & id)
set rs = server.CreateObject("adodb.recordset")
if validate="add" then

	sql = "select * from boo_lecture_record where tid='"&id&"' "
	response.write sql
		rs.Open sql,msconn,adOpenStatic,adLockOptimistic
		if rs.EOF then
			rs.AddNew
			if id<>"" then
				rs("tid")=id
			end if
			if pretest<>"" then
				rs("pretest")=pretest
			end if
			if posttest<>"" then
				rs("posttest")=posttest
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
			showmessage="��ƭ��СC"
		end if

		rs.close	
elseif validate="edit" then
	sql = "select * from boo_lecture_record  where tid='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
	if  not rs.EOF then
		if pretest<>"" then
			rs("pretest")=pretest
		else
			rs("pretest")=null
		end if
		if posttest<>"" then
			rs("posttest")=posttest
		else
			rs("posttest")=null
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
		showmessage="�䤣��ӵ���ơC"
	end if

	rs.close	
else

	sql = "select a.*,b.* from boo_book_lecture a left join  boo_lecture b on a.lid=b.id  where a.id='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
	if rs.EOF then
		response.redirect sender
	else
		sid=trim(rs("sid"))
		name=trim(rs("name"))
		slevel=trim(rs("slevel"))
		grade=trim(rs("grade"))
		class1=trim(rs("class1"))
		department=trim(rs("department"))
		yn=trim(rs("yn"))

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
	end if
	rs.close
	'�B��ҵ{����
	sql = "select * from boo_lecture_record   where tid='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
    if rs.EOF then
		validate="add"
	else
		validate="edit"
		pretest=trim(rs("pretest"))
		posttest=trim(rs("posttest"))

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
	
    if (form1.sid.value=="")
        errmsg += "�Ǹ����ର�ť�\n";
	if (form1.name.value=="")
        errmsg += "�m�W���ର�ť�\n";

    if (errmsg=="")
        return true;
    else
    {
        alert(errmsg);
        return false;
    }
}


function delete_record()
{
    if (confirm("�z�T�w�n�����ӵ����W��?"))
    {
        form1.validate.value="cancel";
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">�ǭ��B��ҵ{�������@</TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="lecturerecordedit.asp" >
			<input type="hidden" value="<%=validate%>" name="validate">
			<input type="hidden" id="sid" name="sid" value="<%=sid%>">
			<input type="hidden" id="id" name="id" value="<%=id%>">
			<input type="hidden" id="sender" name="sender" value="<%=sender%>">
			<input type="hidden" id="name" name="name" value="<%=name%>">
			<input type="hidden" id="slevel" name="slevel" value="<%=slevel%>">
			<input type="hidden" id="grade" name="grade" value="<%=grade%>">
			<input type="hidden" id="class1" name="class1" value="<%=class1%>">
			<input type="hidden" id="department" name="department" value="<%=department%>">
			<input type="hidden" id="yn" name="yn" value="<%=yn%>">
			<input type="hidden" id="subject" name="subject" value="<%=subject%>">
			<input type="hidden" id="date1" name="date1" value="<%=date1%>">
			<input type="hidden" id="stimeh" name="stimeh" value="<%=stimeh%>">
			<input type="hidden" id="etimeh" name="etimeh" value="<%=etimeh%>">
			<input type="hidden" id="stimem" name="stimem" value="<%=stimem%>">
			<input type="hidden" id="etimem" name="etimem" value="<%=etimem%>">
			<input type="hidden" id="place" name="place" value="<%=place%>">
			<input type="hidden" id="speaker" name="speaker" value="<%=speaker%>">
			<input type="hidden" id="content" name="content" value="<%=content%>">
			<input type="hidden" id="sdate" name="sdate" value="<%=sdate%>">
			<input type="hidden" id="edate" name="edate" value="<%=edate%>">

			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>�Ǹ��G</TD>
						<TD>�m�W�G</TD>
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
						<TD>�Ǩ�G</TD>
						<TD>�t�ҡG</TD>
						<TD>�~�šG</TD>
						<TD>�Z�šG</TD>
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
						<TD>
						</TD>
					</TR>
				</TABLE>

			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>�e���G</TD>
						<TD>����G</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=pretest%>" maxlength="3" "  name="pretest" class="inputtext" >
						</TD>
						<TD>
						<input type="text" value="<%=posttest%>" maxlength="3"  name="posttest" class="inputtext" >
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR>
			<TD>
			<BR>
			<input  type="submit"   value="�x�s" class="inputbutton"  >
			<input  type="button" value="��^" class="inputbutton" onclick="window.location='<%=sender%>'">
			<BR><BR>
			</TD>
			</TR>
			</form>
			</TABLE>
			<font color="#FF0000">�B��ҵ{��ԲӸ�T�p�U�G</font>
			<TABLE cellSpacing=2 cellPadding=3 width="90%"  border=0   bgcolor="#F3EFE6" style="border: 1px solid #FF66CC; padding: 0">
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR >
						<TD class="inputlabel">�B��ҵ{�W�١G</TD>
						<TD>
						<%=subject%>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR >
						<TD class="inputlabel">����G</TD><TD><%=date1%>&nbsp;</TD>
						<TD class="inputlabel">�ɶ��_���G</TD>
						
							<TD>
							<TABLE cellSpacing=0 cellPadding=0  border=0 >
							<TR>
							<TD>
								<%=stimeh%>��<%=stimem%>��
							</TD>
							<TD class="inputlabel">&nbsp;~&nbsp;</TD>
							<TD>
								<%=etimeh%>��<%=etimem%>��
							</TD>
							</TR>
							</TABLE>
						</TD>
						<TD class="inputlabel">�a�I�G</TD>
						<TD><%=place%></TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR >
						<TD class="inputlabel"><%if category="L" then response.write "�D���H"  else response.write "�Ѯv"  end if%>�G</TD>
						<TD><%=speaker%></TD>
						<TD class="inputlabel">�}�l���W����G</TD>
						<TD><%=sdate%></TD>
						<TD class="inputlabel">�������W����G</TD>
						<TD><%=edate%></TD>
					</TR>
					<TR>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR >
						<TD class="inputlabel" nowrap><%if category="L" then response.write "���y���e"  else response.write "�ҵ{���e"  end if%>�G</TD>
					</TR>
					<TR>
					<TD><%=content%></TD>
					</TR>
				</TABLE>
			</TD></TR>
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
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
btime=trim(request("btime"))
teachername=trim(request("teachername"))
timeflag=trim(request("timeflag"))
sid=trim(request("sid"))
name=trim(request("name"))
slevel=trim(request("slevel"))
grade=trim(request("grade"))
class1=trim(request("class1"))
department=trim(request("department"))
score=trim(request("score"))
ptime=trim(request("ptime"))
languagecode=trim(request("languagecode"))


subject=trim(request("subject"))
content=trim(request("content"))
feedback=trim(request("feedback"))
teacher=trim(request("teacher"))




today=year(date())-1911 & right("0" & month(date()),2) & right("0" & day(date()),2)

sender=ifnull(trim(request("sender")),"consult.asp" )

set dic = server.CreateObject("scripting.dictionary")
dic.Add "1","��"
dic.Add "2","�@"
dic.Add "3","�G"
dic.Add "4","�T"
dic.Add "5","�|"
dic.Add "6","��"
dic.Add "7","��"


set rs = server.CreateObject("adodb.recordset")
set rs2 = server.CreateObject("adodb.recordset")

if validate="add" then
		sql = "select * from boo_write where tid='"&id&"' "
		rs.Open sql,msconn,adOpenStatic,adLockOptimistic
		if rs.EOF then
			rs.AddNew
			if id<>"" then
				rs("tid")=id
			end if
			if content<>"" then
				rs("content")=content
			end if
			if subject<>"" then
				rs("subject")=subject
			end if
			
			if feedback<>"" then
				rs("feedback")=feedback
			end if
			
			if teacher<>"" then
				rs("teacher")=teacher
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
	sql = "select * from boo_write  where tid='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
	if  not rs.EOF then
		if content<>"" then
			rs("content")=content
		else
			rs("content")=null
		end if
		if subject<>"" then
			rs("subject")=subject
		else
			rs("subject")=null
		end if
		
		if feedback<>"" then
			rs("feedback")=feedback
		else
			rs("feedback")=null
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
	'�w����T
	sql = "select * from boo_book_T_M   where id='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
    if rs.EOF then
		response.redirect sender
	else
		bdate=trim(rs("bdate"))
		btime=trim(rs("btime"))
		ptime=trim(rs("ptime"))
		item=trim(rs("item"))
		teachername=trim(rs("teachername"))
		timeflag=trim(rs("timeflag"))
		sid=trim(rs("sid"))
		name=trim(rs("name"))
		slevel=trim(rs("slevel"))
		grade=trim(rs("grade"))
		class1=trim(rs("class1"))
		department=trim(rs("department"))
		score=ifnull(trim(rs("score")),0)
		orallevel=trim(rs("orallevel"))
		oralset=trim(rs("oralset"))
		topic=trim(rs("topic"))
		briefing=trim(rs("briefing"))
		yn=trim(rs("yn"))
		category=trim(rs("category"))
		languagecode=trim(rs("languagecode"))
		signin=trim(rs("signin"))

	end if
	rs.close
	'�԰Ӥ��e��T
	sql = "select * from boo_write   where tid='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
    if rs.EOF then
		validate="add"
	else
		validate="edit"
		did=trim(rs("id"))

		content=trim(rs("content"))
		subject=trim(rs("subject"))
		
		feedback=trim(rs("feedback"))
		teacher=trim(rs("teacher"))

	end if
	rs.close
end if

if teacher="" or isnull(teacher) or isempty(teacher) then
	teacher=teachername
end if

function dateformat(vdate)
	if vdate<>"" then
		if  Cstr(left(vdate,1)) ="9"  then
		dateformat=cStr(cint(left(vdate,2))+1911 ) & "/" & mid(vdate,3,2) & "/" & right(vdate,2)
		else
		dateformat=cStr(cint(left(vdate,3))+1911 ) & "/" & mid(vdate,4,2) & "/" & right(vdate,2)
		end if
	end if
	
end function
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">�ǭ��g�@�������@�s��</TD>
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
						<TD class="inputlabel">�Ǹ��G</TD><TD><%=sid%></TD>
						<TD class="inputlabel">�m�W�G</TD><TD><%=name%></TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR>
						<TD class="inputlabel">�Ǩ�G</TD><TD><%=slevel%></TD>
						<TD class="inputlabel">�t�ҡG</TD><TD><%=department%></TD>
						<TD class="inputlabel">�~�šG</TD><TD><%=grade%></TD>
						<TD class="inputlabel">�Z�šG</TD><TD><%=class1%></TD>
						<TD class="inputlabel">�j�M�^�˦��Z�G</TD><TD><%=score%></TD>
						<TD></TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR>
						<TD class="inputlabel"><%if category = "T" then response.write "�Ѯv" else response.write "�p�Ѯv" end if%>�G</TD><TD><%=teachername%></TD>
						<TD class="inputlabel">&nbsp;�y���O�G&nbsp;</TD><TD><%=languagecode%></TD>
					</TR>
				</TABLE>
			</TD></TR>
			
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR>
						<TD  class="inputlabel">�w�����ءG&nbsp;&nbsp;</TD><TD><%=item%></TD>
						<TD  class="inputlabel">�w������G</TD><TD><%=bdate%></TD>
						<TD  class="inputlabel">�P���G</TD><TD><%="�]&nbsp;"&dic.Item(cstr(cint(weekday(dateformat(bdate)))))&"&nbsp;�^"%></TD>
						<TD  class="inputlabel">�w���ɬq�G</TD><TD><%=btime%></TD>
						<TD><%=replace(replace(replace(timeflag,"U","�W�@�`(25��)"),"B","�U�@�`(25��)"),"A","�W�U�G�`(50��)")%></TD>
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
			<form id="form1" name="form1" method="post" action="writeedit.asp"   onsubmit="return check_input();">
			<input type="hidden" value="<%=validate%>" name="validate">
			<input type="hidden" value="<%=ptime%>" name="ptime">
			<input type="hidden" value="<%=id%>" name="id">
			<input type="hidden" value="<%=yn%>" name="yn">
			<input type="hidden" value="<%=category%>" name="category">
			<input type="hidden" value="<%=sender%>" name="sender">
			<input type="hidden" value="<%=languagecode%>" name="languagecode">
			<input type="hidden" value="<%=sid%>" name="sid">
			<input type="hidden" value="<%=name%>" name="name">
			<input type="hidden" value="<%=bdate%>" name="bdate">
			<input type="hidden" value="<%=btime%>" name="btime">
			<input type="hidden" value="<%=teachername%>" name="teachername">
			<input type="hidden" value="<%=timeflag%>" name="timeflag">
			<input type="hidden" value="<%=slevel%>" name="slevel">
			<input type="hidden" value="<%=grade%>" name="grade">
			<input type="hidden" value="<%=class1%>" name="class1">
			<input type="hidden" value="<%=department%>" name="department">
			<input type="hidden" value="<%=score%>" name="score">
			<input type="hidden" value="<%=ptime%>" name="ptime">
		
			<TABLE class=normal cellSpacing=0 cellPadding=0 width="100%"  border=0>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>�g�@�D�D(Writing Topic)�G</TD>
					</TR>
					<TR>
						<TD><input type="text" value="<%=subject%>" maxlength="100" size="80" name="subject" class="inputtext"></TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>�g�@���D(Writing Problem)�G</TD>
					</TR>
					<TR>
						<TD>
							<TABLE cellSpacing=1 cellPadding=1  border=0 >
							<TR>
							<TD><textarea name="content" rows="5" cols="100" class="inputtext"  ><%=content%></textarea></TD>
							</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>�Ѯv�^�X(feedback)�G</TD>
					</TR>
					<TR>
						<TD>
							<TABLE cellSpacing=1 cellPadding=1  border=0 >
							<TR>
							<TD><textarea name="feedback" rows="5" cols="100" class="inputtext"  ><%=feedback%></textarea></TD>
							</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR class="inputlabel">
						<TD>�n�E�Ѯv�G</TD>
					</TR>
					<TR>
						<TD><input type="text" value="<%=teacher%>" maxlength="25" size="35" name="teacher" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly ></TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD><BR><input  type="submit" value="�x�s" class="inputbutton" ><input  type="button" value="��^" class="inputbutton" onclick="window.location='<%=sender%>'"></TD></TR>
			
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


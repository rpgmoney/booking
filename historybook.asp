<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 40 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<!-- #INCLUDE virtual="/include/lib/asp/jsCalendar.asp" -->
<!-- #INCLUDE file="lib/parameter.inc" -->
<%
validate=trim(request("validate"))
nextrec=trim(request("nextrec"))
category="T" '�Ѯv�Τp�Ѯv
bdate=trim(request("bdate"))
btime=trim(request("btime"))
item=trim(request("item"))
teachername=trim(request("teachername"))
timeflag=trim(request("timeflag"))
sid=trim(request("sid"))
name=trim(request("name"))
slevel=trim(request("slevel"))
grade=trim(request("grade"))
class1=trim(request("class1"))
department=trim(request("department"))
score=trim(request("score"))
orallevel=trim(request("orallevel"))
oralset=trim(request("oralset"))
topic=trim(request("topic"))
briefing=trim(request("briefing"))
languagecode=trim(request("languagecode"))
yms=trim(request("yms"))
'response.write "orallevel" & orallevel
'response.end
set rs = server.CreateObject("adodb.recordset")
set rs2 = server.CreateObject("adodb.recordset")

btnstatus=""
if timeflag="B" then
	ptime=cdbl(btime+25)
else
	ptime=btime 
end if
'
'response.write "btime=" & btime
if category = "T" then
	sender=ifnull(trim(request("sender")),"bookteacher.asp")
	languagecode="E"
else
	sender=ifnull(trim(request("sender")),"booksteacher.asp")
	item="�f�y"
	timeflag="A"
	
'	if languagecode="E" then '�w���^��p�Ѯv�~�n�ˬd

'		sql = "select * from boo_parameter where priority='1'"
'		rs.Open sql,msconn,adOpenStatic,adLockReadonly
'		if not rs.EOF then
			'�H�n�E�Ѯv���Ĥ@����
'			ww = weekday(bdate)
'			sql1 ="select a.*,b.id as bookid,isnull(b.name,'N') as name,b.department,b.grade,b.class1,item,b.btime,b.bdate  from boo_schedule a "
'			sql1 = sql1 & "left join boo_book_T_M b on a.teacher=b.teachername  and b.btime='"&btime&"' and  b.bdate='"&bdate&"' and b.yn='Y' "
'			sql1 = sql1 & " where b.pid is null and a.category='T' and a.btime='"&btime&"' and a.bweek='"&cint(ww)+1&"'  and a.yn='Y' and b.name is null"
			'response.write sql1
'			rs2.Open sql1,msconn,adOpenStatic,adLockReadonly
'			if not rs2.EOF then
				'�n�E�Ѯv���Ŧ�
'				response.write "<script>alert('�������H�n�E�Ѯv���u����ܡA�f�y�p�Ѯv���ĤG����C');</script>"
'				showmessage="�������H�n�E�Ѯv���u����ܡA�f�y�p�Ѯv���ĤG����C<a href=bookteacher.asp?BOOK_DATE="&bdate & "><font color='blue'>&nbsp;&nbsp;<img src='images/u24.gif' border=0>&nbsp;&nbsp;���ܡy�w���n�E�Ѯv�z</font></a>"
'				btnstatus = "disabled"
				'response.redirect "bookteacher.asp?BOOK_DATE="&bdate
'			end if
'			rs2.close
'		end if

'		rs.close

'	end if

end if



if validate="add" then
	
	sql = "select * from boo_book_T_M where bdate='"&bdate&"' and sid='"&sid&"' and item='"&item&"' and yn='Y'"
	rs2.Open sql,msconn,adOpenStatic,adLockReadonly
	if not rs2.EOF then
		showmessage = "�C�H�@�ѬۦP���{���w���@���A�Фŭ��йw���C�ԲӸ�T�аѦҹw���W�h&nbsp;<a href='#' title='�^�~�y�E�_���ɤ��ߡ@�w���W�h' onclick=""window.showModalDialog('showrule.asp','','dialogWidth=650px;dialogHeight=650px;status=no');"" ><img src='images/icon_question.gif' border='0'></a>"	
	else
		if timeflag="A" then
			sql = "select * from boo_book_T_M where bdate='"&bdate&"' and btime='"&btime&"' and teachername='"&teachername&"' and category='"&category&"' and yn='Y'"
		else
			sql = "select * from boo_book_T_M where bdate='"&bdate&"' and btime='"&btime&"' and ptime='"&ptime&"' and teachername='"&teachername&"' and category='"&category&"' and yn='Y'"
		end if
		'response.write sql
		rs.Open sql,msconn,adOpenStatic,adLockOptimistic
		if rs.EOF then
			rs.AddNew
			id= getguid()
			if id<>"" then
				rs("id")=id
			end if
			if bdate<>"" then
				rs("bdate")=bdate
			end if
			if btime<>"" then
				rs("btime")=btime
			end if
			if ptime<>"" then
				rs("ptime")=ptime 
			end if
			if timeflag<>"" then
				rs("timeflag")=timeflag
			
			end if
			if timeflag<>"" then
				rs("ptime")=ptime
			
			end if
			if teachername<>"" then
				rs("teachername")=teachername
			end if
			if item<>"" then
				rs("item")=item
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
			if score<>"" then
				rs("score")=score
			end if
			if orallevel<>"" then
				rs("orallevel")=orallevel
			end if
			if oralset<>"" then
				rs("oralset")=oralset
			end if
			if topic<>"" then
				rs("topic")=topic
			end if
			if briefing<>"" then
				rs("briefing")=briefing
			end if
			if category<>"" then
				rs("category")=category
			end if
			if languagecode<>"" then
				rs("languagecode")=languagecode
			end if
			if yms<>"" then
				rs("yms")=yms
			end if
			
			
			rs("yn") ="Y"
			rs("initdate") = date()
			rs("inituid") = session("sid")


			rs.Update
			if Err.Number=0 then 
				
			'	response.redirect "bookingtedit.asp?id=" & id & "&category=" & category
			showmessage="�s�W���\�C"
			bdate=""
			btime=""
			item=""
			teachername=""
			timeflag=""
			sid=""
			name=""
			slevel=""
			grade=""
			class1=""
			department=""
			score=""
			orallevel=""
			oralset=""
			topic=""
			briefing=""
			yms=""

			else
				showmessage= Err.Description
			end if

		else
			showmessage="���ɬq�w���H�w���C"
		end if

		rs.close
	end if
	rs2.close
end if

if session("classify")="S" then
	sid = session("sid")
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
	if (form1.item0.checked==true)
        errmsg += "�Ы��w�w������\n";
    if (form1.timeflag.value=="")
        errmsg += "�Ы��w�w���ɬq\n";
	if (form1.bdate.value=="")
        errmsg += "������ର�ť�\n";
    if (form1.btime.value=="")
        errmsg += "�ɬq���ର�ť�\n";
	if (form1.timeflag0.checked==true)
		errmsg += "�Ы��w�`��\n";
	
	if (errmsg=="")
	{
		if (form1.item3.checked==true)
		{
			if (form1.orallevel.value=="")
				errmsg += "�п�ܤf�y�ż�\n";
			if (form1.oralset.value=="")
				errmsg += "�п�ܤf�y�t�C\n";
			if (form1.topic.value=="")
				errmsg += "�п�ܤf�y�D��\n";
			
		
		}else if (form1.item4.checked==true)
		{
			if (form1.briefing.value=="")
				errmsg += "²���D�ؤ��ର�ť�\n";
		}
	
	}
	
	
    if (errmsg=="")
        return true;
    else
    {
        alert(errmsg);
        return false;
    }
}

function ChkStudent()
{
	vWinCal2 = window.open("lib/checkstudent_h.asp?sid="+form1.sid.value+"&languagecode="+ form1.languagecode.value ,"iframe_query", "width=200,height=50,status=no,resizable=no,scrollbars=no");
	vWinCal2.opener = form1;
}
function changesubject()
{
	vWinCal2 = window.open("lib/changesubject.asp?oralset="+form1.oralset.value  ,"iframe_query", "width=200,height=50,status=no,resizable=no,scrollbars=no");
	vWinCal2.opener = form1;
}
function changeschedule()
{
	vWinCal2 = window.open("lib/changeschedule.asp?yms="+form1.yms.value  ,"iframe_query", "width=200,height=50,status=no,resizable=no,scrollbars=no");
	vWinCal2.opener = form1;
}
function clickblock(id)
{
	obj=document.getElementById('area_timeflag');
	obj2=document.getElementById('area_oral');
	obj3=document.getElementById('area_briefing');
	
	if (id=="0")
	{
		form1.timeflag0.checked=true;
		obj.style.display="none";
		obj2.style.display="none";
		obj3.style.display="none";
	}else if (id=="1")
	{
		form1.timeflag0.checked=true;
		obj.style.display="block";
		obj2.style.display="none";
		obj3.style.display="none";
	}else if (id==2)
	{
		form1.timeflag0.checked=true;
		obj.style.display="block";
		obj2.style.display="none";
		obj3.style.display="none";
	}else if (id=="3")
	{
		form1.timeflag3.checked=true;
		obj.style.display="none";
		obj2.style.display="block";
		obj3.style.display="none";
		
		
	}else if (id=="4")
	{
		form1.timeflag3.checked=true;
		obj.style.display="none";
		obj2.style.display="none";
		obj3.style.display="block";
	}else if (id=="5")
	{
		form1.timeflag3.checked=true;
		obj.style.display="none";
		obj2.style.display="none";
		obj3.style.display="none";
	}else if (id=="6")
	{
		form1.timeflag3.checked=true;
		obj.style.display="none";
		obj2.style.display="none";
		obj3.style.display="none";
	}
	
	
}

function ChkScore()
{
	vWinCal2 = window.open("lib/checkscore.asp?sid="+form1.sid.value  ,"iframe_query", "width=200,height=50,status=no,resizable=no,scrollbars=no");
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center"><%if category="T" then response.write "�[�J�ǭ����v�w�������]�Ѯv���{�^" else response.write "�w���p�Ѯv�������{" end if%></TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>
	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="historybook.asp" onsubmit="return check_input();">
			<input type="hidden" value="add" name="validate">
			<input type="hidden" id="nextrec" name="nextrec">
			<input type="hidden" value="<%=category%>"  name="category" >
			<input type="hidden" value="<%=languagecode%>"  name="languagecode" >
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>�Ǹ��G</TD>
						<TD>�m�W�G</TD>
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
						<TD>�Ǩ�G</TD>
						<TD>�t�ҡG</TD>
						<TD>�~�šG</TD>
						<TD>�Z�šG</TD>
						<TD>�j�M�^�˦��Z�G</TD>
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
						<input type="text" value="<%=score%>" maxlength="25"  name="score" class="inputtext"  style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD><%if category = "T" then response.write "�Ѯv" else response.write "�p�Ѯv" end if%>�G</TD><TD>&nbsp;�y���O�G&nbsp;</TD>
					</TR>
					<TR>
						<TD>
						<select name="yms" class="inputtext" onchange="changeschedule();">
						<option value=""> - �Ǧ~�Ǵ� -</option>
						<%=YmsOption(94,Year(dateadd("m",-6,date()))-1911,yms)%>
						</select>
						<select   name="teachername" id="teachername" class="inputtext">
						<option value=""> - �L - </option>
						</select>
						</TD>
						<TD align="center"><%=languagecode%></TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR id="area_item" style="DISPLAY:<%if category="T"  then response.write "block" else response.write "none" end if %>"><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>�w�����ءG<font color="#FF0000">���w�����ؤ����B���ұ���</font></TD>
					</TR>
					<TR>
						<TD>
							<TABLE cellSpacing=0 cellPadding=0  border=0 >
							<TR>
							<TD><input type="radio" name="item" id="item0" onclick="clickblock('0');" class="inputtext" value="" <%if item="" or isnull(item) or isempty(item) then response.write "checked" end if%> ></TD><TD >�����w</TD>
							<TD><input type="radio" name="item" id="item1" onclick="clickblock('1');" class="inputtext" value="�E�_" <%if item="�E�_" then response.write "checked" end if%> ></TD><TD >�E�_</TD>
							<TD><input type="radio" name="item" id="item2" onclick="clickblock('2');" class="inputtext" value="�԰�" <%if item="�԰�" then response.write "checked" end if%>></TD><TD >�԰�</TD>
							<TD><input type="radio" name="item" id="item3" onclick="clickblock('3');" class="inputtext" value="�f�y" <%if item="�f�y" then response.write "checked" end if%>></TD><TD >�f�y</TD>
							<TD><input type="radio" name="item" id="item4" onclick="clickblock('4');" class="inputtext" value="²��" <%if item="²��" then response.write "checked" end if%>></TD><TD >²��</TD>
							<TD><input type="radio" name="item" id="item5" onclick="clickblock('5');" class="inputtext" value="�ֺq" <%if item="�ֺq" then response.write "checked" end if%>></TD><TD >�ֺq</TD>
							<TD><input type="radio" name="item" id="item6" onclick="clickblock('6');" class="inputtext" value="�g�@" <%if item="�g�@" then response.write "checked" end if%>></TD><TD >�g�@</TD>
							</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>�w������G</TD>
						<TD>�w���ɬq�G</TD>
						<TD></TD>
					</TR>
					<TR>
						<TD valign="top">
						<input type="text" value="<%=bdate%>" maxlength="25" size="15"  name="bdate" class="inputtext" readonly>
						<img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('bdate')" class="showhand">
						</TD>
						<TD valign="top">
							<select name="btime" class="inputtext" >
							<option value=""> - �Ы��w -</option>
							<optgroup label="�W��">
							<option value="0810" <%if btime="0810" then response.write "selected" end if %>>8:10��9:00</option>
							<option value="0910" <%if btime="0910" then response.write "selected" end if %>>9:10��10:00</option>
							<option value="1010" <%if btime="1010" then response.write "selected" end if%>>10:10��11:00</option>
							<option value="1110" <%if btime="1110" then response.write "selected" end if%>>11:10��12:00</option>
							</optgroup>
							<optgroup label="�U��">
							<option value="1310" <%if btime="1310" then response.write "selected" end if%>>13:10��14:00</option>
							<option value="1410" <%if btime="1410" then response.write "selected" end if%>>14:10��15:00</option>
							<option value="1510" <%if btime="1510" then response.write "selected" end if%>>15:10��16:00</option>
							<option value="1610" <%if btime="1610" then response.write "selected" end if%>>16:10��17:00</option>
							<option value="1710" <%if btime="1710" then response.write "selected" end if %>>17:10��18:00</option>
							</optgroup>
							</select>
						</TD>
						<TD id="area_timeflag" style="DISPLAY:<%if item="�E�_" or item="�԰�"  then response.write "block" else response.write "none" end if %>">
							<TABLE cellSpacing=0 cellPadding=0  border=0 >
							<TR>
							<TD><input type="radio" name="timeflag" id="timeflag0" class="inputtext" value="" <%if timeflag="" or isnull(timeflag) or isempty(timeflag) then response.write "checked" end if%> ></TD><TD >�����w</TD>
							<TD><input type="radio" name="timeflag" id="timeflag1" class="inputtext" value="U" <%if timeflag="U" then response.write "checked" end if%> ></TD><TD>�W�@�`�]25���^</TD>
							<TD><input type="radio" name="timeflag" id="timeflag2" class="inputtext" value="B" <%if timeflag="B" then response.write "checked" end if%>></TD><TD>�U�@�`�]25���^</TD>
							<TD><input type="radio" name="timeflag" id="timeflag3" class="inputtext" value="A" <%if timeflag="A" then response.write "checked" end if%>></TD><TD>�W�U�G�`(50��)</TD>
							</TR>
							<TR><TD></TD><TD></TD><TD colspan="11"><font color='blue'>�m�Y�ܽШ䥦�P�Ǥ@�_�ѻP�A�п�50�����n</font></TD></TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR id="area_oral" style="DISPLAY:<%if item="�f�y"  then response.write "block" else response.write "none" end if %>"><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						
						<TD>�f�y�t�C�G</TD>
						<TD>�f�y�D�ءG</TD>
						<TD ></TD>
					</TR>
					<TR>
						<TD>
						<select name="oralset" class="inputtext">
						<option value=""> - �Ы��w -</option>
						</select>
						</TD>
						<TD>
						<select name="topic" class="inputtext">
						<option value=""> - �Ы��w -</option>
						
						</select>
						</TD>
						<TD>
						<select name="orallevel" class="inputtext">
						<option value=""> - �Ы��w�f�y�ż� -</option>
						<option value="Level 1" <%if orallevel="Level 1" then response.write "selected" end if%>>Level 1</option>
						<option value="Level 2" <%if orallevel="Level 2" then response.write "selected" end if%>>Level 2</option>
						<option value="Level 3" <%if orallevel="Level 3" then response.write "selected" end if%>>Level 3</option>
						<option value="Level 4" <%if orallevel="Level 4" then response.write "selected" end if%>>Level 4</option>
						</select>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR id="area_briefing" style="DISPLAY:<%if item="²��"  then response.write "block" else response.write "none" end if %>"><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>²���D�ءG</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=briefing%>" maxlength="100" size="55" name="briefing" class="inputtext" >
						</TD>
						
					</TR>
				</TABLE>
			</TD></TR>
			<TR>
			<TD>
			<BR>
			<input  type="submit" value="�w��" <%=btnstatus%> class="inputbutton" >
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
<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 50 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<%
sid = trim(request("sid"))
validate = trim(request("validate"))
q1 = trim(request("q1"))
q2 = trim(request("q2"))
q3 = trim(request("q3"))
q4 = trim(request("q4"))
q5 = trim(request("q5"))
q6 = trim(request("q6"))
q7 = trim(request("q7"))
q8 = trim(request("q8"))
q9 = trim(request("q9"))
q10 = trim(request("q10"))
q11 = trim(request("q11"))
q12 = trim(request("q12"))
q13 = trim(request("q13"))
q14 = trim(request("q14"))
q15 = trim(request("q15"))
q16 = trim(request("q16"))
q17 = trim(request("q17"))
q18 = trim(request("q18"))
q19 = trim(request("q19"))
q20 = trim(request("q20"))
q21 = trim(request("q21"))
q22 = trim(request("q22"))
q23 = trim(request("q23"))
q24 = trim(request("q24"))
q25 = trim(request("q25"))
q26 = trim(request("q26"))
q27 = trim(request("q27"))
q28 = trim(request("q28"))
q29 = trim(request("q29"))
q30 = trim(request("q30"))


'response.write "sid = " & sid

set rs = server.CreateObject("adodb.recordset")

if validate="add" then
	sql = "select * from boo_questionnaire_style where sid='"&sid&"'  and  initdate='"&date()&"' "

	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
    if rs.EOF then
        rs.AddNew
		if sid<>"" then
            rs("sid")=sid
        end if
		if q1<>"" then
            rs("q1")=q1
        end if
		if q2<>"" then
            rs("q2")=q2
        end if
		if q3<>"" then
            rs("q3")=q3
        end if
		if q4<>"" then
            rs("q4")=q4
        end if
		if q5<>"" then
            rs("q5")=q5
        end if
		if q6<>"" then
            rs("q6")=q6
        end if
		if q7<>"" then
            rs("q7")=q7
        end if
		if q8<>"" then
            rs("q8")=q8
        end if
		if q9<>"" then
            rs("q9")=q9
        end if
		if q10<>"" then
            rs("q10")=q10
        end if
		if q11<>"" then
            rs("q11")=q11
        end if
		if q12<>"" then
            rs("q12")=q12
        end if
		if q13<>"" then
            rs("q13")=q13
        end if
		if q14<>"" then
            rs("q14")=q14
        end if
		if q15<>"" then
            rs("q15")=q15
        end if
		if q16<>"" then
            rs("q16")=q16
        end if
		if q17<>"" then
            rs("q17")=q17
        end if
		if q18<>"" then
            rs("q18")=q18
        end if
		if q19<>"" then
            rs("q19")=q19
        end if
		if q20<>"" then
            rs("q20")=q20
        end if
		if q21<>"" then
            rs("q21")=q21
        end if
		if q22<>"" then
            rs("q22")=q22
        end if
		if q23<>"" then
            rs("q23")=q23
        end if
		if q24<>"" then
            rs("q24")=q24
        end if
		if q25<>"" then
            rs("q25")=q25
        end if
		if q26<>"" then
            rs("q26")=q26
        end if
		if q27<>"" then
            rs("q27")=q27
        end if
		if q28<>"" then
            rs("q28")=q28
        end if
		if q29<>"" then
            rs("q29")=q29
        end if
		if q30<>"" then
            rs("q30")=q30
        end if
		rs("initdate") = date()
		rs("inituid") = session("sid")
		rs("yn")="Y"

		rs.Update
        if Err.Number=0 then 
		   sql = "update boo_profile set sytle_yn='Y' where sid='"&sid&"'"
		   msconn.Execute sql
          
        else
            showmessage= Err.Description
        end if
	
	end if
elseif validate="query" then
	sql = "select * from boo_questionnaire_style where sid = '"&sid&"' and yn='Y' order by initdate desc"
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
	if not rs.EOF then
		sid=trim(rs("sid"))
		q1 = trim(rs("q1"))
		q2 = trim(rs("q2"))
		q3 = trim(rs("q3"))
		q4 = trim(rs("q4"))
		q5 = trim(rs("q5"))
		q6 = trim(rs("q6"))
		q7 = trim(rs("q7"))
		q8 = trim(rs("q8"))
		q9 = trim(rs("q9"))
		q10 = trim(rs("q10"))
		q11 = trim(rs("q11"))
		q12 = trim(rs("q12"))
		q13 = trim(rs("q13"))
		q14 = trim(rs("q14"))
		q15 = trim(rs("q15"))
		q16 = trim(rs("q16"))
		q17 = trim(rs("q17"))
		q18 = trim(rs("q18"))
		q19 = trim(rs("q19"))
		q20 = trim(rs("q20"))
		q21 = trim(rs("q21"))
		q22 = trim(rs("q22"))
		q23 = trim(rs("q23"))
		q24 = trim(rs("q24"))
		q25 = trim(rs("q25"))
		q26 = trim(rs("q26"))
		q27 = trim(rs("q27"))
		q28 = trim(rs("q28"))
		q29 = trim(rs("q29"))
		q30 = trim(rs("q30"))
		initdate = trim(rs("initdate"))
	end if
	
end if


%>
<html>
<head>
<title>�iLDCC�^�~�y��O�E�_���ɤ��ߡj</title>


<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT LANGUAGE=javascript FOR=window EVENT=onload>
<!--
<%
if validate="add" then
%>
	//window.self.opener.location.reload();
	window.self.opener.form1.submit();
	window.close();
<%
end if
%>		
//-->
</SCRIPT>
<script language="javascript">

function btn_status()
{
	var obj;
	obj= document.getElementById("btn_close");
	obj.disabled=false;
}

function check_input()
{
    var errmsg=""
	
	
	if (form1.q1_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤ@)\n";
	if (form1.q2_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤG)\n";
	if (form1.q3_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤT)\n";
	if (form1.q4_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���إ|)\n";
	if (form1.q5_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤ�)\n";
	if (form1.q6_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤ�)\n";
	if (form1.q7_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤC)\n";
	if (form1.q8_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤK)\n";
	if (form1.q9_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤE)\n";
	if (form1.q10_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤQ)\n";
	if (form1.q11_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤQ�@)\n";
	if (form1.q12_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤQ�G)\n";
	if (form1.q13_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤQ�T)\n";
	if (form1.q14_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤQ�|)\n";
	if (form1.q15_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤQ��)\n";
	if (form1.q16_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤQ��)\n";
	if (form1.q17_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤQ�C)\n";
	if (form1.q18_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤQ�K)\n";
	if (form1.q19_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤQ�E)\n";
	if (form1.q20_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤG�Q)\n";
	if (form1.q21_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤG�Q�@)\n";
	
	if (form1.q22_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤG�Q�G)\n";
	if (form1.q23_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤG�Q�T)\n";
	if (form1.q24_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤG�Q�|)\n";
	if (form1.q25_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤG�Q��)\n";
	if (form1.q26_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤG�Q��)\n";
	if (form1.q27_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤG�Q�C)\n";
	if (form1.q28_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤG�Q�K)\n";
	if (form1.q29_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤG�Q�E)\n";
	if (form1.q30_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤT�Q)\n";

    if (errmsg=="")
        return true;
    else
    {
        alert(errmsg);
        return false;
    }
}

function wclose()
{
	if (confirm('�A�|���x�s���סA�T�w�n���}�ܡH'))
	{window.close();}
}
</script>
<script language="JavaScript">

function document.onkeydown() 
	{
	if ( event.keyCode==17) 
		{ event.keyCode = 0; 
		event.cancelBubble = true; 
		return false; 
		}
	}

function right(e) {
if (navigator.appName =='Netscape'&&
(e.which ==3|| e.which ==2))
return false;
else if (navigator.appName == 'Microsoft Internet Explorer' &&
(event.button == 2|| event.button ==3)) {
alert("�дL�����z�]���v�A���¡C\n");
return false;
}
return true;
}
document.onmousedown=right;
if (document.layers) window.captureEvents(Event.MOUSEDOWN);
window.onmousedown=right;
	
</script>
</head>
<body bottomMargin=0 leftMargin=0 topMargin=0 rightMargin=0 marginheight="0" marginwidth="0" ONDRAGSTART="window.event.returnValue=false" ONCONTEXTMENU="window.event.returnValue=false" onSelectStart="event.returnValue=false">

<table width="100%" height="100%" border=1 cellpadding=0 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
<tr><td bgcolor=555555 height=35 align="center" class="T2">
<font  color="#FFFFFF">��Perceptual Learning-Style Preference Questionnaire ��</font></td></tr>
<tr valign="top"><td bgcolor=#ECECE3>
<table width="780" align="center"  >
<tr><td  height=35 align="center" class="T2">
<%if validate="query" then%>
<center><input type="button" value="���}"  onclick="window.close();" class="inputbutton" >&nbsp;&nbsp;
��g�ɶ��G<%=initdate%>
<%else%>
�`�N�ƶ��G1. �ݨ��������^�媩�A�ХH�\Ū�^���D�ج��D�A�����ɦA�ѦҤ���<BR>
											 2. �Ҧ��D�جҥH�ǲ߭^�嬰�D
<%end if%>

</td></tr>
<tr><td  valign="top" align="center">
<BR><BR>
	<form id="form1" name="form1" method="post" action="qstyle.asp" onsubmit="return check_input();">
	<input type="hidden" value="add" name="validate">
	<input type="hidden" value="<%=sid%>" name="sid">
	<table width="100%" cellpadding=2 cellspacing=0 border="1" bordercolorlight=#333333 bordercolordark=#dddddd>
	<tr><td width="5%">Item</td><td  width="62%">&nbsp;</td><td  width="8%">�����w</td><td  width="5%">SD</td><td  width="5%">D</td><td  width="5%">U</td><td  width="5%"> A</td><td  width="5%">SA</td></tr>
	<tr><td>1</td>
	<td>When the teacher tells me the instructions I understand better.<BR>��Ѯv�έ^��f�Y�����ɡA�گ����A��</td>
	<td><input type="radio" checked value="" id="q1_0" name="q1"></td>
	<td><input type="radio" value="1" name="q1" <%if q1="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q1" <%if q1="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q1" <%if q1="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q1" <%if q1="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q1" <%if q1="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>2</td>
	<td>I prefer to learn by doing something in class.<BR>�ڤ�����w�b�Ұ�W���Ǭ��ʨӾǲ߭^��</td>
	<td><input type="radio" checked value="" id="q2_0" name="q2"></td>
	<td><input type="radio" value="1" name="q2" <%if q2="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q2" <%if q2="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q2" <%if q2="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q2" <%if q2="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q2" <%if q2="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>3</td>
	<td>I get more work done when I work with others.<BR>�P�P�Ǥ@�_���@�~�Υ\�ҡA�ڪ��Ĳv����n</td>
	<td><input type="radio" checked value="" id="q3_0" name="q3"></td>
	<td><input type="radio" value="1" name="q3" <%if q3="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q3" <%if q3="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q3" <%if q3="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q3" <%if q3="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q3" <%if q3="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>4</td>
	<td>I learn more when I study with a group.<BR>�M�p�դ@�_�ǲߡA�ڷ|�Ǳo����h</td>
	<td><input type="radio" checked value="" id="q4_0" name="q4"></td>
	<td><input type="radio" value="1" name="q4" <%if q4="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q4" <%if q4="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q4" <%if q4="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q4" <%if q4="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q4" <%if q4="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>5</td>
	<td>In class, I learn best when I work with others.<BR>��ڻP�P�Ǥ@�_�����ʮɡA�ڪ��ǲ߮ĪG�̦n</td>
	<td><input type="radio" checked value="" id="q5_0" name="q5"></td>
	<td><input type="radio" value="1" name="q5" <%if q5="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q5" <%if q5="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q5" <%if q5="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q5" <%if q5="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q5" <%if q5="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>6</td>
	<td>I learn better by reading what the teacher writes on the chalkboard.<BR>�p�G�ڥi�H�ݨ�Ѯv�b�ªO�W�g���W�Ҥ��e�A�ڷ|�Ǳo���n</td>
	<td><input type="radio" checked value="" id="q6_0" name="q6"></td>
	<td><input type="radio" value="1" name="q6" <%if q6="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q6" <%if q6="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q6" <%if q6="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q6" <%if q6="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q6" <%if q6="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>7</td>
	<td>When someone tells me how to do something in class, I learn it better.<BR>�Ұ�W�A�Y���H�i�D�ڸӦp�󰵡A�ڷ|�ǲ߱o����n</td>
	<td><input type="radio" checked value="" id="q7_0" name="q7"></td>
	<td><input type="radio" value="1" name="q7" <%if q7="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q7" <%if q7="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q7" <%if q7="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q7" <%if q7="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q7" <%if q7="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>8</td>
	<td>When I do things in class, I learn better.<BR>�b�Ұ�W�����ʡA�ڷ|�ǲߪ����n</td>
	<td><input type="radio" checked value="" id="q8_0" name="q8"></td>
	<td><input type="radio" value="1" name="q8" <%if q8="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q8" <%if q8="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q8" <%if q8="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q8" <%if q8="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q8" <%if q8="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>9</td>
	<td>I remember things I have heard in class better than things I have read.<BR>�b�Ұ󤤡A�ڸ��e���O���ť�쪺���e�ө��Ū�����e</td>
	<td><input type="radio" checked value="" id="q9_0" name="q9"></td>
	<td><input type="radio" value="1" name="q9" <%if q9="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q9" <%if q9="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q9" <%if q9="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q9" <%if q9="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q9" <%if q9="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>10</td>
	<td>When I read instructions, I remember them better.<BR>��ڦۤv�\Ū�����ɡA�ڤ���O�o��</td>
	<td><input type="radio" checked value="" id="q10_0" name="q10"></td>
	<td><input type="radio" value="1" name="q10" <%if q10="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q10" <%if q10="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q10" <%if q10="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q10" <%if q10="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q10" <%if q10="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>11</td>
	<td>I learn more when I can make a model of something.<BR>�Y�گ��ڰʤⰵ�A�ڷ|�Ǳo��h</td>
	<td><input type="radio" checked value="" id="q11_0" name="q11"></td>
	<td><input type="radio" value="1" name="q11" <%if q11="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q11" <%if q11="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q11" <%if q11="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q11" <%if q11="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q11" <%if q11="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>12</td>
	<td>I understand better when I read instructions.<BR>�ڦۤv�\Ū���������A��</td>
	<td><input type="radio" checked value="" id="q12_0" name="q12"></td>
	<td><input type="radio" value="1" name="q12" <%if q12="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q12" <%if q12="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q12" <%if q12="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q12" <%if q12="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q12" <%if q12="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>13</td>
	<td>When I study alone, I remember things better.<BR>�ڦۤv�@�H�ǲ߮ɸ��e���O��ҾǪ�</td>
	<td><input type="radio" checked value="" id="q13_0" name="q13"></td>
	<td><input type="radio" value="1" name="q13"  <%if q13="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q13"  <%if q13="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q13"  <%if q13="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q13"  <%if q13="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q13"  <%if q13="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>14</td>
	<td>I learn more when I make something for a class project.<BR>�ڦp�G�˦۰ѻP�B�z���w���ҷ~�A�ڷ|�Ǳo���h</td>
	<td><input type="radio" checked value="" id="q14_0" name="q14"></td>
	<td><input type="radio" value="1" name="q14" <%if q14="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q14" <%if q14="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q14" <%if q14="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q14" <%if q14="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q14" <%if q14="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>15</td>
	<td>I enjoy learning in class by doing experiments.<BR>�ڳ��w�b�Ұ󤤡A�ǥѹ�ڰѻP���ʨӾǲ�</td>
	<td><input type="radio" checked value="" id="q15_0" name="q15"></td>
	<td><input type="radio" value="1" name="q15" <%if q15="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q15" <%if q15="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q15" <%if q15="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q15" <%if q15="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q15" <%if q15="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>16</td>
	<td>I learn better when I make drawings as I study.<BR>�����e�ϡA�ڷ|�Ǳo��n</td>
	<td><input type="radio" checked value="" id="q16_0" name="q16"></td>
	<td><input type="radio" value="1" name="q16" <%if q16="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q16" <%if q16="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q16" <%if q16="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q16" <%if q16="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q16" <%if q16="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>17</td>
	<td>I learn better in class when the teacher gives a lecture.<BR>��Ѯv���H���Ҥ覡�i��A�ھǱo���h</td>
	<td><input type="radio" checked value="" id="q17_0" name="q17"></td>
	<td><input type="radio" value="1" name="q17" <%if q17="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q17" <%if q17="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q17" <%if q17="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q17" <%if q17="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q17" <%if q17="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>18</td>
	<td>When I work alone, I learn better.<BR>��ڿW�۾ǲߧھǱo���n</td>
	<td><input type="radio" checked value="" id="q18_0" name="q18"></td>
	<td><input type="radio" value="1" name="q18"  <%if q18="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q18" <%if q18="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q18" <%if q18="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q18" <%if q18="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q18" <%if q18="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>19</td>
	<td>I understand things better in class when I participate in role-playing.<BR>�Ұ󤤡A�p�ڰѻP�����t�����ʡA�ڷ|�A�ѱo��h</td>
	<td><input type="radio" checked value="" id="q19_0" name="q19"></td>
	<td><input type="radio" value="1" name="q19" <%if q19="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q19" <%if q19="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q19" <%if q19="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q19" <%if q19="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q19" <%if q19="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>20</td>
	<td>I learn better in class when I listen to someone.<BR>�Ұ󤤡A���ť�Ѯv�ΦP�Ǹѻ��A�ڷ|�A�ѧ�h</td>
	<td><input type="radio" checked value="" id="q20_0" name="q20"></td>
	<td><input type="radio" value="1" name="q20" <%if q20="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q20" <%if q20="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q20" <%if q20="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q20" <%if q20="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q20" <%if q20="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>21</td>
	<td>I enjoy working on an assignment with two or three classmates.<BR>�ڳ��w��B�T��P�Ǥ@�_���@�~</td>
	<td><input type="radio" checked value="" id="q21_0" name="q21"></td>
	<td><input type="radio" value="1" name="q21"  <%if q21="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q21"  <%if q21="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q21"  <%if q21="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q21"  <%if q21="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q21"  <%if q21="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>22</td>
	<td>When I build something, I remember what I have learned better.<BR>��ڦۤv�@�@�ǪF��A�ڸ����O��q���ҾǨ쪺</td>
	<td><input type="radio" checked value="" id="q22_0" name="q22"></td>
	<td><input type="radio" value="1" name="q22" <%if q22="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q22" <%if q22="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q22" <%if q22="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q22" <%if q22="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q22" <%if q22="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>23</td>
	<td>I prefer to study with others.<BR>�ڸ����w�M�O�H�@�_�ǲ�</td>
	<td><input type="radio" checked value="" id="q23_0" name="q23"></td>
	<td><input type="radio" value="1" name="q23" <%if q23="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q23" <%if q23="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q23" <%if q23="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q23" <%if q23="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q23" <%if q23="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>24</td>
	<td>I learn better by reading than by listening to someone.<BR>��_ť�O�H���ѡA�ڦۤv�\Ū��Ǳo��n</td>
	<td><input type="radio" checked value="" id="q24_0" name="q24"></td>
	<td><input type="radio" value="1" name="q24" <%if q24="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q24" <%if q24="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q24" <%if q24="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q24" <%if q24="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q24" <%if q24="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>25</td>
	<td>I enjoy making something for a class project.<BR>�ڳ��w�t�d�ʤⰵ���w���ҷ~</td>
	<td><input type="radio" checked value="" id="q25_0" name="q25"></td>
	<td><input type="radio" value="1" name="q25" <%if q25="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q25" <%if q25="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q25" <%if q25="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q25" <%if q25="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q25" <%if q25="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>26</td>
	<td>I learn best in class when I can participate in related activities.<BR>�ڪ��ǲ߮ĪG���n�p��ѻP�������ʮ�</td>
	<td><input type="radio" checked value="" id="q26_0" name="q26"></td>
	<td><input type="radio" value="1" name="q26" <%if q26="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q26" <%if q26="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q26" <%if q26="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q26" <%if q26="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q26" <%if q26="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>27</td>
	<td>In class, I work better when I work alone.<BR>�ڦۤvŪ�Ѿǲߤ�M�P�Ǥ@�_Ū�A�ĪG���n</td>
	<td><input type="radio" checked value="" id="q27_0" name="q27"></td>
	<td><input type="radio" value="1" name="q27" <%if q27="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q27" <%if q27="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q27" <%if q27="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q27" <%if q27="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q27" <%if q27="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>28</td>
	<td>I prefer working on projects by myself.<BR>�ڤ�����w�ۤv�W�ߧ@�~</td>
	<td><input type="radio" checked value="" id="q28_0" name="q28"></td>
	<td><input type="radio" value="1" name="q28" <%if q28="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q28" <%if q28="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q28" <%if q28="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q28" <%if q28="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q28" <%if q28="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>29</td>
	<td>I learn more by reading textbooks than by listening to lectures.<BR>��_��ť�½ҡA�ڦۤv�\Ū�Ь�ѷ|�Ǳo����h</td>
	<td><input type="radio" checked value="" id="q29_0" name="q29"></td>
	<td><input type="radio" value="1" name="q29" <%if q29="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q29" <%if q29="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q29" <%if q29="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q29" <%if q29="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q29" <%if q29="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>30</td>
	<td>I prefer to work by myself.<BR>�ڤ�����w�ۤvŪ��</td>
	<td><input type="radio" checked value="" id="q30_0" name="q30"></td>
	<td><input type="radio" value="1" name="q30" <%if q30="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q30" <%if q30="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q30" <%if q30="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q30" <%if q30="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q30" <%if q30="5" then response.write "checked" end if%>></td>
	</tr>

	</table>

</td></tr>
</table>
<BR><BR>
<%if validate<>"query" then%>
<center><input type="button" value="���}"  id="btn_close" onclick="wclose();" class="inputbutton" >&nbsp;&nbsp;
<input type="submit" value="�e�X����"  class="inputbutton" >
<%else%>
<center><input type="button" value="���}"  onclick="window.close();" class="inputbutton" >&nbsp;&nbsp;

<%end if%>
<BR><BR>
</form>
</td></tr>
<tr><td bgcolor=#555555 height=24 align=right><font Color="#FFFFFF">�w���������D�Ь�LDCC--�\���@ ����7403 </font></td></tr></table>

</body>
</html>
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->
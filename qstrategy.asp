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
q31 = trim(request("q31"))
q32 = trim(request("q32"))
q33 = trim(request("q33"))
q34 = trim(request("q34"))
q35 = trim(request("q35"))
q36 = trim(request("q36"))
q37 = trim(request("q37"))
q38 = trim(request("q38"))
q39 = trim(request("q39"))
q40 = trim(request("q40"))
q41 = trim(request("q41"))
q42 = trim(request("q42"))
q43 = trim(request("q43"))
q44 = trim(request("q44"))
q45 = trim(request("q45"))
q46 = trim(request("q46"))
q47 = trim(request("q47"))
q48 = trim(request("q48"))
q49 = trim(request("q49"))
q50 = trim(request("q50"))


'response.write "sid = " & sid

set rs = server.CreateObject("adodb.recordset")

if validate="add" then
	sql = "select * from boo_questionnaire_strategy where sid='"&sid&"'  and  initdate='"&date()&"' "

	'response.write sql
	'response.end
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
		if q31<>"" then
            rs("q31")=q31
        end if
		if q32<>"" then
            rs("q32")=q32
        end if
		if q33<>"" then
            rs("q33")=q33
        end if
		if q34<>"" then
            rs("q34")=q34
        end if
		if q35<>"" then
            rs("q35")=q35
        end if
		if q36<>"" then
            rs("q36")=q36
        end if
		if q37<>"" then
            rs("q37")=q37
        end if
		if q38<>"" then
            rs("q38")=q38
        end if
		if q39<>"" then
            rs("q39")=q39
        end if
		if q40<>"" then
            rs("q40")=q40
        end if
		if q41<>"" then
            rs("q41")=q41
        end if
		if q42<>"" then
            rs("q42")=q42
        end if
		if q43<>"" then
            rs("q43")=q43
        end if
		if q44<>"" then
            rs("q44")=q44
        end if
		if q45<>"" then
            rs("q45")=q45
        end if
		if q46<>"" then
            rs("q46")=q46
        end if
		if q47<>"" then
            rs("q47")=q47
        end if
		if q48<>"" then
            rs("q48")=q48
        end if
		if q49<>"" then
            rs("q49")=q49
        end if
		if q50<>"" then
            rs("q50")=q50
        end if
		rs("initdate") = date()
		rs("inituid") = session("sid")
		rs("yn")="Y"

		rs.Update
        if Err.Number=0 then 
		   sql = "update boo_profile set strategy_yn='Y' where sid='"&sid&"'"
		   msconn.Execute sql
          
        else
            showmessage= Err.Description
        end if
	
	end if
elseif validate="query" then
	sql = "select * from boo_questionnaire_strategy where sid = '"&sid&"' and yn='Y' order by initdate desc"
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
		q31 = trim(rs("q31"))
		q32 = trim(rs("q32"))
		q33 = trim(rs("q33"))
		q34 = trim(rs("q34"))
		q35 = trim(rs("q35"))
		q36 = trim(rs("q36"))
		q37 = trim(rs("q37"))
		q38 = trim(rs("q38"))
		q39 = trim(rs("q39"))
		q40 = trim(rs("q40"))
		q41 = trim(rs("q41"))
		q42 = trim(rs("q42"))
		q43 = trim(rs("q43"))
		q44 = trim(rs("q44"))
		q45 = trim(rs("q45"))
		q46 = trim(rs("q46"))
		q47 = trim(rs("q47"))
		q48 = trim(rs("q48"))
		q49 = trim(rs("q49"))
		q50 = trim(rs("q50"))
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

	if (form1.q31_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤT�Q�@)\n";
	if (form1.q32_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤT�Q�G)\n";
	if (form1.q33_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤT�Q�T)\n";
	if (form1.q34_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤT�Q�|)\n";
	if (form1.q35_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤT�Q��)\n";
	if (form1.q36_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤT�Q��)\n";
	if (form1.q37_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤT�Q�C)\n";
	if (form1.q38_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤT�Q�K)\n";
	if (form1.q39_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤT�Q�E)\n";
	if (form1.q40_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���إ|�Q)\n";
	
	if (form1.q41_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���إ|�Q�@)\n";
	if (form1.q42_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���إ|�Q�G)\n";
	if (form1.q43_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���إ|�Q�T)\n";
	if (form1.q44_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���إ|�Q�|)\n";
	if (form1.q45_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���إ|�Q��)\n";
	if (form1.q46_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���إ|�Q��)\n";
	if (form1.q47_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���إ|�Q�C)\n";
	if (form1.q48_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���إ|�Q�K)\n";
	if (form1.q49_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���إ|�Q�E)\n";
	if (form1.q50_0.checked==true)
        errmsg += "You have item not to answer. (�A�����ץ���ܡA���ؤ��Q)\n";
	


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
<font  color="#FFFFFF">��STRATEGY INVENTORY FOR LANGUAGE LEARNING (SILL) ��</font></td></tr>
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
	<form id="form1" name="form1" method="post" action="qstrategy.asp" onsubmit="return check_input();">
	<input type="hidden" value="add" name="validate">
	<input type="hidden" value="<%=sid%>" name="sid">
	<font class="T2" ><B>Part A�G</B></font><br>
	<table width="100%" cellpadding=2 cellspacing=0 border="1" bordercolorlight=#333333 bordercolordark=#dddddd>
	<tr><td width="5%">Item</td><td  width="62%">&nbsp;</td><td  width="8%">�����w</td><td  width="5%">1</td><td  width="5%">2</td><td  width="5%">3</td><td  width="5%"> 4</td><td  width="5%">5</td></tr>
	<tr><td>1</td>
	<td>I think of relationships between what I already know and new things I learn in English.<BR>�ڷ|�h��ҾǹL���M�s�Ǫ��^�y���������Y�C</td>
	<td><input type="radio" checked value="" id="q1_0" name="q1"></td>
	<td><input type="radio" value="1" name="q1" <%if q1="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q1" <%if q1="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q1" <%if q1="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q1" <%if q1="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q1" <%if q1="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>2</td>
	<td>I use new English words in a sentence so I can remember them. <BR>���F�O��s�Ǫ��^�y��r�A�ڷ|�յۥγo�ǥͦr�ӳy�y�C</td>
	<td><input type="radio" checked value="" id="q2_0" name="q2"></td>
	<td><input type="radio" value="1" name="q2" <%if q2="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q2" <%if q2="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q2" <%if q2="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q2" <%if q2="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q2" <%if q2="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>3</td>
	<td>I connect the sound of a new English word and an image or picture of the word to help me remember the word. <BR>�ڷ|�b�������Q�X�i�H�t�X�^�y�n�����Ϥ��ηN�H�A�H�K�O��Y�ӳ�r�C </td>
	<td><input type="radio" checked value="" id="q3_0" name="q3"></td>
	<td><input type="radio" value="1" name="q3" <%if q3="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q3" <%if q3="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q3" <%if q3="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q3" <%if q3="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q3" <%if q3="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>4</td>
	<td>I remember a new English word by making a mental picture of a situation in which the word might be used. <BR>�ڷ|�b�����s�y�X�Y�ӥͦr�X�{�����ҡA�H�o�ؤ�k���r�I�_�ӡC</td>
	<td><input type="radio" checked value="" id="q4_0" name="q4"></td>
	<td><input type="radio" value="1" name="q4" <%if q4="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q4" <%if q4="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q4" <%if q4="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q4" <%if q4="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q4" <%if q4="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>5</td>
	<td>I use rhymes to remember new English words. <BR>�ڷ|�ϥΩ������覡�ӰO��ͦr�C</td>
	<td><input type="radio" checked value="" id="q5_0" name="q5"></td>
	<td><input type="radio" value="1" name="q5" <%if q5="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q5" <%if q5="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q5" <%if q5="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q5" <%if q5="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q5" <%if q5="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>6</td>
	<td>I use flashcards to remember new English words.<BR>�ڷ|�ϥΰ{�ܥd�ӭI�ͦr�C</td>
	<td><input type="radio" checked value="" id="q6_0" name="q6"></td>
	<td><input type="radio" value="1" name="q6" <%if q6="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q6" <%if q6="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q6" <%if q6="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q6" <%if q6="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q6" <%if q6="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>7</td>
	<td>I physically act out new English words.<BR>�ڷ|��ͦr�Ϊ���t�X�ӡC </td>
	<td><input type="radio" checked value="" id="q7_0" name="q7"></td>
	<td><input type="radio" value="1" name="q7" <%if q7="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q7" <%if q7="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q7" <%if q7="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q7" <%if q7="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q7" <%if q7="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>8</td>
	<td>I review English lessons often. <BR>�ڱ`�`�Ʋ߭^�y�ҵ{�C</td>
	<td><input type="radio" checked value="" id="q8_0" name="q8"></td>
	<td><input type="radio" value="1" name="q8" <%if q8="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q8" <%if q8="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q8" <%if q8="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q8" <%if q8="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q8" <%if q8="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>9</td>
	<td>I remember new English words or phrases by remembering their location on the page, on the board, or on a street sign.<BR>�ڷ|���ӥͦr�Τ��y�X�{�b�ҥ��B�ªO�άO��D�ݪO����m�A�ӰO��ͦr�Τ��y�C</td>
	<td><input type="radio" checked value="" id="q9_0" name="q9"></td>
	<td><input type="radio" value="1" name="q9" <%if q9="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q9" <%if q9="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q9" <%if q9="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q9" <%if q9="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q9" <%if q9="5" then response.write "checked" end if%>></td>
	</tr>
	</table>
	<br><font class="T2" ><B>Part B�G</B></font><br>
	<table width="100%" cellpadding=2 cellspacing=0 border="1" bordercolorlight=#333333 bordercolordark=#dddddd>
	<tr><td width="5%">Item</td><td  width="62%">&nbsp;</td><td  width="8%">�����w</td><td  width="5%">1</td><td  width="5%">2</td><td  width="5%">3</td><td  width="5%"> 4</td><td  width="5%">5</td></tr>
	<tr><td>10</td>
	<td>I say or write new English words several times.<BR>�ڷ|���ƻ��μg�^��ͦr�n�X���C</td>
	<td><input type="radio" checked value="" id="q10_0" name="q10"></td>
	<td><input type="radio" value="1" name="q10" <%if q10="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q10" <%if q10="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q10" <%if q10="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q10" <%if q10="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q10" <%if q10="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>11</td>
	<td>I try to talk like native English speakers. <BR>�ڷ|�Q��^�y���o���H�^�y�����y���H�@�ˡC</td>
	<td><input type="radio" checked value="" id="q11_0" name="q11"></td>
	<td><input type="radio" value="1" name="q11" <%if q11="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q11" <%if q11="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q11" <%if q11="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q11" <%if q11="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q11" <%if q11="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>12</td>
	<td>I practice the sounds of English. <BR>�ڷ|�m�߭^�y���o���C</td>
	<td><input type="radio" checked value="" id="q12_0" name="q12"></td>
	<td><input type="radio" value="1" name="q12" <%if q12="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q12" <%if q12="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q12" <%if q12="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q12" <%if q12="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q12" <%if q12="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>13</td>
	<td>I use the English words I know in different ways. <BR>�ڷ|��ǹL���^��r�Φb���P���譱�W�C</td>
	<td><input type="radio" checked value="" id="q13_0" name="q13"></td>
	<td><input type="radio" value="1" name="q13"  <%if q13="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q13"  <%if q13="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q13"  <%if q13="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q13"  <%if q13="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q13"  <%if q13="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>14</td>
	<td>I start conversations in English. <BR>�ڷ|�H�^�y�}�ҹ�ܡC</td>
	<td><input type="radio" checked value="" id="q14_0" name="q14"></td>
	<td><input type="radio" value="1" name="q14" <%if q14="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q14" <%if q14="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q14" <%if q14="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q14" <%if q14="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q14" <%if q14="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>15</td>
	<td>I watch English TV shows spoken in English or go to movies spoken in English. <BR>�ڷ|�ݥH�^�y�o�����q���`�ةιq�v�C</td>
	<td><input type="radio" checked value="" id="q15_0" name="q15"></td>
	<td><input type="radio" value="1" name="q15" <%if q15="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q15" <%if q15="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q15" <%if q15="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q15" <%if q15="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q15" <%if q15="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>16</td>
	<td>I read for pleasure in English. <BR>�ھ\Ū�^�尵���𶢬��ʡC</td>
	<td><input type="radio" checked value="" id="q16_0" name="q16"></td>
	<td><input type="radio" value="1" name="q16" <%if q16="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q16" <%if q16="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q16" <%if q16="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q16" <%if q16="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q16" <%if q16="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>17</td>
	<td>I write notes, messages, letters, or reports in English. <BR>�ڷ|�H�^�y�ӰO���O�B�T���B�ѫH�άO���i�C</td>
	<td><input type="radio" checked value="" id="q17_0" name="q17"></td>
	<td><input type="radio" value="1" name="q17" <%if q17="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q17" <%if q17="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q17" <%if q17="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q17" <%if q17="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q17" <%if q17="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>18</td>
	<td>I first skim an English passage (read over the passage quickly) then go back and read carefully. <BR>�ڷ|����Ū�^�y���峹(�ܧ֦a��峹�ݹL�@�M)�A�M��A�^�ӲӬݡC</td>
	<td><input type="radio" checked value="" id="q18_0" name="q18"></td>
	<td><input type="radio" value="1" name="q18"  <%if q18="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q18" <%if q18="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q18" <%if q18="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q18" <%if q18="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q18" <%if q18="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>19</td>
	<td>I look for words in my own language that are similar to new words in English.<BR>�ڷ|�b�ڪ����y�̧�M�M�^�y�ۦP���ͦr�C</td>
	<td><input type="radio" checked value="" id="q19_0" name="q19"></td>
	<td><input type="radio" value="1" name="q19" <%if q19="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q19" <%if q19="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q19" <%if q19="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q19" <%if q19="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q19" <%if q19="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>20</td>
	<td>I try to find patterns in English. <BR>�ڷ|��X�^�y���Ҧ��C</td>
	<td><input type="radio" checked value="" id="q20_0" name="q20"></td>
	<td><input type="radio" value="1" name="q20" <%if q20="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q20" <%if q20="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q20" <%if q20="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q20" <%if q20="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q20" <%if q20="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>21</td>
	<td>I find the meaning of an English word by dividing it into parts that I understand.<BR>�ڷ|��^�y��Ѷ}�ӡA��X�ۤv���������A�ǥH�F�ѳ�r���N��C</td>
	<td><input type="radio" checked value="" id="q21_0" name="q21"></td>
	<td><input type="radio" value="1" name="q21"  <%if q21="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q21"  <%if q21="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q21"  <%if q21="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q21"  <%if q21="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q21"  <%if q21="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>22</td>
	<td>I try not to translate word-for-word. <BR>�ڤ��|�v�r�v�y½Ķ�C</td>
	<td><input type="radio" checked value="" id="q22_0" name="q22"></td>
	<td><input type="radio" value="1" name="q22" <%if q22="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q22" <%if q22="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q22" <%if q22="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q22" <%if q22="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q22" <%if q22="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>23</td>
	<td>I make summaries of information that I hear or read in English. <BR>�ڷ|��ť��άOŪ�쪺�^���T�����K�n�C</td>
	<td><input type="radio" checked value="" id="q23_0" name="q23"></td>
	<td><input type="radio" value="1" name="q23" <%if q23="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q23" <%if q23="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q23" <%if q23="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q23" <%if q23="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q23" <%if q23="5" then response.write "checked" end if%>></td>
	</tr>
	</table>
	<br><font class="T2" ><B>Part C�G</B></font><br>
	<table width="100%" cellpadding=2 cellspacing=0 border="1" bordercolorlight=#333333 bordercolordark=#dddddd>
	<tr><td width="5%">Item</td><td  width="62%">&nbsp;</td><td  width="8%">�����w</td><td  width="5%">1</td><td  width="5%">2</td><td  width="5%">3</td><td  width="5%"> 4</td><td  width="5%">5</td></tr>
	<tr><td>24</td>
	<td>To understand unfamiliar English words, I make guesses.<BR>�J�줣���x���^���r�A�ڷ|�h�q�����N��C</td>
	<td><input type="radio" checked value="" id="q24_0" name="q24"></td>
	<td><input type="radio" value="1" name="q24" <%if q24="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q24" <%if q24="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q24" <%if q24="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q24" <%if q24="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q24" <%if q24="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>25</td>
	<td>When I can think of a word during a conversation in English, I use gestures.<BR>�b��ܤ��A�ڦp�G�Q���X�Y�Ӧr�^���򻡡A�ڷ|�ϥΪ��M�ʧ@�C</td>
	<td><input type="radio" checked value="" id="q25_0" name="q25"></td>
	<td><input type="radio" value="1" name="q25" <%if q25="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q25" <%if q25="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q25" <%if q25="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q25" <%if q25="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q25" <%if q25="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>26</td>
	<td>I make up new words if I do not know the right ones in English. <BR>�p�G�ڤ����D�^�y�ӫ�򻡡A�ڷ|�ۤv�гy�s�r�C</td>
	<td><input type="radio" checked value="" id="q26_0" name="q26"></td>
	<td><input type="radio" value="1" name="q26" <%if q26="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q26" <%if q26="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q26" <%if q26="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q26" <%if q26="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q26" <%if q26="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>27</td>
	<td>I read English without looking up every new word.<BR>�\Ū���L�{���A�ڤ@�J��ͦr�N���W�d�r��C</td>
	<td><input type="radio" checked value="" id="q27_0" name="q27"></td>
	<td><input type="radio" value="1" name="q27" <%if q27="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q27" <%if q27="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q27" <%if q27="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q27" <%if q27="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q27" <%if q27="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>28</td>
	<td>I try to guess what the other person will say next in English.<BR>�ڷ|�έ^�y�յۥh�q�O�H���۷|������C</td>
	<td><input type="radio" checked value="" id="q28_0" name="q28"></td>
	<td><input type="radio" value="1" name="q28" <%if q28="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q28" <%if q28="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q28" <%if q28="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q28" <%if q28="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q28" <%if q28="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>29</td>
	<td>I can think of an English word, I use a word or phrase that means the same thing. <BR>�p�G�ڷQ���_�ӬY�ӭ^���r�A�ڷ|�ΧO���r�Τ��y����z�P�˪��N��C</td>
	<td><input type="radio" checked value="" id="q29_0" name="q29"></td>
	<td><input type="radio" value="1" name="q29" <%if q29="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q29" <%if q29="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q29" <%if q29="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q29" <%if q29="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q29" <%if q29="5" then response.write "checked" end if%>></td>
	</tr>
	</table>
	<br><font class="T2" ><B>Part D�G</B></font><br>
	<table width="100%" cellpadding=2 cellspacing=0 border="1" bordercolorlight=#333333 bordercolordark=#dddddd>
	<tr><td width="5%">Item</td><td  width="62%">&nbsp;</td><td  width="8%">�����w</td><td  width="5%">1</td><td  width="5%">2</td><td  width="5%">3</td><td  width="5%"> 4</td><td  width="5%">5</td></tr>
	<tr><td>30</td>
	<td>I try to find as many ways as I can to use my English.<BR>�ڷ|���q����|�m�߭^�y�C</td>
	<td><input type="radio" checked value="" id="q30_0" name="q30"></td>
	<td><input type="radio" value="1" name="q30" <%if q30="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q30" <%if q30="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q30" <%if q30="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q30" <%if q30="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q30" <%if q30="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>31</td>
	<td>I notice my English mistakes and I use that information to help me do better.<BR>�ڷ|�`�N�کҥǪ����~�A�Ǧ����U�ۤv�Ǳo��n�C</td>
	<td><input type="radio" checked value="" id="q31_0" name="q31"></td>
	<td><input type="radio" value="1" name="q31" <%if q31="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q31" <%if q31="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q31" <%if q31="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q31" <%if q31="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q31" <%if q31="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>32</td>
	<td>I pay attention when someone is speaking English. <BR>���H�b���^�y�ɡA�|�ް_�ڪ��`�N�C</td>
	<td><input type="radio" checked value="" id="q32_0" name="q32"></td>
	<td><input type="radio" value="1" name="q32" <%if q32="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q32" <%if q32="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q32" <%if q32="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q32" <%if q32="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q32" <%if q32="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>33</td>
	<td>I try to find out how to be a better learner of English. <BR>�ڷ|�Q��k���ۤv������n���^�y�ǲߪ̡C</td>
	<td><input type="radio" checked value="" id="q33_0" name="q33"></td>
	<td><input type="radio" value="1" name="q33" <%if q33="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q33" <%if q33="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q33" <%if q33="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q33" <%if q33="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q33" <%if q33="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>34</td>
	<td>I plan my schedule so I will have enough time to study English.<BR>�ڷ|�n�n�W�e�ɶ��A�H�K���������ɶ��ǭ^�y�C</td>
	<td><input type="radio" checked value="" id="q34_0" name="q34"></td>
	<td><input type="radio" value="1" name="q34" <%if q34="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q34" <%if q34="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q34" <%if q34="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q34" <%if q34="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q34" <%if q34="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>35</td>
	<td>I look for people I can talk to in English.  <BR>�ڷ|���έ^�y�͸ܪ��H�m�߭^�y�C</td>
	<td><input type="radio" checked value="" id="q35_0" name="q35"></td>
	<td><input type="radio" value="1" name="q35" <%if q35="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q35" <%if q35="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q35" <%if q35="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q35" <%if q35="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q35" <%if q35="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>36</td>
	<td>I look for opportunities to read as much as possible in English. <BR>�ڷ|���q����|�\Ū�^�y�C</td>
	<td><input type="radio" checked value="" id="q36_0" name="q36"></td>
	<td><input type="radio" value="1" name="q36" <%if q36="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q36" <%if q36="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q36" <%if q36="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q36" <%if q36="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q36" <%if q36="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>37</td>
	<td>I have clear goals for improving my English skills. <BR>���p��W�i�^�y��O�ڦ��۷�M�����ؼСC</td>
	<td><input type="radio" checked value="" id="q37_0" name="q37"></td>
	<td><input type="radio" value="1" name="q37" <%if q37="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q37" <%if q37="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q37" <%if q37="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q37" <%if q37="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q37" <%if q37="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>38</td>
	<td>I think about my progress in learning English. <BR>�ڷ|�h��ҧڦb�ǲ߭^�y�W���i�B�{�סC</td>
	<td><input type="radio" checked value="" id="q38_0" name="q38"></td>
	<td><input type="radio" value="1" name="q38" <%if q38="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q38" <%if q38="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q38" <%if q38="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q38" <%if q38="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q38" <%if q38="5" then response.write "checked" end if%>></td>
	</tr>
	</table>
	<br><font class="T2" ><B>Part E�G</B></font><br>
	<table width="100%" cellpadding=2 cellspacing=0 border="1" bordercolorlight=#333333 bordercolordark=#dddddd>
	<tr><td width="5%">Item</td><td  width="62%">&nbsp;</td><td  width="8%">�����w</td><td  width="5%">1</td><td  width="5%">2</td><td  width="5%">3</td><td  width="5%"> 4</td><td  width="5%">5</td></tr>
	<tr><td>39</td>
	<td>I try to relax whenever I feel afraid of using English. <BR>�C��ڷP��`�ȭn�έ^�y�ɡA�ڷ|���q���P�C</td>
	<td><input type="radio" checked value="" id="q39_0" name="q39"></td>
	<td><input type="radio" value="1" name="q39" <%if q39="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q39" <%if q39="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q39" <%if q39="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q39" <%if q39="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q39" <%if q39="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>40</td>
	<td>I encourage myself to speak English even when I am afraid of making a mistake. <BR>�Y�ϧګܩȷ|�����A���٬O���y�ۤv�h�}�f���^�y�C</td>
	<td><input type="radio" checked value="" id="q40_0" name="q40"></td>
	<td><input type="radio" value="1" name="q40" <%if q40="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q40" <%if q40="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q40" <%if q40="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q40" <%if q40="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q40" <%if q40="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>41</td>
	<td>I give myself a reward or treat when I do well in English. <BR>��ڦb�^�y�譱���}�n��{�ɡA�ڷ|����ۤv�C</td>
	<td><input type="radio" checked value="" id="q41_0" name="q41"></td>
	<td><input type="radio" value="1" name="q41" <%if q41="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q41" <%if q41="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q41" <%if q41="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q41" <%if q41="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q41" <%if q41="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>42</td>
	<td>I notice if I am tense or nervous when I am studying or using English. <BR>�ڷ|�`�N��ڦb��Ū�Ψϥέ^�y�ɬO�_�|��i�C</td>
	<td><input type="radio" checked value="" id="q42_0" name="q42"></td>
	<td><input type="radio" value="1" name="q42" <%if q42="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q42" <%if q42="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q42" <%if q42="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q42" <%if q42="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q42" <%if q42="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>43</td>
	<td>I write down my feelings in a language learning diary. <BR>�ڷ|��ڪ��Pı�O���b�y���ǲߤ�O�̡C</td>
	<td><input type="radio" checked value="" id="q43_0" name="q43"></td>
	<td><input type="radio" value="1" name="q43" <%if q43="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q43" <%if q43="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q43" <%if q43="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q43" <%if q43="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q43" <%if q43="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>44</td>
	<td>I talk to someone else about how I feel when I am learning English. <BR>��ڦb�ǭ^�y�ɡA�ڷ|�i�D�O�H�ڪ��Pı�C</td>
	<td><input type="radio" checked value="" id="q44_0" name="q44"></td>
	<td><input type="radio" value="1" name="q44" <%if q44="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q44" <%if q44="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q44" <%if q44="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q44" <%if q44="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q44" <%if q44="5" then response.write "checked" end if%>></td>
	</tr>
	</table>
	<br><font class="T2" ><B>Part F�G</B></font><br>
	<table width="100%" cellpadding=2 cellspacing=0 border="1" bordercolorlight=#333333 bordercolordark=#dddddd>
	<tr><td width="5%">Item</td><td  width="62%">&nbsp;</td><td  width="8%">�����w</td><td  width="5%">1</td><td  width="5%">2</td><td  width="5%">3</td><td  width="5%"> 4</td><td  width="5%">5</td></tr>
	<tr><td>45</td>
	<td>If I do not understand something in English, I ask the other person to slow down or say it again. <BR>�J��ť�������^��A�ڷ|�ХL��C�t�סA�άO�A���@���C</td>
	<td><input type="radio" checked value="" id="q45_0" name="q45"></td>
	<td><input type="radio" value="1" name="q45" <%if q45="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q45" <%if q45="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q45" <%if q45="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q45" <%if q45="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q45" <%if q45="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>46</td>
	<td>I ask English speakers to correct me when I talk. <BR>��ڻ��^�y�ɡA�ڷ|�ХH�^�y�����y���H�ȥ��ڪ����~�C</td>
	<td><input type="radio" checked value="" id="q46_0" name="q46"></td>
	<td><input type="radio" value="1" name="q46" <%if q46="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q46" <%if q46="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q46" <%if q46="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q46" <%if q46="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q46" <%if q46="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>47</td>
	<td>I practice English with other students. <BR>�ڷ|�M�O���ǥͽm�߭^�y�C</td>
	<td><input type="radio" checked value="" id="q47_0" name="q47"></td>
	<td><input type="radio" value="1" name="q47" <%if q47="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q47" <%if q47="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q47" <%if q47="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q47" <%if q47="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q47" <%if q47="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>48</td>
	<td>I ask for help from English speakers.  <BR>�ڷ|�D�U�H�^�y�����y���H�C</td>
	<td><input type="radio" checked value="" id="q48_0" name="q48"></td>
	<td><input type="radio" value="1" name="q48" <%if q48="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q48" <%if q48="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q48" <%if q48="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q48" <%if q48="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q48" <%if q48="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>49</td>
	<td>I ask questions in English.  <BR>�ڷ|�H�^�y�Ӱݰ��D�C</td>
	<td><input type="radio" checked value="" id="q49_0" name="q49"></td>
	<td><input type="radio" value="1" name="q49" <%if q49="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q49" <%if q49="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q49" <%if q49="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q49" <%if q49="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q49" <%if q49="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>50</td>
	<td>I try to learn about the culture of English speakers. <BR>�ڷ|�Q���D�^�y�t��a����ơC</td>
	<td><input type="radio" checked value="" id="q50_0" name="q50"></td>
	<td><input type="radio" value="1" name="q50" <%if q50="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q50" <%if q50="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q50" <%if q50="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q50" <%if q50="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q50" <%if q50="5" then response.write "checked" end if%>></td>
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
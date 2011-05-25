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
<title>【LDCC英外語能力診斷輔導中心】</title>


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
        errmsg += "You have item not to answer. (你有答案未選擇，項目一)\n";
	if (form1.q2_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目二)\n";
	if (form1.q3_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目三)\n";
	if (form1.q4_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目四)\n";
	if (form1.q5_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目五)\n";
	if (form1.q6_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目六)\n";
	if (form1.q7_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目七)\n";
	if (form1.q8_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目八)\n";
	if (form1.q9_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目九)\n";
	if (form1.q10_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目十)\n";
	if (form1.q11_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目十一)\n";
	if (form1.q12_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目十二)\n";
	if (form1.q13_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目十三)\n";
	if (form1.q14_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目十四)\n";
	if (form1.q15_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目十五)\n";
	if (form1.q16_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目十六)\n";
	if (form1.q17_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目十七)\n";
	if (form1.q18_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目十八)\n";
	if (form1.q19_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目十九)\n";
	if (form1.q20_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目二十)\n";
	if (form1.q21_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目二十一)\n";
	
	if (form1.q22_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目二十二)\n";
	if (form1.q23_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目二十三)\n";
	if (form1.q24_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目二十四)\n";
	if (form1.q25_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目二十五)\n";
	if (form1.q26_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目二十六)\n";
	if (form1.q27_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目二十七)\n";
	if (form1.q28_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目二十八)\n";
	if (form1.q29_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目二十九)\n";
	if (form1.q30_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目三十)\n";

	if (form1.q31_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目三十一)\n";
	if (form1.q32_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目三十二)\n";
	if (form1.q33_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目三十三)\n";
	if (form1.q34_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目三十四)\n";
	if (form1.q35_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目三十五)\n";
	if (form1.q36_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目三十六)\n";
	if (form1.q37_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目三十七)\n";
	if (form1.q38_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目三十八)\n";
	if (form1.q39_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目三十九)\n";
	if (form1.q40_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目四十)\n";
	
	if (form1.q41_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目四十一)\n";
	if (form1.q42_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目四十二)\n";
	if (form1.q43_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目四十三)\n";
	if (form1.q44_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目四十四)\n";
	if (form1.q45_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目四十五)\n";
	if (form1.q46_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目四十六)\n";
	if (form1.q47_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目四十七)\n";
	if (form1.q48_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目四十八)\n";
	if (form1.q49_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目四十九)\n";
	if (form1.q50_0.checked==true)
        errmsg += "You have item not to answer. (你有答案未選擇，項目五十)\n";
	


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
	if (confirm('你尚未儲存答案，確定要離開嗎？'))
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
alert("請尊重智慧財產權，謝謝。\n");
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
<font  color="#FFFFFF">※STRATEGY INVENTORY FOR LANGUAGE LEARNING (SILL) ※</font></td></tr>
<tr valign="top"><td bgcolor=#ECECE3>
<table width="780" align="center"  >
<tr><td  height=35 align="center" class="T2">
<%if validate="query" then%>
<center><input type="button" value="離開"  onclick="window.close();" class="inputbutton" >&nbsp;&nbsp;
填寫時間：<%=initdate%>
<%else%>
注意事項：1. 問卷本身為英文版，請以閱讀英文題目為主，不懂時再參考中文<BR>
											 2. 所有題目皆以學習英文為主
<%end if%>

</td></tr>
<tr><td  valign="top" align="center">
<BR><BR>
	<form id="form1" name="form1" method="post" action="qstrategy.asp" onsubmit="return check_input();">
	<input type="hidden" value="add" name="validate">
	<input type="hidden" value="<%=sid%>" name="sid">
	<font class="T2" ><B>Part A：</B></font><br>
	<table width="100%" cellpadding=2 cellspacing=0 border="1" bordercolorlight=#333333 bordercolordark=#dddddd>
	<tr><td width="5%">Item</td><td  width="62%">&nbsp;</td><td  width="8%">未指定</td><td  width="5%">1</td><td  width="5%">2</td><td  width="5%">3</td><td  width="5%"> 4</td><td  width="5%">5</td></tr>
	<tr><td>1</td>
	<td>I think of relationships between what I already know and new things I learn in English.<BR>我會去思考學過的和新學的英語之間的關係。</td>
	<td><input type="radio" checked value="" id="q1_0" name="q1"></td>
	<td><input type="radio" value="1" name="q1" <%if q1="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q1" <%if q1="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q1" <%if q1="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q1" <%if q1="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q1" <%if q1="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>2</td>
	<td>I use new English words in a sentence so I can remember them. <BR>為了記住新學的英語單字，我會試著用這些生字來造句。</td>
	<td><input type="radio" checked value="" id="q2_0" name="q2"></td>
	<td><input type="radio" value="1" name="q2" <%if q2="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q2" <%if q2="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q2" <%if q2="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q2" <%if q2="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q2" <%if q2="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>3</td>
	<td>I connect the sound of a new English word and an image or picture of the word to help me remember the word. <BR>我會在腦海中想出可以配合英語聲音的圖片或意象，以便記住某個單字。 </td>
	<td><input type="radio" checked value="" id="q3_0" name="q3"></td>
	<td><input type="radio" value="1" name="q3" <%if q3="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q3" <%if q3="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q3" <%if q3="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q3" <%if q3="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q3" <%if q3="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>4</td>
	<td>I remember a new English word by making a mental picture of a situation in which the word might be used. <BR>我會在腦中製造出某個生字出現的情境，以這種方法把單字背起來。</td>
	<td><input type="radio" checked value="" id="q4_0" name="q4"></td>
	<td><input type="radio" value="1" name="q4" <%if q4="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q4" <%if q4="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q4" <%if q4="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q4" <%if q4="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q4" <%if q4="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>5</td>
	<td>I use rhymes to remember new English words. <BR>我會使用押韻的方式來記住生字。</td>
	<td><input type="radio" checked value="" id="q5_0" name="q5"></td>
	<td><input type="radio" value="1" name="q5" <%if q5="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q5" <%if q5="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q5" <%if q5="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q5" <%if q5="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q5" <%if q5="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>6</td>
	<td>I use flashcards to remember new English words.<BR>我會使用閃示卡來背生字。</td>
	<td><input type="radio" checked value="" id="q6_0" name="q6"></td>
	<td><input type="radio" value="1" name="q6" <%if q6="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q6" <%if q6="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q6" <%if q6="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q6" <%if q6="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q6" <%if q6="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>7</td>
	<td>I physically act out new English words.<BR>我會把生字用肢體演出來。 </td>
	<td><input type="radio" checked value="" id="q7_0" name="q7"></td>
	<td><input type="radio" value="1" name="q7" <%if q7="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q7" <%if q7="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q7" <%if q7="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q7" <%if q7="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q7" <%if q7="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>8</td>
	<td>I review English lessons often. <BR>我常常複習英語課程。</td>
	<td><input type="radio" checked value="" id="q8_0" name="q8"></td>
	<td><input type="radio" value="1" name="q8" <%if q8="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q8" <%if q8="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q8" <%if q8="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q8" <%if q8="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q8" <%if q8="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>9</td>
	<td>I remember new English words or phrases by remembering their location on the page, on the board, or on a street sign.<BR>我會按照生字或片語出現在課本、黑板或是街道看板的位置，來記住生字或片語。</td>
	<td><input type="radio" checked value="" id="q9_0" name="q9"></td>
	<td><input type="radio" value="1" name="q9" <%if q9="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q9" <%if q9="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q9" <%if q9="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q9" <%if q9="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q9" <%if q9="5" then response.write "checked" end if%>></td>
	</tr>
	</table>
	<br><font class="T2" ><B>Part B：</B></font><br>
	<table width="100%" cellpadding=2 cellspacing=0 border="1" bordercolorlight=#333333 bordercolordark=#dddddd>
	<tr><td width="5%">Item</td><td  width="62%">&nbsp;</td><td  width="8%">未指定</td><td  width="5%">1</td><td  width="5%">2</td><td  width="5%">3</td><td  width="5%"> 4</td><td  width="5%">5</td></tr>
	<tr><td>10</td>
	<td>I say or write new English words several times.<BR>我會重複說或寫英文生字好幾次。</td>
	<td><input type="radio" checked value="" id="q10_0" name="q10"></td>
	<td><input type="radio" value="1" name="q10" <%if q10="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q10" <%if q10="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q10" <%if q10="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q10" <%if q10="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q10" <%if q10="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>11</td>
	<td>I try to talk like native English speakers. <BR>我會想把英語說得像以英語為母語的人一樣。</td>
	<td><input type="radio" checked value="" id="q11_0" name="q11"></td>
	<td><input type="radio" value="1" name="q11" <%if q11="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q11" <%if q11="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q11" <%if q11="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q11" <%if q11="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q11" <%if q11="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>12</td>
	<td>I practice the sounds of English. <BR>我會練習英語的發音。</td>
	<td><input type="radio" checked value="" id="q12_0" name="q12"></td>
	<td><input type="radio" value="1" name="q12" <%if q12="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q12" <%if q12="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q12" <%if q12="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q12" <%if q12="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q12" <%if q12="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>13</td>
	<td>I use the English words I know in different ways. <BR>我會把學過的英文字用在不同的方面上。</td>
	<td><input type="radio" checked value="" id="q13_0" name="q13"></td>
	<td><input type="radio" value="1" name="q13"  <%if q13="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q13"  <%if q13="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q13"  <%if q13="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q13"  <%if q13="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q13"  <%if q13="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>14</td>
	<td>I start conversations in English. <BR>我會以英語開啟對話。</td>
	<td><input type="radio" checked value="" id="q14_0" name="q14"></td>
	<td><input type="radio" value="1" name="q14" <%if q14="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q14" <%if q14="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q14" <%if q14="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q14" <%if q14="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q14" <%if q14="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>15</td>
	<td>I watch English TV shows spoken in English or go to movies spoken in English. <BR>我會看以英語發音的電視節目或電影。</td>
	<td><input type="radio" checked value="" id="q15_0" name="q15"></td>
	<td><input type="radio" value="1" name="q15" <%if q15="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q15" <%if q15="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q15" <%if q15="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q15" <%if q15="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q15" <%if q15="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>16</td>
	<td>I read for pleasure in English. <BR>我閱讀英文做為休閒活動。</td>
	<td><input type="radio" checked value="" id="q16_0" name="q16"></td>
	<td><input type="radio" value="1" name="q16" <%if q16="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q16" <%if q16="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q16" <%if q16="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q16" <%if q16="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q16" <%if q16="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>17</td>
	<td>I write notes, messages, letters, or reports in English. <BR>我會以英語來記筆記、訊息、書信或是報告。</td>
	<td><input type="radio" checked value="" id="q17_0" name="q17"></td>
	<td><input type="radio" value="1" name="q17" <%if q17="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q17" <%if q17="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q17" <%if q17="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q17" <%if q17="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q17" <%if q17="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>18</td>
	<td>I first skim an English passage (read over the passage quickly) then go back and read carefully. <BR>我會先略讀英語的文章(很快地把文章看過一遍)，然後再回來細看。</td>
	<td><input type="radio" checked value="" id="q18_0" name="q18"></td>
	<td><input type="radio" value="1" name="q18"  <%if q18="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q18" <%if q18="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q18" <%if q18="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q18" <%if q18="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q18" <%if q18="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>19</td>
	<td>I look for words in my own language that are similar to new words in English.<BR>我會在我的母語裡找尋和英語相同的生字。</td>
	<td><input type="radio" checked value="" id="q19_0" name="q19"></td>
	<td><input type="radio" value="1" name="q19" <%if q19="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q19" <%if q19="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q19" <%if q19="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q19" <%if q19="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q19" <%if q19="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>20</td>
	<td>I try to find patterns in English. <BR>我會找出英語的模式。</td>
	<td><input type="radio" checked value="" id="q20_0" name="q20"></td>
	<td><input type="radio" value="1" name="q20" <%if q20="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q20" <%if q20="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q20" <%if q20="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q20" <%if q20="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q20" <%if q20="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>21</td>
	<td>I find the meaning of an English word by dividing it into parts that I understand.<BR>我會把英語拆解開來，找出自己懂的部份，藉以了解單字的意思。</td>
	<td><input type="radio" checked value="" id="q21_0" name="q21"></td>
	<td><input type="radio" value="1" name="q21"  <%if q21="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q21"  <%if q21="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q21"  <%if q21="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q21"  <%if q21="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q21"  <%if q21="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>22</td>
	<td>I try not to translate word-for-word. <BR>我不會逐字逐句翻譯。</td>
	<td><input type="radio" checked value="" id="q22_0" name="q22"></td>
	<td><input type="radio" value="1" name="q22" <%if q22="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q22" <%if q22="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q22" <%if q22="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q22" <%if q22="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q22" <%if q22="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>23</td>
	<td>I make summaries of information that I hear or read in English. <BR>我會把聽到或是讀到的英文資訊做成摘要。</td>
	<td><input type="radio" checked value="" id="q23_0" name="q23"></td>
	<td><input type="radio" value="1" name="q23" <%if q23="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q23" <%if q23="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q23" <%if q23="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q23" <%if q23="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q23" <%if q23="5" then response.write "checked" end if%>></td>
	</tr>
	</table>
	<br><font class="T2" ><B>Part C：</B></font><br>
	<table width="100%" cellpadding=2 cellspacing=0 border="1" bordercolorlight=#333333 bordercolordark=#dddddd>
	<tr><td width="5%">Item</td><td  width="62%">&nbsp;</td><td  width="8%">未指定</td><td  width="5%">1</td><td  width="5%">2</td><td  width="5%">3</td><td  width="5%"> 4</td><td  width="5%">5</td></tr>
	<tr><td>24</td>
	<td>To understand unfamiliar English words, I make guesses.<BR>遇到不熟悉的英文單字，我會去猜它的意思。</td>
	<td><input type="radio" checked value="" id="q24_0" name="q24"></td>
	<td><input type="radio" value="1" name="q24" <%if q24="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q24" <%if q24="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q24" <%if q24="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q24" <%if q24="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q24" <%if q24="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>25</td>
	<td>When I can think of a word during a conversation in English, I use gestures.<BR>在對話中，我如果想不出某個字英文怎麼說，我會使用表情和動作。</td>
	<td><input type="radio" checked value="" id="q25_0" name="q25"></td>
	<td><input type="radio" value="1" name="q25" <%if q25="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q25" <%if q25="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q25" <%if q25="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q25" <%if q25="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q25" <%if q25="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>26</td>
	<td>I make up new words if I do not know the right ones in English. <BR>如果我不知道英語該怎麼說，我會自己創造新字。</td>
	<td><input type="radio" checked value="" id="q26_0" name="q26"></td>
	<td><input type="radio" value="1" name="q26" <%if q26="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q26" <%if q26="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q26" <%if q26="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q26" <%if q26="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q26" <%if q26="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>27</td>
	<td>I read English without looking up every new word.<BR>閱讀的過程中，我一遇到生字就馬上查字典。</td>
	<td><input type="radio" checked value="" id="q27_0" name="q27"></td>
	<td><input type="radio" value="1" name="q27" <%if q27="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q27" <%if q27="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q27" <%if q27="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q27" <%if q27="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q27" <%if q27="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>28</td>
	<td>I try to guess what the other person will say next in English.<BR>我會用英語試著去猜別人接著會說什麼。</td>
	<td><input type="radio" checked value="" id="q28_0" name="q28"></td>
	<td><input type="radio" value="1" name="q28" <%if q28="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q28" <%if q28="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q28" <%if q28="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q28" <%if q28="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q28" <%if q28="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>29</td>
	<td>I can think of an English word, I use a word or phrase that means the same thing. <BR>如果我想不起來某個英文單字，我會用別的字或片語來轉述同樣的意思。</td>
	<td><input type="radio" checked value="" id="q29_0" name="q29"></td>
	<td><input type="radio" value="1" name="q29" <%if q29="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q29" <%if q29="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q29" <%if q29="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q29" <%if q29="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q29" <%if q29="5" then response.write "checked" end if%>></td>
	</tr>
	</table>
	<br><font class="T2" ><B>Part D：</B></font><br>
	<table width="100%" cellpadding=2 cellspacing=0 border="1" bordercolorlight=#333333 bordercolordark=#dddddd>
	<tr><td width="5%">Item</td><td  width="62%">&nbsp;</td><td  width="8%">未指定</td><td  width="5%">1</td><td  width="5%">2</td><td  width="5%">3</td><td  width="5%"> 4</td><td  width="5%">5</td></tr>
	<tr><td>30</td>
	<td>I try to find as many ways as I can to use my English.<BR>我會儘量找機會練習英語。</td>
	<td><input type="radio" checked value="" id="q30_0" name="q30"></td>
	<td><input type="radio" value="1" name="q30" <%if q30="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q30" <%if q30="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q30" <%if q30="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q30" <%if q30="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q30" <%if q30="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>31</td>
	<td>I notice my English mistakes and I use that information to help me do better.<BR>我會注意我所犯的錯誤，藉此幫助自己學得更好。</td>
	<td><input type="radio" checked value="" id="q31_0" name="q31"></td>
	<td><input type="radio" value="1" name="q31" <%if q31="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q31" <%if q31="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q31" <%if q31="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q31" <%if q31="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q31" <%if q31="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>32</td>
	<td>I pay attention when someone is speaking English. <BR>當有人在說英語時，會引起我的注意。</td>
	<td><input type="radio" checked value="" id="q32_0" name="q32"></td>
	<td><input type="radio" value="1" name="q32" <%if q32="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q32" <%if q32="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q32" <%if q32="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q32" <%if q32="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q32" <%if q32="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>33</td>
	<td>I try to find out how to be a better learner of English. <BR>我會想辦法讓自己成為更好的英語學習者。</td>
	<td><input type="radio" checked value="" id="q33_0" name="q33"></td>
	<td><input type="radio" value="1" name="q33" <%if q33="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q33" <%if q33="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q33" <%if q33="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q33" <%if q33="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q33" <%if q33="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>34</td>
	<td>I plan my schedule so I will have enough time to study English.<BR>我會好好規畫時間，以便有足夠的時間學英語。</td>
	<td><input type="radio" checked value="" id="q34_0" name="q34"></td>
	<td><input type="radio" value="1" name="q34" <%if q34="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q34" <%if q34="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q34" <%if q34="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q34" <%if q34="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q34" <%if q34="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>35</td>
	<td>I look for people I can talk to in English.  <BR>我會找能用英語談話的人練習英語。</td>
	<td><input type="radio" checked value="" id="q35_0" name="q35"></td>
	<td><input type="radio" value="1" name="q35" <%if q35="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q35" <%if q35="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q35" <%if q35="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q35" <%if q35="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q35" <%if q35="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>36</td>
	<td>I look for opportunities to read as much as possible in English. <BR>我會儘量找機會閱讀英語。</td>
	<td><input type="radio" checked value="" id="q36_0" name="q36"></td>
	<td><input type="radio" value="1" name="q36" <%if q36="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q36" <%if q36="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q36" <%if q36="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q36" <%if q36="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q36" <%if q36="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>37</td>
	<td>I have clear goals for improving my English skills. <BR>對於如何增進英語能力我有相當清楚的目標。</td>
	<td><input type="radio" checked value="" id="q37_0" name="q37"></td>
	<td><input type="radio" value="1" name="q37" <%if q37="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q37" <%if q37="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q37" <%if q37="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q37" <%if q37="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q37" <%if q37="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>38</td>
	<td>I think about my progress in learning English. <BR>我會去思考我在學習英語上的進步程度。</td>
	<td><input type="radio" checked value="" id="q38_0" name="q38"></td>
	<td><input type="radio" value="1" name="q38" <%if q38="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q38" <%if q38="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q38" <%if q38="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q38" <%if q38="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q38" <%if q38="5" then response.write "checked" end if%>></td>
	</tr>
	</table>
	<br><font class="T2" ><B>Part E：</B></font><br>
	<table width="100%" cellpadding=2 cellspacing=0 border="1" bordercolorlight=#333333 bordercolordark=#dddddd>
	<tr><td width="5%">Item</td><td  width="62%">&nbsp;</td><td  width="8%">未指定</td><td  width="5%">1</td><td  width="5%">2</td><td  width="5%">3</td><td  width="5%"> 4</td><td  width="5%">5</td></tr>
	<tr><td>39</td>
	<td>I try to relax whenever I feel afraid of using English. <BR>每當我感到害怕要用英語時，我會儘量放輕鬆。</td>
	<td><input type="radio" checked value="" id="q39_0" name="q39"></td>
	<td><input type="radio" value="1" name="q39" <%if q39="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q39" <%if q39="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q39" <%if q39="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q39" <%if q39="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q39" <%if q39="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>40</td>
	<td>I encourage myself to speak English even when I am afraid of making a mistake. <BR>即使我很怕會說錯，我還是鼓勵自己多開口說英語。</td>
	<td><input type="radio" checked value="" id="q40_0" name="q40"></td>
	<td><input type="radio" value="1" name="q40" <%if q40="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q40" <%if q40="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q40" <%if q40="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q40" <%if q40="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q40" <%if q40="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>41</td>
	<td>I give myself a reward or treat when I do well in English. <BR>當我在英語方面有良好表現時，我會犒賞自己。</td>
	<td><input type="radio" checked value="" id="q41_0" name="q41"></td>
	<td><input type="radio" value="1" name="q41" <%if q41="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q41" <%if q41="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q41" <%if q41="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q41" <%if q41="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q41" <%if q41="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>42</td>
	<td>I notice if I am tense or nervous when I am studying or using English. <BR>我會注意當我在研讀或使用英語時是否會緊張。</td>
	<td><input type="radio" checked value="" id="q42_0" name="q42"></td>
	<td><input type="radio" value="1" name="q42" <%if q42="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q42" <%if q42="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q42" <%if q42="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q42" <%if q42="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q42" <%if q42="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>43</td>
	<td>I write down my feelings in a language learning diary. <BR>我會把我的感覺記錄在語言學習日記裡。</td>
	<td><input type="radio" checked value="" id="q43_0" name="q43"></td>
	<td><input type="radio" value="1" name="q43" <%if q43="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q43" <%if q43="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q43" <%if q43="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q43" <%if q43="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q43" <%if q43="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>44</td>
	<td>I talk to someone else about how I feel when I am learning English. <BR>當我在學英語時，我會告訴別人我的感覺。</td>
	<td><input type="radio" checked value="" id="q44_0" name="q44"></td>
	<td><input type="radio" value="1" name="q44" <%if q44="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q44" <%if q44="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q44" <%if q44="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q44" <%if q44="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q44" <%if q44="5" then response.write "checked" end if%>></td>
	</tr>
	</table>
	<br><font class="T2" ><B>Part F：</B></font><br>
	<table width="100%" cellpadding=2 cellspacing=0 border="1" bordercolorlight=#333333 bordercolordark=#dddddd>
	<tr><td width="5%">Item</td><td  width="62%">&nbsp;</td><td  width="8%">未指定</td><td  width="5%">1</td><td  width="5%">2</td><td  width="5%">3</td><td  width="5%"> 4</td><td  width="5%">5</td></tr>
	<tr><td>45</td>
	<td>If I do not understand something in English, I ask the other person to slow down or say it again. <BR>遇到聽不懂的英文，我會請他放慢速度，或是再講一次。</td>
	<td><input type="radio" checked value="" id="q45_0" name="q45"></td>
	<td><input type="radio" value="1" name="q45" <%if q45="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q45" <%if q45="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q45" <%if q45="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q45" <%if q45="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q45" <%if q45="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>46</td>
	<td>I ask English speakers to correct me when I talk. <BR>當我說英語時，我會請以英語為母語的人糾正我的錯誤。</td>
	<td><input type="radio" checked value="" id="q46_0" name="q46"></td>
	<td><input type="radio" value="1" name="q46" <%if q46="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q46" <%if q46="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q46" <%if q46="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q46" <%if q46="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q46" <%if q46="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>47</td>
	<td>I practice English with other students. <BR>我會和別的學生練習英語。</td>
	<td><input type="radio" checked value="" id="q47_0" name="q47"></td>
	<td><input type="radio" value="1" name="q47" <%if q47="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q47" <%if q47="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q47" <%if q47="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q47" <%if q47="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q47" <%if q47="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>48</td>
	<td>I ask for help from English speakers.  <BR>我會求助以英語為母語的人。</td>
	<td><input type="radio" checked value="" id="q48_0" name="q48"></td>
	<td><input type="radio" value="1" name="q48" <%if q48="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q48" <%if q48="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q48" <%if q48="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q48" <%if q48="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q48" <%if q48="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>49</td>
	<td>I ask questions in English.  <BR>我會以英語來問問題。</td>
	<td><input type="radio" checked value="" id="q49_0" name="q49"></td>
	<td><input type="radio" value="1" name="q49" <%if q49="1" then response.write "checked" end if%>></td>
	<td><input type="radio" value="2" name="q49" <%if q49="2" then response.write "checked" end if%>></td>
	<td><input type="radio" value="3" name="q49" <%if q49="3" then response.write "checked" end if%>></td>
	<td><input type="radio" value="4" name="q49" <%if q49="4" then response.write "checked" end if%>></td>
	<td><input type="radio" value="5" name="q49" <%if q49="5" then response.write "checked" end if%>></td>
	</tr>
	<tr><td>50</td>
	<td>I try to learn about the culture of English speakers. <BR>我會想知道英語系國家的文化。</td>
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
<center><input type="button" value="離開"  id="btn_close" onclick="wclose();" class="inputbutton" >&nbsp;&nbsp;
<input type="submit" value="送出答案"  class="inputbutton" >
<%else%>
<center><input type="button" value="離開"  onclick="window.close();" class="inputbutton" >&nbsp;&nbsp;

<%end if%>
<BR><BR>
</form>
</td></tr>
<tr><td bgcolor=#555555 height=24 align=right><font Color="#FFFFFF">預約相關問題請洽LDCC--許蕙婷 分機7403 </font></td></tr></table>

</body>
</html>
<!-- #INCLUDE virtual="/include/lib/conn/msconnclose.asp" -->
<%@Language=VBScript LCID=1033%>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<%
sid = trim(request("sid"))
%>

<html>
<head>
<title>【LDCC英外語能力診斷輔導中心】</title>


<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<script language="javascript">

function btn_status()
{
	var obj;
	obj= document.getElementById("btn_start");
	obj.disabled=false;
}

</script>
</head>
<body bottomMargin=0 leftMargin=0 topMargin=0 rightMargin=0 marginheight="0" marginwidth="0" >

<table width="100%" height="100%" border=1 cellpadding=0 cellspacing=0 bordercolorlight=#333333 bordercolordark=#dddddd>
<tr><td bgcolor=555555 height=35 align="center" class="T2">
<font  color="#FFFFFF">※Perceptual Learning-Style Preference Questionnaire ※</font></td></tr>
<tr valign="top"><td bgcolor=#ECECE3>
<table width="780" align="center"  >
<tr><td  height=35 align="center" class="T2">Direction</td></tr>
<tr><td  valign="top" align="center">
<BR><BR>
	<table width="75%" cellpadding=3 cellspacing=3>
	<tr><td></td></tr>
	<tr><td>This questionnaire has been designed to help you identify the way(s) you learn best--the way(s) you prefer to learn.</td></tr>
	<tr><td>Read each statement on the following pages. Please respond to the statements <font color="blue"><B>AS THEY APPLY TO YOUR STUDY OF ENGLISH</B></font>.</td></tr>
	<tr><td>Decide whether you agree or disagree with each statement. For example, if you strong agree, mark:</td></tr>
	<tr><td>
	<table width="100%" height="100%" border=1 cellpadding=2 cellspacing=2 bordercolorlight=#333333 bordercolordark=#dddddd>
	<tr><td width="20%">SD<br>Strongly Disagree</td><td width="20%">D<br>Disagree</td><td width="20%">U<br>Undecided</td>
	<td width="20%">A<br>Agree</td><td width="20%">SA<br>Strongly Agree</td></tr>
	<tr align="center"><td><input type="radio" value="1" name="q1"></td><td><input type="radio" value="2" name="q1"></td>
	<td><input type="radio" value="3" name="q1"></td><td><input type="radio" value="4" name="q1"></td><td><input type="radio" value="5" checked name="q1"></td></tr>
	</table>
	</td></tr>
	<tr><td>Please respond to each statement quickly, without too much thought. Try not to change your responses after you choose them. Please answer all the questions. Please use a pen to mark your choices. </td></tr>
	

	</table>

</td></tr>
<tr><td  height=35 align="center" class="T2">說明</td></tr>
<tr><td  valign="top" align="center">
<BR><BR>
	<table width="75%" cellpadding=3 cellspacing=3>
	<tr><td>本問卷設計是為了幫助同學瞭解自己在學習英文時較常使用或傾向使用的學習方法。請仔細閱讀每一項陳述，並<font color="blue"><B>根據你學習英文的現況來回答問題</B></font>。回答問題請在答案欄中勾取您的答案(5、4、3、2、1)。例如，你的答案是非常同意，請做答如下:</td></tr>
	<tr><td>
	<table width="100%" height="100%" border=1 cellpadding=2 cellspacing=2 bordercolorlight=#333333 bordercolordark=#dddddd>
	<tr><td width="20%">SD<br>Strongly Disagree</td><td width="20%">D<br>Disagree</td><td width="20%">U<br>Undecided</td>
	<td width="20%">A<br>Agree</td><td width="20%">SA<br>Strongly Agree</td></tr>
	<tr align="center"><td><input type="radio" value="1" name="q2"></td><td><input type="radio" value="2" name="q2"></td>
	<td><input type="radio" value="3" name="q2"></td><td><input type="radio" value="4" name="q2"></td><td><input type="radio" value="5" checked name="q2"></td></tr>
	</table>
	</td></tr>
	<tr><td>請勿花太多時間思考，盡量迅速作答，也請避免在作答後更改答案。 </td></tr>
	

	</table>

</td></tr>
</table>
<BR><BR>
<center><input type="checkbox" onclick="btn_status();">我已詳細閱讀<input type="button" value="開始作答" onclick="window.location.href='qstyle.asp?sid=<%=sid%>'" id="btn_start" class="inputbutton" disabled>&nbsp;&nbsp;<input type="button" value="離開"  id="btn_close" onclick="window.close();" class="inputbutton" >
<BR><BR>
</td></tr>
<tr><td bgcolor=#555555 height=24 align=right><font Color="#FFFFFF">預約相關問題請洽LDCC--許蕙婷 分機7403 </font></td></tr></table>

</body>
</html>

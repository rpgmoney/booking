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
<font  color="#FFFFFF">※STRATEGY INVENTORY FOR LANGUAGE LEARNING (SILL) ※</font></td></tr>
<tr valign="top"><td bgcolor=#ECECE3>
<table width="780" align="center" >
<tr><td  height=35 align="center" class="T2"><BR>語言學習策略查核表<BR><BR>Direction 說明</td></tr>
<tr><td  valign="top" align="center">
<BR>
	<table width="100%" cellpadding=3 cellspacing=3>
	<tr><td></td></tr>
	<tr><td>This form of the Strategy Inventory for Language Learning (SILL) is for students of English as a second or foreign language. You will find statements about learning English. Please read each one and write the response (1, 2, 3, 4 or 5) that tells HOW TRUE OF YOU THE STATEMENT IS on the worksheet for answering and scoring.</td></tr>
	<tr><td>這份語言學習策略查核表是為EFL學生所設計。內容關於英語學習狀況等陳述。請仔細閱讀每項陳述。依據每一項陳述對於你的真實性，把答案(1, 2, 3, 4, 5)寫在答案分數工作單上。</td></tr>
	<tr><td>
	<BR>
		<table width="90%" height="100%" border=1 cellpadding=2 cellspacing=2 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td>1.</td><td>Never or almost never true of me.</td><td>我從來都沒有或是幾乎沒有。</td></tr>
		<tr><td>2.</td><td>Usually not true of me.</td><td>我通常沒有。</td></tr>
		<tr><td>3.</td><td>Somewhat true of me.</td><td>有點像我。</td></tr>
		<tr><td>4.</td><td>Usually true of me.</td><td>我通常是這樣。</td></tr>
		<tr><td>5.</td><td>Always or almost always true of me.</td><td>我一直都是這樣，或是幾乎一向如此。</td></tr>
		</table>
	</td></tr>
	<tr><td>
	<BR>
		<table width="90%" height="100%" border=0 cellpadding=2 cellspacing=2 >
		<tr><td>1.</td><td><font color="blue"><B>NEVER OR ALMOST NEVER TRUE OF ME</B></font> means that the statement is very rarely true of yoy.</td></tr>
		<tr><td></td><td>「我從來都沒有或是幾乎沒有」表示該陳述的正確性很低。</td></tr>
		<tr><td>2.</td><td> <font color="blue"><B>USUALLY NOT TRUE OF ME</B></font> means that the statement is true less than half the time.</td></tr>
		<tr><td></td><td>「我通常沒有」表示該陳述的正確性沒有超過一半。</td></tr>
		<tr><td>3.</td><td><font color="blue"><B>SOMEWHAT TRUE OF ME</B></font> means that the statement is true of you about half the time.</td></tr>
		<tr><td></td><td>「有點像我」表示該陳述的正確性為一半。</td></tr>
		<tr><td>4.</td><td><font color="blue"><B>USUALLY TRUE OF ME</B></font> means that the statement is true more than half the time.</td></tr>
		<tr><td></td><td>「我通常是這樣」表示該陳述的正確性已超過一半。</td></tr>
		<tr><td>5.</td><td><font color="blue"><B>ALWAYS OR ALMOST ALWAYS TRUE OF ME</B></font> means that the statement is true of you almost always.</td></tr>
		<tr><td></td><td>「我一直都是這樣，或是幾乎一向如此」表示該陳述的正確性幾乎百分之百。</td></tr>
		</table>
	</td></tr>
	</table>
</td></tr>
<tr><td  valign="top" align="center">
<BR>
	<table width="100%" cellpadding=3 cellspacing=3>
	<tr><td>Answer in terms of how well the statement describes you. Do not answer how you think you should be, or what other people do. There are no right or wrong answers to these statements. Work as quickly as you can without being careless. This usually takes about 20-30 minutes to complete. If you have any questions, let the teacher know immediately.</td></tr>
	<tr><td>你的回答是根據該陳述有多麼像你的程度。不要依照你認為自己應該是什麼樣子或是別人是怎麼認為的來回答。這些陳述並沒有對或錯的標準答案。在謹慎小心的情況下，快速作答。這份問卷通常需花二十到三十分鐘。如果有問題，馬上告知你的老師。 </td></tr>
	</table>
</td></tr>
</table>
<BR><BR>
<center><input type="checkbox" onclick="btn_status();">我已詳細閱讀<input type="button" value="開始作答" onclick="window.location.href='qstrategy.asp?sid=<%=sid%>'" id="btn_start" class="inputbutton" disabled>&nbsp;&nbsp;<input type="button" value="離開"  id="btn_close" onclick="window.close();" class="inputbutton" >
<BR><BR>
</td></tr>
<tr><td bgcolor=#555555 height=24 align=right><font Color="#FFFFFF">預約相關問題請洽LDCC--許蕙婷 分機7403 </font></td></tr></table>

</body>
</html>

<%@Language=VBScript LCID=1033%>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<%
sid = trim(request("sid"))
%>

<html>
<head>
<title>�iLDCC�^�~�y��O�E�_���ɤ��ߡj</title>


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
<font  color="#FFFFFF">��STRATEGY INVENTORY FOR LANGUAGE LEARNING (SILL) ��</font></td></tr>
<tr valign="top"><td bgcolor=#ECECE3>
<table width="780" align="center" >
<tr><td  height=35 align="center" class="T2"><BR>�y���ǲߵ����d�֪�<BR><BR>Direction ����</td></tr>
<tr><td  valign="top" align="center">
<BR>
	<table width="100%" cellpadding=3 cellspacing=3>
	<tr><td></td></tr>
	<tr><td>This form of the Strategy Inventory for Language Learning (SILL) is for students of English as a second or foreign language. You will find statements about learning English. Please read each one and write the response (1, 2, 3, 4 or 5) that tells HOW TRUE OF YOU THE STATEMENT IS on the worksheet for answering and scoring.</td></tr>
	<tr><td>�o���y���ǲߵ����d�֪�O��EFL�ǥͩҳ]�p�C���e����^�y�ǲߪ��p�����z�C�ХJ�Ӿ\Ū�C�����z�C�̾ڨC�@�����z���A���u��ʡA�⵪��(1, 2, 3, 4, 5)�g�b���פ��Ƥu�@��W�C</td></tr>
	<tr><td>
	<BR>
		<table width="90%" height="100%" border=1 cellpadding=2 cellspacing=2 bordercolorlight=#333333 bordercolordark=#dddddd>
		<tr><td>1.</td><td>Never or almost never true of me.</td><td>�ڱq�ӳ��S���άO�X�G�S���C</td></tr>
		<tr><td>2.</td><td>Usually not true of me.</td><td>�ڳq�`�S���C</td></tr>
		<tr><td>3.</td><td>Somewhat true of me.</td><td>���I���ڡC</td></tr>
		<tr><td>4.</td><td>Usually true of me.</td><td>�ڳq�`�O�o�ˡC</td></tr>
		<tr><td>5.</td><td>Always or almost always true of me.</td><td>�ڤ@�����O�o�ˡA�άO�X�G�@�V�p���C</td></tr>
		</table>
	</td></tr>
	<tr><td>
	<BR>
		<table width="90%" height="100%" border=0 cellpadding=2 cellspacing=2 >
		<tr><td>1.</td><td><font color="blue"><B>NEVER OR ALMOST NEVER TRUE OF ME</B></font> means that the statement is very rarely true of yoy.</td></tr>
		<tr><td></td><td>�u�ڱq�ӳ��S���άO�X�G�S���v��ܸӳ��z�����T�ʫܧC�C</td></tr>
		<tr><td>2.</td><td> <font color="blue"><B>USUALLY NOT TRUE OF ME</B></font> means that the statement is true less than half the time.</td></tr>
		<tr><td></td><td>�u�ڳq�`�S���v��ܸӳ��z�����T�ʨS���W�L�@�b�C</td></tr>
		<tr><td>3.</td><td><font color="blue"><B>SOMEWHAT TRUE OF ME</B></font> means that the statement is true of you about half the time.</td></tr>
		<tr><td></td><td>�u���I���ڡv��ܸӳ��z�����T�ʬ��@�b�C</td></tr>
		<tr><td>4.</td><td><font color="blue"><B>USUALLY TRUE OF ME</B></font> means that the statement is true more than half the time.</td></tr>
		<tr><td></td><td>�u�ڳq�`�O�o�ˡv��ܸӳ��z�����T�ʤw�W�L�@�b�C</td></tr>
		<tr><td>5.</td><td><font color="blue"><B>ALWAYS OR ALMOST ALWAYS TRUE OF ME</B></font> means that the statement is true of you almost always.</td></tr>
		<tr><td></td><td>�u�ڤ@�����O�o�ˡA�άO�X�G�@�V�p���v��ܸӳ��z�����T�ʴX�G�ʤ����ʡC</td></tr>
		</table>
	</td></tr>
	</table>
</td></tr>
<tr><td  valign="top" align="center">
<BR>
	<table width="100%" cellpadding=3 cellspacing=3>
	<tr><td>Answer in terms of how well the statement describes you. Do not answer how you think you should be, or what other people do. There are no right or wrong answers to these statements. Work as quickly as you can without being careless. This usually takes about 20-30 minutes to complete. If you have any questions, let the teacher know immediately.</td></tr>
	<tr><td>�A���^���O�ھڸӳ��z���h�򹳧A���{�סC���n�̷ӧA�{���ۤv���ӬO����ˤl�άO�O�H�O���{�����Ӧ^���C�o�ǳ��z�èS����ο����зǵ��סC�b�ԷV�p�ߪ����p�U�A�ֳt�@���C�o���ݨ��q�`�ݪ�G�Q��T�Q�����C�p�G�����D�A���W�i���A���Ѯv�C </td></tr>
	</table>
</td></tr>
</table>
<BR><BR>
<center><input type="checkbox" onclick="btn_status();">�ڤw�ԲӾ\Ū<input type="button" value="�}�l�@��" onclick="window.location.href='qstrategy.asp?sid=<%=sid%>'" id="btn_start" class="inputbutton" disabled>&nbsp;&nbsp;<input type="button" value="���}"  id="btn_close" onclick="window.close();" class="inputbutton" >
<BR><BR>
</td></tr>
<tr><td bgcolor=#555555 height=24 align=right><font Color="#FFFFFF">�w���������D�Ь�LDCC--�\���@ ����7403 </font></td></tr></table>

</body>
</html>

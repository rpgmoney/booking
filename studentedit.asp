<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 50 %>
<% Response.CacheControl = "No-cache" %>
<!-- #INCLUDE FILE="checkaccount.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/syconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<!-- #INCLUDE virtual="/include/lib/inc/lib.inc" -->
<!-- #INCLUDE virtual="/include/lib/asp/jsCalendar.asp" -->
<%
validate=trim(request("validate"))
sid=trim(request("sid"))
id=trim(request("id"))
name=trim(request("name"))
birthday=trim(request("birthday"))
sex=trim(request("sex"))
mail=trim(request("mail"))
cell=trim(request("cell"))
slevel=trim(request("slevel"))
grade=trim(request("grade"))
class1=trim(request("class1"))
department=trim(request("department"))
purpose=trim(request("purpose"))
purpose_remark=trim(request("purpose_remark"))
howknow=trim(request("howknow"))
howknow_remark=trim(request("howknow_remark"))
note=trim(request("note"))
btncontrol=trim(request("btncontrol"))
sytle_yn=trim(request("sytle_yn"))
strategy_yn=trim(request("strategy_yn"))
enable=trim(request("enable"))
classify=trim(request("classify"))
sender=ifnull(trim(request("sender")),"studentlist.asp")

if sid="" or isnull(sid) or isempty(sid) then
	sid=session("sid")
end if
if sid="S224955279" then sid="1096300068" end if


btnstatus=""'控制儲存鈕
set rs = server.CreateObject("adodb.recordset")
if validate="edit" then
	sql = "select * from boo_profile where sid='"&sid&"'"
	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
    if not  rs.EOF then
		if name<>"" then
            rs("name")=name
        end if
		if birthday<>"" then
            rs("birthday")=birthday
        end if
		if sex<>"" then
            rs("sex")=sex
        end if
		if mail<>"" then
            rs("mail")=mail
        end if
		if cell<>"" then
            rs("cell")=cell
        end if
		if slevel<>"" then
            rs("slevel")=slevel
        end if
		if department<>"" then
			 rs("department")=department
		end if
		if grade<>"" then
            rs("grade")=grade
        end if
		if class1<>"" then
            rs("class1")=class1
        else
			rs("class1")=null
		end if
		if purpose<>"" then
            rs("purpose")=purpose
        end if
		if purpose_remark<>"" then
            rs("purpose_remark")=purpose_remark
        else
			rs("purpose_remark")=null
		end if
		if howknow<>"" then
            rs("howknow")=howknow
        end if
		if howknow_remark<>"" then
            rs("howknow_remark")=howknow_remark
        else
			rs("howknow_remark")=null
		end if
		if note<>"" then
            rs("note")=note
        else
			 rs("note")=null
		end if
		if enable<>"" then
            rs("enable")=enable
        end if
		if classify<>"" then
            rs("classify")=classify
        end if
		
		
		'rs("initdate") = date()
		rs.Update
        if Err.Number=0 then 
           '註冊之後填寫問卷 
          ' response.redirect "qstyle.asp?sid=" & sid 
        else
            showmessage= Err.Description
        end if
	else
		showmessage="找不到該筆資料。"
	end if
elseif validate="delete" then
	sql = "delete from boo_profile where id='"&id&"'"
	msconn.Execute sql
	sql1 = "delete from boo_questionnaire_strategy where id='"&id&"'"
	msconn.Execute sql
	sql2 = "delete from boo_questionnaire_style where id='"&id&"'"
	msconn.Execute sql

	if Err.Number=0 then 
		response.redirect sender
	else
		 showmessage= Err.Description
	end if
else

	sql = "select * from boo_profile where sid='"&sid&"' and classify in ('S','E')"
'response.write sql
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
	if not rs.EOF then
		sid=trim(rs("sid"))
		id = trim(rs("id"))
		name=trim(rs("name"))
		birthday=trim(rs("birthday"))
		sex=trim(rs("sex"))
		mail=trim(rs("mail"))
		cell=trim(rs("cell"))
		slevel=trim(rs("slevel"))
		grade=trim(rs("grade"))
		class1=trim(rs("class1"))
		department=trim(rs("department"))
		purpose=trim(rs("purpose"))
		purpose_remark=trim(rs("purpose_remark"))
		howknow=trim(rs("howknow"))
		howknow_remark=trim(rs("howknow_remark"))
		note=trim(rs("note"))		
		sytle_yn=trim(rs("sytle_yn"))
		strategy_yn=trim(rs("strategy_yn"))
		enable=trim(rs("enable"))
		classify = trim(rs("classify"))
	else
		showmessage="你不是學員喔，此介面只提供學員編輯個人資料。" & "<br>"
		btnstatus = "disabled"
	end if
	rs.close

end if

if btnstatus="" then
	if sytle_yn<>"Y" then
		showmessage = showmessage & "尚未填寫學習風格問卷<br>"
	end if
	if strategy_yn<>"Y" then
		showmessage = showmessage & "尚未填寫學習策略問卷"
	end if
	if strategy_yn<>"Y" or strategy_yn<>"Y" then
		showmessage = showmessage & "，問卷填寫完成後才可預約。"
	end if
	if sytle_yn="Y"  and strategy_yn="Y" and session("questionnaire")<>"Y" then
		'可直接預約
		'response.redirect ""
		session("questionnaire")="Y"
		showmessage = "該人員已完成註冊，可以預約中心的療程。"
	end if
end if
'系所
StrDepartment="<option value=''> - 尚未指定 - </option>"
set rsLoad = server.CreateObject("adodb.recordset")
sql ="select * from s90_unit where unt_std='Y' order by unt_sort_seq  "
rsLoad.Open sql,syconn,adOpenStatic,adLockReadonly

while not rsLoad.EOF
	if department=rsLoad("unt_name_abr") then
		StrDepartment=StrDepartment&"<option value="""&rsLoad("unt_name_abr")&""" selected>"&  rsLoad("unt_name_abr")&"</option>"
	else
		StrDepartment=StrDepartment&"<option value="""&rsLoad("unt_name_abr")&""" >"&  rsLoad("unt_name_abr")&"</option>"
	end if 
	rsLoad.MoveNext 
wend
rsLoad.close

'學制
Strslevel="<option value=''> - 尚未指定 - </option>"
sql ="select * from s90_degree "
rsLoad.Open sql,syconn,adOpenStatic,adLockReadonly

while not rsLoad.EOF
	if slevel=rsLoad("dgr_name") then
		Strslevel=Strslevel&"<option value="""&rsLoad("dgr_name")&""" selected>"&  rsLoad("dgr_name")&"</option>"
	else
		Strslevel=Strslevel&"<option value="""&rsLoad("dgr_name")&""" >"&  rsLoad("dgr_name")&"</option>"
	end if 
	rsLoad.MoveNext 
wend
rsLoad.close

set rsLoad=nothing

%>
<HTML>
<HEAD>
<TITLE> 【LDCC英外語能力診斷輔導中心】 </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/stm31.js" type="text/javascript"></SCRIPT>
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>
<script language="javascript">

function check_input()
{
    var errmsg=""
	
	if (form1.cell.value=="")
        errmsg += "電話不能為空白\n";
    if (form1.birthday.value=="")
        errmsg += "生日不能為空白\n";
	
    if (form1.grade.value=="" )
        errmsg += "年級不能為空白\n";
	if (form1.mail.value=="")
        errmsg += "E-mail不能為空白\n";
	if (form1.purpose1.checked==true)
        errmsg += "來訪中心的目的是purpose of visiting this center必須選擇\n";
	if (form1.howknow1.checked==true)
        errmsg += "你如何得知本中心How do you know about this center必須選擇\n";
	
	if (errmsg == "")
	{
		var obj=document.getElementById("grade");
		if (obj.value!=""){
			objvalue=obj.value;
			if ( !isint(obj.value))
				errmsg += "年級必須為半形數字\n";  
		}
	}
	
    if (errmsg=="")
	{
        form1.validate.value = "edit";
		form1.submit();
	}
    else
    {
        alert(errmsg);
    }
}

function DeleteDc()
{
	var errmsg=""
	
	
	
	if (confirm("確定要刪除該筆資料嗎？")){
		form1.validate.value="delete";
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">學員個人基本資料 </TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="studentedit.asp" >
			<input type="hidden" value="" name="validate">
			<input type="hidden" id="sender" name="sender" value="<%=sender%>">
			<input type="hidden" id="btncontrol" name="btncontrol" value="<%=btncontrol%>">
			<input type="hidden" id="sytle_yn" name="sytle_yn" value="<%=sytle_yn%>">
			<input type="hidden" id="strategy_yn" name="strategy_yn" value="<%=strategy_yn%>">
			<input type="hidden" id="id" name="id" value="<%=id%>">
			
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>
						<%if sytle_yn<>"Y" then %>
						<input  type="button" <%=btnstatus%> value="填寫學習風格問卷" onclick="window.open('qstylerule.asp?sid=<%=sid%>','','fullscreen=1,scrollbars=1')" class="inputbutton" style='color:bc546f;font-size:9pt' onMouseOver="this.style.color='#ff6666'" onMouseOut="this.style.color='#bc546f'" >
						<%else%>
						<input  type="button"  <%=btnstatus%>  value="我的風格問卷" onclick="window.open('qstyle.asp?sid=<%=sid%>&validate=query','','fullscreen=1,scrollbars=1')" class="inputbutton" style='color:009900;font-size=9pt' onMouseOver="this.style.color='#00cc00'" onMouseOut="this.style.color='#009900'" >
						<%end if%>
						</TD>
						<TD>
						<%if strategy_yn<>"Y" then %>
						<input  type="button" <%=btnstatus%>  value="填寫學習策略問卷" onclick="window.open('qstrategyrule.asp?sid=<%=sid%>','','fullscreen=1,scrollbars=1')" class="inputbutton" style='color:bc546f;font-size:9pt' onMouseOver="this.style.color='#ff6666'" onMouseOut="this.style.color='#bc546f'" >
						<%else%>
						<input  type="button" <%=btnstatus%>  value="我的策略問卷" onclick="window.open('qstrategy.asp?sid=<%=sid%>&validate=query','','fullscreen=1,scrollbars=1')" class="inputbutton" style='color:009900;font-size=9pt' onMouseOver="this.style.color='#00cc00'" onMouseOut="this.style.color='#009900'" >
						<%end if%>
						</TD>
					</TR>
					
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>學號：</TD>
						<TD>姓名：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=sid%>" maxlength="25" size="35" name="sid" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						<TD>
						<input type="text" value="<%=name%>" maxlength="25" size="35" name="name" class="inputtext"  >
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>電話：</TD>
						<TD>生日：</TD>
						<TD></TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=cell%>" maxlength="30" size="30"  name="cell" class="inputtext" >
						</TD>
						<TD>
						<input type="text" value="<%=birthday%>" maxlength="25"  name="birthday" class="inputtext" readonly>
						<img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('birthday')" class="showhand">
						&nbsp;
						</TD>
						<TD>
							<TABLE cellSpacing=0 cellPadding=0  border=0 >
							<TR>
							<TD class="inputlabel">性別：</TD>
							<TD><input type="radio" name="sex" class="inputtext" value="M" <%if sex="M" then response.write "checked" end if%> ></TD><TD>男</TD>
							<TD><input type="radio" name="sex" class="inputtext" value="F" <%if sex="F" then response.write "checked" end if%>></TD><TD>女</TD>
							</TR>
							</TABLE>
						</TD>
						
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>學制：</TD>
						<TD>系所：</TD>
						<TD>年級：</TD>
						
						<TD>班級：</TD>
					</TR>
					<TR>
						<TD>
						<select name="slevel" class="inputtext">
						<%=Strslevel%>
						</select>
						</TD>
						<TD>
						<select name="department" class="inputtext">
						<%=StrDepartment%>
						</select>
						</TD>
						<TD>
						<input type="text" value="<%=grade%>" maxlength="10" size="10"  name="grade" class="inputtext" >
						</TD>
						
						<TD>
						<input type="text" value="<%=class1%>" maxlength="25"  name="class1" class="inputtext" >
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>Email：(請填寫學校Email，以防遺漏信件。)</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=mail%>" maxlength="100"  size="50" name="mail" class="inputtext" >
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR style="DISPLAY:<%if btncontrol="Y" then response.write "block" else response.write "none" end if%>"><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>開啟登入權限：</TD>
						<TD>權限：</TD>
					</TR>
					<TR>
						<TD>
						<select name="enable" class="inputtext">
						<option value="" <%if enable="all" then response.write "selected" end if%>> - 未指定 - </option>
						<option value="Y" <%if enable="Y" then response.write "selected" end if%>> - 是 - </option>
						<option value="N" <%if enable="N" then response.write "selected" end if%>> - 否 - </option>
						</select>
						</TD>
						<TD>
						<select name="classify" class="inputtext">
						<option value="S" <%if classify="S" then response.write "selected" end if%>>學員</option>
						<option value="E" <%if classify="E" then response.write "selected" end if%>>小老師</option>
						</select>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 width="98%" >
				<TR>
				<TD width="55%">
					<TABLE cellSpacing=0 cellPadding=0  border=0 >
						<TR class="inputlabel">
							<TD colspan="2">來訪中心的目的是purpose of visiting this center：</TD>
						</TR>
						<TR>
							<TD><input type="radio" name="purpose" id="purpose1" class="inputtext" checked value="" <%if purpose="" then response.write "checked" end if%> ></TD>
							<TD>未指定</TD>
						</TR>
						<TR>
							<TD><input type="radio" name="purpose" class="inputtext" value="dc" <%if purpose="dc" then response.write "checked" end if%> ></TD>
							<TD>診斷諮商(英文學習方法)Diagnosis and Consultation(dc)</TD>
						</TR>
						<TR>
							<TD><input type="radio" name="purpose" class="inputtext" value="software" <%if purpose="software" then response.write "checked" end if%> ></TD>
							<TD>使用英語自學軟體Englisg Learning Software(software)</TD>
						</TR>
						<TR>
							<TD><input type="radio" name="purpose" class="inputtext" value="op" <%if purpose="op" then response.write "checked" end if%> ></TD>
							<TD>口語練習Oral Practice(op)</TD>
						</TR>
						<TR>
							<TD><input type="radio" name="purpose" class="inputtext" value="workshops" <%if purpose="workshops" then response.write "checked" end if%> ></TD>
							<TD>英語學習方法講座Workshops</TD>
						</TR>
						<TR>
							<TD><input type="radio" name="purpose" class="inputtext" value="test-prep" <%if purpose="test-prep" then response.write "checked" end if%> ></TD>
							<TD>語測模擬測驗Simulation Tests(test-prep)</TD>
						</TR>
						<TR>
							<TD><input type="radio" name="purpose" class="inputtext" value="other" <%if purpose="other" then response.write "checked" end if%> ></TD>
							<TD>其他 Other &nbsp;<input type="text" value="<%=purpose_remark%>" maxlength="25" size="50" name="purpose_remark" class="inputtext" ></TD>
						</TR>
					</TABLE>
				</TD>
				<TD  rowspan="3" width="45%" height="100%">
				
				<!-- 歷年學習問卷 -->
					<TABLE cellSpacing=1 cellPadding=0  width="100%" height="100%" border=1 bgColor=#FFFFF4 align="center" bordercolor="#f4c60d">
					<TR height="10"><TD ><font color="#CC3300">歷年學習策略問卷</font>&nbsp;&nbsp;</TD>
					<TD ><font color="#CC3300">歷年學習風格問卷</font>&nbsp;&nbsp;</TD></TR>
						<TR valign="top">
						<TD >
						<%
							set rs1 = server.CreateObject("adodb.recordset")
							sql = "select * from boo_questionnaire_strategy where sid='"&sid&"' order by initdate desc"
							rs1.Open sql,msconn,adOpenStatic,adLockReadonly
							while not rs1.EOF
								response.write rs1("initdate") & "<input type='button' value='問卷分析'  onclick=""window.open('qstrategyreport.asp?qstid=" & rs1("qstid") &"','_blank','height=800, resizable=0, scrollbars=1, menubar=1, toolbar=1, top=10')"" class='inputbutton'><br>"
								rs1.MoveNext
							wend
							rs1.close
						%>&nbsp;
						</TD>
						<TD>
						<%
							sql = "select * from boo_questionnaire_style where sid='"&sid&"' order by initdate desc"
							rs1.Open sql,msconn,adOpenStatic,adLockReadonly
							while not rs1.EOF
								response.write rs1("initdate") & "<input type='button' value='問卷分析'  onclick=""window.open('qstylereport.asp?qsid=" & rs1("qsid") &"','_blank','height=800, resizable=0, scrollbars=1, menubar=1, toolbar=1, top=10')"" class='inputbutton'><br>"
								rs1.MoveNext
							wend
							rs1.close
						%>&nbsp;
					</TD></TR>
					</TABLE>
				<!-- 歷年學習問卷 -->
				</TD></TR>
				<TR>
				<TD>
					<TABLE cellSpacing=0 cellPadding=0  border=0 >
						<TR class="inputlabel">
							<TD colspan="2">你如何得知本中心How do you know about this center：</TD>
						</TR>
						<TR>
							<TD><input type="radio" name="howknow" id="howknow1" class="inputtext" checked value="" <%if howknow="" then response.write "checked" end if%> ></TD>
							<TD>未指定</TD>
						</TR>
						<TR>
							<TD><input type="radio" name="howknow" class="inputtext" value="brochures" <%if howknow="brochures" then response.write "checked" end if%> ></TD>
							<TD>小冊子 Brochures</TD>
						</TR>
						<TR>
							<TD><input type="radio" name="howknow" class="inputtext" value="Teachers or Classmates" <%if howknow="Teachers or Classmates" then response.write "checked" end if%> ></TD>
							<TD>老師或同學朋友告知 Teachers or Classmates</TD>
						</TR>
						<TR>
							<TD><input type="radio" name="howknow" class="inputtext" value="Website" <%if howknow="Website" then response.write "checked" end if%> ></TD>
							<TD>從網路上 Website</TD>
						</TR>
						
						<TR>
							<TD><input type="radio" name="howknow" class="inputtext" value="other" <%if howknow="other" then response.write "checked" end if%> ></TD>
							<TD>其他 Other <input type="text" value="<%=howknow_remark%>" maxlength="25"  name="howknow_remark" size="50" class="inputtext" ></TD>
						</TR>
					</TABLE>
				</TD>
				</TR>
				<TR>
				<TD>
					<TABLE cellSpacing=0 cellPadding=0  border=0 >
						<TR class="inputlabel">
							<TD>備註 Note：</TD>
						</TR>
						<TR>
							<TD>
							<textarea  cols="70" rows="6" name="note" class="inputtext" ><%=note%></textarea>
							</TD>
						</TR>
					</TABLE>
				</TD>
				</TR>
				</TABLE>
			</TD></TR>
			<TR>
			<TD>
			<BR>
			<input  type="button" value="確認後修改" <%=btnstatus%> onclick="check_input();" class="inputbutton" >
			<%if sytle_yn="Y" then %><input  type="button" value="重新填寫學習風格問卷" onclick="window.open('qstylerule.asp?sid=<%=sid%>','','fullscreen=1,scrollbars=1')" class="inputbutton"><%end if%>
			<%if strategy_yn="Y" then %><input  type="button" value="重新填寫學習策略問卷"  onclick="window.open('qstrategyrule.asp?sid=<%=sid%>','','fullscreen=1,scrollbars=1')"  class="inputbutton"><%end if%>
			<%if btncontrol="Y" then%>
			<input  type="button" value="刪除" onclick="DeleteDc();" class="inputbutton" <%if session("classify")<>"A" then response.write "disabled" end if %> >
			<input  type="button" value="返回" class="inputbutton" onclick="window.location='<%=sender%>'">
			<%end if%>
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
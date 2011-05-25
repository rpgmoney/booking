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
strength=trim(request("strength"))
needed=trim(request("needed"))
effect=trim(request("effect"))


content=trim(request("content"))
content_remark=trim(request("content_remark"))

feedback=trim(request("feedback"))
teacher=trim(request("teacher"))
backcase=trim(request("backcase"))



today=year(date())-1911 & right("0" & month(date()),2) & right("0" & day(date()),2)

sender=ifnull(trim(request("sender")),"consult.asp" )

set dic = server.CreateObject("scripting.dictionary")
dic.Add "1","日"
dic.Add "2","一"
dic.Add "3","二"
dic.Add "4","三"
dic.Add "5","四"
dic.Add "6","五"
dic.Add "7","六"


set rs = server.CreateObject("adodb.recordset")
set rs2 = server.CreateObject("adodb.recordset")

if validate="add" then
		sql = "select * from boo_consult where tid='"&id&"' "
		rs.Open sql,msconn,adOpenStatic,adLockOptimistic
		if rs.EOF then
			rs.AddNew
			if id<>"" then
				rs("tid")=id
			end if
			if content<>"" then
				rs("content")=content
			end if
			if content_remark<>"" then
				rs("content_remark")=content_remark
			end if
			
			if feedback<>"" then
				rs("feedback")=feedback
			end if
			
			if teacher<>"" then
				rs("teacher")=teacher
			end if
			if backcase<>"" then
				rs("backcase")=backcase
			end if
			if strength<>"" then
				rs("strength")=strength
			end if
			if needed<>"" then
				rs("needed")=needed
			end if
			if effect<>"" then
				rs("effect")=effect
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
			showmessage="資料重覆。"
		end if

		rs.close	

elseif validate="edit" then
	sql = "select * from boo_consult  where tid='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
	if  not rs.EOF then
		if content<>"" then
			rs("content")=content
		end if
		if content_remark<>"" then
			rs("content_remark")=content_remark
		else
			rs("content_remark")=null
		end if
		
		if feedback<>"" then
			rs("feedback")=feedback
		else
			rs("feedback")=null
		end if
		if backcase<>"" then
			rs("backcase")=backcase
		else
			rs("backcase")=null
		end if
		if strength<>"" then
			rs("strength")=strength
		else
			rs("strength")=null
		end if
		if needed<>"" then
			rs("needed")=needed
		else
			rs("needed")=null
		end if
		if effect<>"" then
			rs("effect")=effect
		else
			rs("effect")=null
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
		showmessage="找不到該筆資料。"
	end if

	rs.close	
else
	'預約資訊
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
		consult=trim(rs("consult"))
		yn=trim(rs("yn"))
		category=trim(rs("category"))
		languagecode=trim(rs("languagecode"))
		signin=trim(rs("signin"))

	end if
	rs.close
	'諮商內容資訊
	sql = "select * from boo_consult   where tid='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
    if rs.EOF then
		validate="add"
	else
		validate="edit"
		did=trim(rs("id"))
		content=trim(rs("content"))
		content_remark=trim(rs("content_remark"))
		
		feedback=trim(rs("feedback"))
		teacher=trim(rs("teacher"))
		backcase=trim(rs("backcase"))
		strength=trim(rs("strength"))
		needed=trim(rs("needed"))
		effect=trim(rs("effect"))

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
<TITLE> 【LDCC英外語能力診斷輔導中心】 </TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK rel=stylesheet Type="text/css" href="lib\default.css">
<SCRIPT language="JavaScript1.2" src="/include/lib/js/stm31.js" type="text/javascript"></SCRIPT>
<SCRIPT language="JavaScript1.2" src="/include/lib/js/lib.js" type="text/javascript"></SCRIPT>
<script language="javascript">

function check_input()
{
    var errmsg=""
	var tmpobj1=""
	var icounter=0
	for ( var i=1;i<=10;i++){
			tmpobj1 = document.getElementById("content"+i);
			if (tmpobj1.checked== true){icounter++;}
	}
	if (icounter==0){errmsg += "請選擇診斷內容\n";}
	if (form1.strength.value=="")
		errmsg += "優點/缺點不能為空白\n";
	if (form1.needed.value=="")
		errmsg += "須改善之處不能為空白\n";
	if (form1.effect.value=="")
		errmsg += "預期成效不能為空白\n";
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">學員諮商紀錄維護編輯</TD>
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
						<TD class="inputlabel">學號：</TD><TD><%=sid%></TD>
						<TD class="inputlabel">姓名：</TD><TD><%=name%></TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR>
						<TD class="inputlabel">學制：</TD><TD><%=slevel%></TD>
						<TD class="inputlabel">系所：</TD><TD><%=department%></TD>
						<TD class="inputlabel">年級：</TD><TD><%=grade%></TD>
						<TD class="inputlabel">班級：</TD><TD><%=class1%></TD>
						<TD class="inputlabel">大專英檢成績：</TD><TD><%=score%></TD>
						<TD></TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR>
						<TD class="inputlabel"><%if category = "T" then response.write "老師" else response.write "小老師" end if%>：</TD><TD><%=teachername%></TD>
						<TD class="inputlabel">&nbsp;語言別：&nbsp;</TD><TD><%=languagecode%></TD>
					</TR>
				</TABLE>
			</TD></TR>
			
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR>
						<TD  class="inputlabel">預約項目：&nbsp;&nbsp;</TD><TD><%=item%></TD>
						<TD  class="inputlabel">預約日期：</TD><TD><%=bdate%></TD>
						<TD  class="inputlabel">星期：</TD><TD><%="（&nbsp;"&dic.Item(cstr(cint(weekday(dateformat(bdate)))))&"&nbsp;）"%></TD>
						<TD  class="inputlabel">預約時段：</TD><TD><%=btime%></TD>
						<TD><%=replace(replace(replace(timeflag,"U","上一節(25分)"),"B","下一節(25分)"),"A","上下二節(50分)")%></TD>
					</TR>
					
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR >
						<TD class="inputlabel">諮商主題：</TD><TD><%=consult%></TD>
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
			<form id="form1" name="form1" method="post" action="consultedit.asp"   onsubmit="return check_input();">
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
						<TD>諮商內容：</TD>
					</TR>
					<TR>
						<TD>
							<TABLE cellSpacing=1 cellPadding=1  border=0 >
							<TR>
							<TD><input type="checkbox" name="content" id="content1"  class="inputtext" value="Listening" <%if Instr(content,"Listening")<>0  then response.write "checked" end if%> ></TD><TD >Listening</TD>
							<TD><input type="checkbox" name="content" id="content2"  class="inputtext" value="Speaking" <%if Instr(content,"Speaking")<> 0  then response.write "checked" end if%>></TD><TD >Speaking</TD>
							<TD><input type="checkbox" name="content" id="content3"  class="inputtext" value="Reading" <%if Instr(content,"Reading") <> 0 then response.write "checked" end if%>></TD><TD >Reading</TD>
							<TD><input type="checkbox" name="content" id="content4"  class="inputtext" value="Writing" <%if  Instr(content,"Writing") <> 0  then response.write "checked" end if%>></TD><TD >Writing</TD>
							<TD><input type="checkbox" name="content" id="content5"  class="inputtext" value="Grammar" <%if Instr(content,"Grammar") <>0  then response.write "checked" end if%>></TD><TD >Grammar</TD>
							<TD><input type="checkbox" name="content" id="content6"  class="inputtext" value="Pronunciation" <%if  Instr(content,"Pronunciation") <> 0  then response.write "checked" end if%>></TD><TD >Pronunciation</TD>
							</TR>
							<TR>
							<TD><input type="checkbox" name="content" id="content7"  class="inputtext" value="Test-taking" <%if Instr(content,"Test-taking") <> 0  then response.write "checked" end if%>></TD><TD >Test-taking</TD>
							<TD><input type="checkbox" name="content" id="content8"  class="inputtext" value="Public Speaking" <%if  Instr(content,"Public Speaking") <> 0  then response.write "checked" end if%>></TD><TD >Public Speaking</TD>
							<TD><input type="checkbox" name="content" id="content9"  class="inputtext" value="Presentation Skills" <%if  Instr(content,"Presentation Skills") <> 0 then response.write "checked" end if%>></TD><TD >Presentation Skills</TD>
							<TD><input type="checkbox" name="content" id="content10"  class="inputtext" value="Other" <%if  Instr(content,"Other") <> 0 then response.write "checked" end if%>></TD><TD >Other</TD>
							<TD colspan="4"><input type="text" value="<%=content_remark%>"  size="25" maxlength="50" class="inputtext"  ></TD>
							</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>優點/缺點Strength(s)/Weakness(es)：</TD>
					</TR>
					<TR>
						<TD>
							<TABLE cellSpacing=1 cellPadding=1  border=0 >
							<TR>
							<TD><textarea name="strength" rows="5" cols="100" class="inputtext"  ><%=strength%></textarea></TD>
							</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>須改善之處Tmprovement(s) Needed：</TD>
					</TR>
					<TR>
						<TD>
							<TABLE cellSpacing=1 cellPadding=1  border=0 >
							<TR>
							<TD><textarea name="needed" rows="5" cols="100" class="inputtext"  ><%=needed%></textarea></TD>
							</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>預期成效Anticipated Effects：</TD>
					</TR>
					<TR>
						<TD>
							<TABLE cellSpacing=1 cellPadding=1  border=0 >
							<TR>
							<TD><textarea name="effect" rows="5" cols="100" class="inputtext"  ><%=effect%></textarea></TD>
							</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>備註(remark)：</TD>
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
						<TD>駐診老師：</TD>
						<TD>回診個案：</TD>
					</TR>
					<TR>
						<TD><input type="text" value="<%=teacher%>" maxlength="25" size="35" name="teacher" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly ></TD>
						<TD align="center"><input type="checkbox" name="backcase" id="backcase"  class="inputtext" value="Y" <%if backcase="Y" then response.write "checked" end if%>></TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD><BR><input  type="submit" value="儲存" class="inputbutton" ><input  type="button" value="返回" class="inputbutton" onclick="window.location='<%=sender%>'"></TD></TR>
			
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


<%@Language=VBScript LCID=1033%>
<% Session.TimeOut = 90 %>
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
op=trim(request("op"))
optime=trim(request("optime"))
optime_b=ifnull(trim(request("optime_b")),0)
optime_c=ifnull(trim(request("optime_c")),0)
pp=trim(request("pp"))
pptime=trim(request("pptime"))
pptime_b=ifnull(trim(request("pptime_b")),0)
pptime_c=ifnull(trim(request("pptime_c")),0)
crkp=trim(request("crkp"))
crkptime=trim(request("crkptime"))
crkptime_b=ifnull(trim(request("crkptime_b")),0)
crkptime_c=ifnull(trim(request("crkptime_c")),0)
write=trim(request("write"))
writetime=trim(request("writetime"))
writetime_b=ifnull(trim(request("writetime_b")),0)
writetime_c=ifnull(trim(request("writetime_c")),0)
reading=trim(request("reading"))
readtime=trim(request("readtime"))
readtime_b=ifnull(trim(request("readtime_b")),0)
readtime_c=ifnull(trim(request("readtime_c")),0)
note=trim(request("note"))
backdate=trim(request("backdate"))
teacher=trim(request("teacher"))
backcase=trim(request("backcase"))


today=year(date())-1911 & right("0" & month(date()),2) & right("0" & day(date()),2)

sender=ifnull(trim(request("sender")),"diagnosis.asp" )

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
		sql = "select * from boo_diagnosis where tid='"&id&"' "
		'response.write sql
		
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
			if optime<>"" then
				rs("optime")=optime
				rs("op")="Y"
			end if
			if pptime<>"" then
				rs("pptime")=pptime
				rs("pp")="Y"
			end if
			if crkptime<>"" then
				rs("crkptime")=crkptime
				rs("crkp")="Y"
			end if
			if writetime<>"" then
				rs("writetime")=writetime
				rs("write")="Y"
			end if
			if readtime<>"" then
				rs("readtime")=readtime
				rs("reading")="Y"
			end if
			if note<>"" then
				rs("note")=note
			end if
			if backdate<>"" then
				rs("backdate")=backdate
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
				
				'軟體
				swcounter=trim(request("swcounter"))'共幾筆
				sqlinsert=""
				for k=1 to swcounter
					softwareid=trim(request("softwareid" & k ))
					softwaretime=trim(request("softwaretime" & k ))

					if  softwaretime<>"" then
						sqlinsert = sqlinsert & "insert into boo_diagnosis_softwore  (tid,category,sid,times,modifyuid,modifydate,inituid,initdate) values('"&id&"','S','"&softwareid&"','"&softwaretime&"','"&session("sid")&"','"&date()&"','"&session("sid")&"','"&date()&"');"
					end if
				Next
				msconn.execute sqlinsert
				'response.write "sqlinsert=" & sqlinsert & "<br>"
				'補充教材
				excounter=trim(request("excounter"))'共幾筆
				sqlinsert=""
				for k=1 to excounter
					extraid=trim(request("extraid" & k ))
					extratime=trim(request("extratime" & k ))
					if  extratime<>"" then
						sqlinsert = sqlinsert & "insert into boo_diagnosis_softwore  (tid,category,sid,times,modifyuid,modifydate,inituid,initdate) values('"&id&"','T','"&extraid&"','"&extratime&"','"&session("sid")&"','"&date()&"','"&session("sid")&"','"&date()&"');"
					end if
				Next
				'response.write "sqlinsert=" & sqlinsert & "<br>"
				if sqlinsert<>"" then
					msconn.begintrans
					msconn.execute sqlinsert
					if err.number=0 then
						msconn.committrans
						response.redirect sender
					else
						msconn.rollbacktrans
					end if
				end if
				
			else
				showmessage= Err.Description
			end if

		else
			showmessage="資料重覆。"
		end if
		'response.end
		rs.close	

elseif validate="edit" then
	sql = "select * from boo_diagnosis where tid='"&id&"' "
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
		if optime<>"" then
			rs("optime")=optime
			rs("op")="Y"
		else
			rs("optime")=null
			rs("op")=null
		end if
		if pptime<>"" then
			rs("pptime")=pptime
			rs("pp")="Y"
		else
			rs("pptime")=null
			rs("pp")=null
		end if
		if crkptime<>"" then
			rs("crkptime")=crkptime
			rs("crkp")="Y"
		else
			rs("crkptime")=null
			rs("crkp")=null
		end if
		if writetime<>"" then
			rs("writetime")=writetime
			rs("write")="Y"
		else
			rs("writetime")=null
			rs("write")=null
		end if
		if readtime<>"" then
			rs("readtime")=readtime
			rs("reading")="Y"
		else
			rs("readtime")=null
			rs("reading")=null
		end if
		
		if note<>"" then
			rs("note")=note
		else
			rs("note")=null
		end if
		if backdate<>"" then
			rs("backdate")=backdate
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
			'軟體
			set rsq = server.CreateObject("adodb.recordset")
			swcounter=trim(request("swcounter"))'共幾筆
			sqlinsert=""
			for k=1 to swcounter
				softwareid=trim(request("softwareid" & k ))
				softwaretime=trim(request("softwaretime" & k ))

				if  softwaretime<>"" then
					sqlq ="select * from boo_diagnosis_softwore where tid='"&id&"' and sid='"&softwareid&"'"
					rsq.Open sqlq,msconn,adOpenStatic,adLockReadonly
					if rsq.EOF then
						sqlinsert = sqlinsert & "insert into boo_diagnosis_softwore  (tid,category,sid,times,modifyuid,modifydate,inituid,initdate) values('"&id&"','S','"&softwareid&"','"&softwaretime&"','"&session("sid")&"','"&date()&"','"&session("sid")&"','"&date()&"');"
					else
						sqlinsert = sqlinsert & "update  boo_diagnosis_softwore  set times='"&softwaretime&"' where tid='"&id&"' and sid='"&softwareid&"'; "
					end if
					rsq.close
				else
					sqlinsert = sqlinsert & "update  boo_diagnosis_softwore  set times=0 where tid='"&id&"' and sid='"&softwareid&"'; "
				end if
			Next
			msconn.execute sqlinsert
			'response.write "sqlinsert=" & sqlinsert & "<br>"
			'補充教材
			excounter=trim(request("excounter"))'共幾筆
			sqlinsert=""
			for k=1 to excounter
				extraid=trim(request("extraid" & k ))
				extratime=trim(request("extratime" & k ))
				if  extratime<>"" then
					sqlq ="select * from boo_diagnosis_softwore where tid='"&id&"' and sid='"&extraid&"'"
					rsq.Open sqlq,msconn,adOpenStatic,adLockReadonly
					if rsq.EOF then
						sqlinsert = sqlinsert & "insert into boo_diagnosis_softwore  (tid,category,sid,times,modifyuid,modifydate,inituid,initdate) values('"&id&"','T','"&extraid&"','"&extratime&"','"&session("sid")&"','"&date()&"','"&session("sid")&"','"&date()&"');"
					else
						sqlinsert = sqlinsert & "update  boo_diagnosis_softwore  set times='"&extratime&"' where tid='"&id&"' and sid='"&extraid&"'; "
					end if
					rsq.close
				else
					sqlinsert = sqlinsert & "update  boo_diagnosis_softwore  set times=0 where tid='"&id&"' and sid='"&extraid&"'; "
				end if
			Next
			'response.write "sqlinsert=" & sqlinsert & "<br>"
			sqlinsert = sqlinsert & "delete from  boo_diagnosis_softwore  where times=0 and times_b=0 and times_c=0  and  tid='"&id&"'"
			if sqlinsert<>"" then
					msconn.begintrans
					msconn.execute sqlinsert
					if err.number=0 then
						msconn.committrans
						response.redirect sender
					else
						msconn.rollbacktrans
					end if
			end if
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
		yn=trim(rs("yn"))
		category=trim(rs("category"))
		languagecode=trim(rs("languagecode"))
		signin=trim(rs("signin"))
		


	end if
	rs.close
	'處方資訊
	sql = "select * from boo_diagnosis   where tid='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
    if rs.EOF then
		validate="add"
	else
		validate="edit"
		did=trim(rs("id"))
		content=trim(rs("content"))
		content_remark=trim(rs("content_remark"))
		op=trim(rs("op"))
		optime=trim(rs("optime"))
		optime_b=ifnull(trim(rs("optime_b")),0)
		optime_c=ifnull(trim(rs("optime_c")),0)
		pp=trim(rs("pp"))
		pptime=trim(rs("pptime"))
		pptime_b=ifnull(trim(rs("pptime_b")),0)
		pptime_c=ifnull(trim(rs("pptime_c")),0)
		crkp=trim(rs("crkp"))
		crkptime=trim(rs("crkptime"))
		crkptime_b=ifnull(trim(rs("crkptime_b")),0)
		crkptime_c=ifnull(trim(rs("crkptime_c")),0)
		write=trim(rs("write"))
		writetime=trim(rs("writetime"))
		writetime_b=ifnull(trim(rs("writetime_b")),0)
		writetime_c=ifnull(trim(rs("writetime_c")),0)
		reading=trim(rs("reading"))
		readtime=trim(rs("readtime"))
		readtime_b=ifnull(trim(rs("readtime_b")),0)
		readtime_c=ifnull(trim(rs("readtime_c")),0)
		note=trim(rs("note"))
		backdate=trim(rs("backdate"))
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


	if (form1.backdate.value=="")
		errmsg += "回診日期不能為空白\n";
	if (form1.strength.value=="")
		errmsg += "優點/缺點不能為空白\n";
	if (form1.needed.value=="")
		errmsg += "須改善之處不能為空白\n";
	if (form1.effect.value=="")
		errmsg += "預期成效不能為空白\n";
	
    if (errmsg=="")
	{
        
		
		return true;

	}
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">學員診斷處方籤紀錄編輯</TD>
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
						<TD  class="inputlabel">星期：</TD><TD><%="（&nbsp;"&dic.Item(cstr(cint(weekday(NumberToDateFormat(bdate)))))&"&nbsp;）"%></TD>
						<TD  class="inputlabel">預約時段：</TD><TD><%=btime%></TD>
						<TD><%=replace(replace(replace(timeflag,"U","上一節(25分)"),"B","下一節(25分)"),"A","上下二節(50分)")%></TD>
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
			<form id="form1" name="form1" method="post" action="diagnosisedit.asp"   onsubmit="return check_input();">
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
						<TD>診斷內容：</TD>
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
						<TD>須改善之處 Improvement(s) Needed：</TD>
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
						<TD>建議（Recommendation）：</TD>
					</TR>
					<TR>
						<TD>
							<TABLE cellSpacing=1 cellPadding=2  border=0 >
							<TR><TD  class="inputlabel">口語：</TD>
							<TD>
							<select name="optime"  class="inputtext"  >
							<option value="">未指定</option>
							<%=NumOption2(1,10,optime)%>
							</select>
							</TD>
							<TD>次(一次為上下二節50分)</TD>
							<TD  class="inputlabel">已預約：</TD>
							<TD><input type="text" value="<%=optime_b%>" maxlength="25" size="5" name="optime_b" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly ></TD>
							<TD  class="inputlabel">已完成：</TD>
							<TD><input type="text" value="<%=optime_c%>" maxlength="25" size="5" name="optime_c" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly ></TD>
							
							</TR>
							<TR>
							<TD class="inputlabel">簡報：</TD>
							<TD>
							<select name="pptime"  class="inputtext"  >
							<option value="">未指定</option>
							<%=NumOption2(1,10,pptime)%>
							</select>
							</TD>
							<TD>次(一次為上下二節50分)</TD>
							<TD  class="inputlabel">已預約：</TD>
							<TD><input type="text" value="<%=pptime_b%>" maxlength="25" size="5" name="pptime_b" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly ></TD>
							<TD  class="inputlabel">已完成：</TD>
							<TD><input type="text" value="<%=pptime_c%>" maxlength="25" size="5" name="pptime_c" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly ></TD>
							
							</TR>
							<TR><TD class="inputlabel">詩歌：</TD>
							<TD>
							<select name="crkptime"  class="inputtext"  >
							<option value="">未指定</option>
							<%=NumOption2(1,10,crkptime)%>
							</select>
							</TD>
							<TD>次(一次為上下二節50分)</TD>
							<TD  class="inputlabel">已預約：</TD>
							<TD><input type="text" value="<%=crkptime_b%>" maxlength="25" size="5" name="crkptime_b" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly ></TD>
							<TD  class="inputlabel">已完成：</TD>
							<TD><input type="text" value="<%=crkptime_c%>" maxlength="25" size="5" name="crkptime_c" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly ></TD>
							
							</TR>
							<TR>
							<TD class="inputlabel">寫作諮商：</TD>
							<TD>
							<select name="writetime"  class="inputtext"  >
							<option value="">未指定</option>
							<%=NumOption2(1,10,writetime)%>
							</select>
							</TD>
							<TD>次(一次為上下二節50分)</TD>
							<TD  class="inputlabel">已預約：</TD>
							<TD><input type="text" value="<%=writetime_b%>" maxlength="25" size="5" name="writetime_b" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly ></TD>
							<TD  class="inputlabel">已完成：</TD>
							<TD><input type="text" value="<%=writetime_c%>" maxlength="25" size="5" name="writetime_c" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly ></TD>
							
							</TR>
							<TR><TD class="inputlabel">閱讀諮商：</TD>
							<TD>
							<select name="readtime"  class="inputtext"  >
							<option value="">未指定</option>
							<%=NumOption2(1,10,readtime)%>
							</select>
							</TD>
							<TD>次(一次為上下二節50分)</TD>
							<TD  class="inputlabel">已預約：</TD>
							<TD><input type="text" value="<%=readtime_b%>" maxlength="25" size="5" name="readtime_b" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly ></TD>
							<TD  class="inputlabel">已完成：</TD>
							<TD><input type="text" value="<%=readtime_c%>" maxlength="25" size="5" name="readtime_c" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly ></TD>
							</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>

			</TD></TR>
			
		<tr class="inputlabel"><td>自學軟體：</td></tr>
		<TR><TD>
			<TABLE cellSpacing=1 cellPadding=2 width="90%"  border=0    >
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<%
					'軟體
					set rsLoad = server.CreateObject("adodb.recordset")
					sql = "select a.*,b.times,b.times_b,b.times_c  from "
					sql = sql & " ( "
					sql = sql & " select * from boo_software where yn='Y' and category='S' "
					sql = sql & " ) a left join "
					sql = sql & " ( "
					sql = sql & " select  sid,times,times_b,times_c  from boo_diagnosis_softwore  where  tid='"&id&"' and  category='S' "
					sql = sql & " ) b on a.id=b.sid  order by floor,software"

					'response.write sql
					rsLoad.Open sql,msconn,adOpenStatic,adLockReadonly
					Str=""
					i = 0
					while not rsLoad.EOF
						i= i +1
					%>
						<TR><TD></TD><TD>&nbsp;<% =rsLoad("floor")%>&nbsp;-&nbsp;<% =rsLoad("software")%></TD>
						<TD>
							<input type="hidden" name="softwareid<%=i%>" value="<%=rsLoad("id")%>">
							<select name="softwaretime<%=i%>"  class="inputtext"  >
							<option value="">未指定</option>
							<option value="60" <%if   rsLoad("times")="60" then  response.write "selected" end if%>>60(1小時)</option>
							<option value="120" <%if   rsLoad("times")="120" then  response.write "selected" end if%>>120(2小時)</option>
							<option value="180" <%if   rsLoad("times")="180" then  response.write "selected" end if%>>180(3小時)</option>
							<option value="240" <%if   rsLoad("times")="240" then  response.write "selected" end if%>>240(4小時)</option>
							<option value="300" <%if   rsLoad("times")="300" then  response.write "selected" end if%>>300(5小時)</option>
							<option value="360" <%if   rsLoad("times")="360" then  response.write "selected" end if%>>360(6小時)</option>
							<option value="420" <%if   rsLoad("times")="420" then  response.write "selected" end if%>>420(7小時)</option>
							<option value="480" <%if   rsLoad("times")="480" then  response.write "selected" end if%>>480(8小時)</option>
							<option value="540" <%if   rsLoad("times")="540" then  response.write "selected" end if%>>540(9小時)</option>
							<option value="600" <%if   rsLoad("times")="600" then  response.write "selected" end if%>>600(10小時)</option>

							<option value="660" <%if   rsLoad("times")="660" then  response.write "selected" end if%>>660(11小時)</option>
							<option value="720" <%if   rsLoad("times")="720" then  response.write "selected" end if%>>720(12小時)</option>
							<option value="780" <%if   rsLoad("times")="780" then  response.write "selected" end if%>>780(13小時)</option>
							<option value="840" <%if   rsLoad("times")="840" then  response.write "selected" end if%>>840(14小時)</option>
							<option value="900" <%if   rsLoad("times")="900" then  response.write "selected" end if%>>900(15小時)</option>
							<option value="960" <%if   rsLoad("times")="960" then  response.write "selected" end if%>>960(16小時)</option>
							<option value="1020" <%if   rsLoad("times")="1020" then  response.write "selected" end if%>>1020(17小時)</option>
							<option value="1080" <%if   rsLoad("times")="1080" then  response.write "selected" end if%>>1080(18小時)</option>
							<option value="1140" <%if   rsLoad("times")="1140" then  response.write "selected" end if%>>1140(19小時)</option>
							<option value="1200" <%if   rsLoad("times")="1200" then  response.write "selected" end if%>>1200(20小時)</option>

							</select>
							</TD>
							<TD>分鐘&nbsp;&nbsp;</TD>
							<TD  class="inputlabel">已預約：</TD>
							<TD><input type="text" value="<%=rsLoad("times_b")%>" maxlength="25" size="5" name="softtime_b<%=i%>" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly ></TD>
							<TD  class="inputlabel">已完成：</TD>
							<TD><input type="text" value="<%=rsLoad("times_c")%>" maxlength="25" size="5" name="softtime_c<%=i%>" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly ></TD>
						
						</TR>
					<%
						rsLoad.MoveNext 
					wend
					rsLoad.close
					
					response.write Str
					%>
					<input type="hidden" value="<%=i%>" name="swcounter">
				</TABLE>
			</TD></TR>
			</TABLE>
		</TD></TR>
		<tr class="inputlabel"><td>補充教材：</td></tr>
		<TR><TD>
			<TABLE cellSpacing=1 cellPadding=2 width="90%"  border=0    >
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<%
					'補充教材
					sql = "select a.*,b.times,b.times_b,b.times_c  from "
					sql = sql & " ( "
					sql = sql & " select * from boo_software where yn='Y' and category='T' "
					sql = sql & " ) a left join "
					sql = sql & " ( "
					sql = sql & " select  sid,times,times_b,times_c  from boo_diagnosis_softwore  where  tid='"&id&"' and  category='T' "
					sql = sql & " ) b on a.id=b.sid "
					rsLoad.Open sql,msconn,adOpenStatic,adLockReadonly
					Str=""
					j = 0 
					while not rsLoad.EOF
					j= j +1
					%>
						<TR><TD></TD><TD>&nbsp;<% =rsLoad("software")%></TD>
						<TD>
							<input type="hidden" name="extraid<%=j%>" value="<%=rsLoad("id")%>">
							<select name="extratime<%=j%>"  class="inputtext"  >
							<option value="">未指定</option>
							<option value="60" <%if   rsLoad("times")="60" then  response.write "selected" end if%>>60(1小時)</option>
							<option value="120" <%if   rsLoad("times")="120" then  response.write "selected" end if%>>120(2小時)</option>
							<option value="180" <%if   rsLoad("times")="180" then  response.write "selected" end if%>>180(3小時)</option>
							<option value="240" <%if   rsLoad("times")="240" then  response.write "selected" end if%>>240(4小時)</option>
							<option value="300" <%if   rsLoad("times")="300" then  response.write "selected" end if%>>300(5小時)</option>
							<option value="360" <%if   rsLoad("times")="360" then  response.write "selected" end if%>>360(6小時)</option>
							<option value="420" <%if   rsLoad("times")="420" then  response.write "selected" end if%>>420(7小時)</option>
							<option value="480" <%if   rsLoad("times")="480" then  response.write "selected" end if%>>480(8小時)</option>
							<option value="540" <%if   rsLoad("times")="540" then  response.write "selected" end if%>>540(9小時)</option>
							<option value="600" <%if   rsLoad("times")="600" then  response.write "selected" end if%>>600(10小時)</option>

							<option value="660" <%if   rsLoad("times")="660" then  response.write "selected" end if%>>660(11小時)</option>
							<option value="720" <%if   rsLoad("times")="720" then  response.write "selected" end if%>>720(12小時)</option>
							<option value="780" <%if   rsLoad("times")="780" then  response.write "selected" end if%>>780(13小時)</option>
							<option value="840" <%if   rsLoad("times")="840" then  response.write "selected" end if%>>840(14小時)</option>
							<option value="900" <%if   rsLoad("times")="900" then  response.write "selected" end if%>>900(15小時)</option>
							<option value="960" <%if   rsLoad("times")="960" then  response.write "selected" end if%>>960(16小時)</option>
							<option value="1020" <%if   rsLoad("times")="1020" then  response.write "selected" end if%>>1020(17小時)</option>
							<option value="1080" <%if   rsLoad("times")="1080" then  response.write "selected" end if%>>1080(18小時)</option>
							<option value="1140" <%if   rsLoad("times")="1140" then  response.write "selected" end if%>>1140(19小時)</option>
							<option value="1200" <%if   rsLoad("times")="1200" then  response.write "selected" end if%>>1200(20小時)</option>
							</select>
							</TD>
							<TD>分鐘&nbsp;&nbsp;</TD>
							<TD  class="inputlabel">已預約：</TD>
							<TD><input type="text" value="<%=rsLoad("times_b")%>" maxlength="25" size="5" name="extratime_b<%=j%>" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly ></TD>
							<TD  class="inputlabel">已完成：</TD>
							<TD><input type="text" value="<%=rsLoad("times_c")%>" maxlength="25" size="5" name="extratime_c<%=j%>" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly ></TD>
						
						</TR>
					<%
						rsLoad.MoveNext 
					wend
					rsLoad.close
					
					set rsLoad=nothing
					response.write Str
					%>
					<input type="hidden" value="<%=j%>" name="excounter">
				</TABLE>
			</TD></TR>
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
						<TD>備註說明：</TD>
					</TR>
					<TR>
						<TD>
							<TABLE cellSpacing=1 cellPadding=1  border=0 >
							<TR>
							<TD><textarea name="note" rows="5" cols="100" class="inputtext"  ><%=note%></textarea></TD>
							</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR class="inputlabel">
						<TD>回診日期：</TD>
						<TD>駐診老師：</TD>
						<TD>回診個案：</TD>
					</TR>
					<TR>
						<TD><input type="text" id="backdate" name="backdate" value="<%=backdate%>"  maxlength="6" class="inputtext" readonly><img src="/include/lib/images/calendar.gif" onClick="jsCalendar_PopWinEng('backdate')" class="showhand"></TD>
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


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
related=trim(request("related"))
idea=trim(request("idea"))
comment=trim(request("comment"))
others=trim(request("others"))

pronunciation=trim(request("pronunciation"))
fluency=trim(request("fluency"))
vocabulary=trim(request("vocabulary"))
grammar=trim(request("grammar"))
overall=trim(request("overall"))

teacher=trim(request("teacher"))




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
		sql = "select * from boo_op where tid='"&id&"' "
		rs.Open sql,msconn,adOpenStatic,adLockOptimistic
		if rs.EOF then
			rs.AddNew
			if id<>"" then
				rs("tid")=id
			end if
			if pronunciation<>"" then
				rs("pronunciation")=pronunciation
			end if
			if fluency<>"" then
				rs("fluency")=fluency
			end if
			
			if vocabulary<>"" then
				rs("vocabulary")=vocabulary
			end if
			if grammar<>"" then
				rs("grammar")=grammar
			end if
			if overall<>"" then
				rs("overall")=overall
			end if

			if teacher<>"" then
				rs("teacher")=teacher
			end if
			if related<>"" then
				rs("related")=related
			end if
			if idea<>"" then
				rs("idea")=idea
			end if
			if comment<>"" then
				rs("comment")=comment
			end if
			if others<>"" then
				rs("others")=others
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
	sql = "select * from boo_op  where tid='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockOptimistic
	if  not rs.EOF then
		if pronunciation<>"" then
			rs("pronunciation")=pronunciation
		end if
		if fluency<>"" then
			rs("fluency")=fluency
		end if
		
		if vocabulary<>"" then
			rs("vocabulary")=vocabulary
		end if
		if grammar<>"" then
			rs("grammar")=grammar
		end if
		if overall<>"" then
			rs("overall")=overall
		end if
		if related<>"" then
			rs("related")=related
		else
			rs("related")=null
		end if
		if idea<>"" then
			rs("idea")=idea
		else
			rs("idea")=null
		end if
		if comment<>"" then
			rs("comment")=comment
		else
			rs("comment")=null
		end if
		if others<>"" then
			rs("others")=others
		else
			rs("others")=null
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
		yn=trim(rs("yn"))
		category=trim(rs("category"))
		languagecode=trim(rs("languagecode"))
		signin=trim(rs("signin"))
		orallevel=trim(rs("orallevel"))
		oralset=trim(rs("oralset"))
		topic=trim(rs("topic"))


	end if
	rs.close
	'口語內容資訊
	sql = "select * from boo_op   where tid='"&id&"' "
	rs.Open sql,msconn,adOpenStatic,adLockReadonly
    if rs.EOF then
		validate="add"
	else
		validate="edit"
		did=trim(rs("id"))

		pronunciation=trim(rs("pronunciation"))
		fluency=trim(rs("fluency"))
		vocabulary=trim(rs("vocabulary"))
		grammar=trim(rs("grammar"))
		overall=trim(rs("overall"))
		teacher=trim(rs("teacher"))
		related=trim(rs("related"))
		idea=trim(rs("idea"))
		comment=trim(rs("comment"))
		others=trim(rs("others"))
	end if
	rs.close
end if

if teacher="" or isnull(teacher) or isempty(teacher) then
	teacher=teachername
end if


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
	if (form1.pronunciation.value=="")
		errmsg += "發音不能為空白\n";
	if (form1.fluency.value=="")
		errmsg += "流利度不能為空白\n";
	if (form1.vocabulary.value=="")
		errmsg += "單字不能為空白\n";
	if (form1.grammar.value=="")
		errmsg += "文法不能為空白\n";
	if (form1.overall.value=="")
		errmsg += "綜合不能為空白\n";
	if (form1.related.value=="")
		errmsg += "相關字彙不能為空白\n";
	if (form1.idea.value=="")
		errmsg += "好主意不能為空白\n";
	if (form1.comment.value=="")
		errmsg += "評語/建議不能為空白\n";
    if (errmsg=="")
        return true;
    else
    {
        alert(errmsg);
        return false;
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center">學員口說練習紀錄編輯</TD>
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
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR>
					<TD class="inputlabel">口語級數：</TD><TD><%=orallevel%></TD>
					<TD class="inputlabel">口語系列：</TD><TD><%=oralset%></TD>
					<TD class="inputlabel">口語題目：</TD><TD><%=topic%></TD>
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
			<form id="form1" name="form1" method="post" action="opedit.asp"   onsubmit="return check_input();">
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
			<input type="hidden" value="<%=orallevel%>" name="orallevel">
			<input type="hidden" value="<%=oralset%>" name="oralset">
			<input type="hidden" value="<%=topic%>" name="topic">
			<TABLE class=normal cellSpacing=0 cellPadding=0 width="100%"  border=0>
			<TR><TD>
					
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR class="inputlabel">
						<TD>發音(pronunciation)：</TD>
						<TD>流利度(fluency)：</TD>
					</TR>
					<TR>
						<TD>
						<select name="pronunciation" class="inputtext"  >
						<option value=""> - 未指定 - </option>
						<option value="未評" <% if pronunciation="未評" then response.write "selected" end if %>>未評</option>
						<option value="1" <% if pronunciation="1" then response.write "selected" end if %>>1 - Utterances are almost incomprehensible</option>
						<option value="2" <% if pronunciation="2" then response.write "selected" end if %>>2 - Severe Interference</option>
						<option value="3" <% if pronunciation="3" then response.write "selected" end if %>>3 - Substantial interference</option>
						<option value="4" <% if pronunciation="4" then response.write "selected" end if %>>4 - Occasional interference</option>
						<option value="5" <% if pronunciation="5" then response.write "selected" end if %>>5 - Little or no interference(from Chinese)</option>
						</select>
						
						</TD>
						<TD>
							<select name="fluency" class="inputtext"  >
							<option value=""> - 未指定 - </option>
							<option value="未評" <% if fluency="未評" then response.write "selected" end if %>>未評</option>
							<option value="1" <% if fluency="1" then response.write "selected" end if %>>1 - Complete communication breakdown</option>
							<option value="2" <% if fluency="2" then response.write "selected" end if %>>2 - Many pauses with communication breakdown</option>
							<option value="3" <% if fluency="3" then response.write "selected" end if %>>3 - Frequent hesitation, but no significant breakdown of communication</option>
							<option value="4" <% if fluency="4" then response.write "selected" end if %>>4 - Slight hesitation, natural pauses</option>
							<option value="5" <% if fluency="5" then response.write "selected" end if %>>5 - No hesitation, natural pauses</option>
							</select>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR class="inputlabel">
						<TD>單字(vocabulary)：</TD>
						<TD>文法(grammar)：</TD>
					</TR>
					<TR>
						<TD>
							<select name="vocabulary" class="inputtext"  >
							<option value=""> - 未指定 - </option>
							<option value="未評" <% if vocabulary="未評" then response.write "selected" end if %>>未評</option>
							<option value="1" <% if vocabulary="1" then response.write "selected" end if %>>1 - Few examples of correct usage; relies on Chinese words</option>
							<option value="2" <% if vocabulary="2" then response.write "selected" end if %>>2 - Substantial errors; little use of various vocabulary items</option>
							<option value="3" <% if vocabulary="3" then response.write "selected" end if %>>3 - Frequent errors; little use of various vocabulary items</option>
							<option value="4" <% if vocabulary="4" then response.write "selected" end if %>>4 - Mostly accuate; some use of various vocabulary items</option>
							<option value="5" <% if vocabulary="5" then response.write "selected" end if %>>5 - Accurate usage; showsknowledge of various vocabulary items</option>
							</select>
						</TD>
						<TD>
							<select name="grammar" class="inputtext"  >
							<option value=""> - 未指定 - </option>
							<option value="未評" <% if grammar="未評" then response.write "selected" end if %>>未評</option>
							<option value="1" <% if grammar="1" then response.write "selected" end if %>>1 - Few examples of correct usage</option>
							<option value="2" <% if grammar="2" then response.write "selected" end if %>>2 - Many errors;affect ability to communicate</option>
							<option value="3" <% if grammar="3" then response.write "selected" end if %>>3 - Several errors;some breakdown of communication</option>
							<option value="4" <% if grammar="4" then response.write "selected" end if %>>4 - One or two significant errors</option>
							<option value="5" <% if grammar="5" then response.write "selected" end if %>>5 - No significant errors</option>
							</select>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR class="inputlabel">
						
						<TD>綜合(overall)：</TD>
						<TD>駐診老師：</TD>
					</TR>
					<TR>
						
						<TD>
							<select name="overall" class="inputtext"  style="width:150">
							<option value=""> - 未指定 - </option>
							<option value="未評" <% if overall="未評" then response.write "selected" end if %>>未評</option>
							<option value="1" <% if overall="1" then response.write "selected" end if %>>1 - Limited</option>
							<option value="2" <% if overall="2" then response.write "selected" end if %>>2 - Modest</option>
							<option value="3" <% if overall="3" then response.write "selected" end if %>>3 - Competent</option>
							<option value="4" <% if overall="4" then response.write "selected" end if %>>4 - Good</option>
							<option value="5" <% if overall="5" then response.write "selected" end if %>>5 - Expert</option>
							</select>
						</TD>
						<TD><input type="text" value="<%=teacher%>" maxlength="25" size="35" name="teacher" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly ></TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
			<BR>
			<p class="T3">「口語練習」紀錄 Register of 「Oral Practice」</p>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>1、相關字彙 Related Vocabulary：</TD>
					</TR>
					<TR>
						<TD>
							<TABLE cellSpacing=1 cellPadding=1  border=0 >
							<TR>
							<TD><textarea name="related" rows="5" cols="100" class="inputtext"  ><%=related%></textarea></TD>
							</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>2、評語（優／缺點）Comment(s)：</TD>
					</TR>
					<TR>
						<TD>
							<TABLE cellSpacing=1 cellPadding=1  border=0 >
							<TR>
							<TD><textarea name="idea" rows="5" cols="100" class="inputtext"  ><%=idea%></textarea></TD>
							</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>3、待改善之處Improvement(s) Needed：</TD>
					</TR>
					<TR>
						<TD>
							<TABLE cellSpacing=1 cellPadding=1  border=0 >
							<TR>
							<TD><textarea name="comment" rows="5" cols="100" class="inputtext"  ><%=comment%></textarea></TD>
							</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>4、其他Others：</TD>
					</TR>
					<TR>
						<TD>
							<TABLE cellSpacing=1 cellPadding=1  border=0 >
							<TR>
							<TD><textarea name="others" rows="5" cols="100" class="inputtext"  ><%=others%></textarea></TD>
							</TR>
							</TABLE>
						</TD>
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


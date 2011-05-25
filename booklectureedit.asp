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
nextrec=trim(request("nextrec"))
id=trim(request("id"))
category=trim(request("category")) 

sid=trim(request("sid")) 
name=trim(request("name")) 
slevel=trim(request("slevel")) 
grade=trim(request("grade")) 
class1=trim(request("class1")) 
department=trim(request("department")) 

sender=ifnull(trim(request("sender")),"booklecturelist.asp?category=" & category&"&lid=" & id)

set rs = server.CreateObject("adodb.recordset")
if validate="cancel" then

	sqld="update boo_book_lecture set yn='N',canceldate= Convert(varchar(10),Getdate(),111),canceluid='"& session("sid")&"'   where id='"&id&"'"
	msconn.Execute sqld
	if Err.number<>0 then
	
        showmessage= Err.Description
    end if

end if

sql = "select a.*,b.* from boo_book_lecture a left join  boo_lecture b on a.lid=b.id  where a.id='"&id&"' "
rs.Open sql,msconn,adOpenStatic,adLockReadonly
if rs.EOF then
	response.redirect sender
else
	sid=trim(rs("sid"))
	name=trim(rs("name"))
	slevel=trim(rs("slevel"))
	grade=trim(rs("grade"))
	class1=trim(rs("class1"))
	department=trim(rs("department"))
	yn=trim(rs("yn"))
	canceldate=trim(rs("canceldate"))
	canceluid=trim(rs("canceluid"))
	signin=trim(rs("signin"))

	subject=trim(rs("subject")) 
	date1=trim(rs("date1")) 
	stimeh=left(trim(rs("stime")) ,2)
	etimeh=left(trim(rs("etime")) ,2)
	stimem=right(trim(rs("stime")) ,2)
	etimem=right(trim(rs("etime")) ,2)
	place=trim(rs("place")) 
	speaker=trim(rs("speaker")) 
	content=trim(rs("content")) 
	sdate=trim(rs("sdate")) 
	edate=trim(rs("edate")) 
end if
rs.close


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
	
    if (form1.sid.value=="")
        errmsg += "學號不能為空白\n";
	if (form1.name.value=="")
        errmsg += "姓名不能為空白\n";

    if (errmsg=="")
        return true;
    else
    {
        alert(errmsg);
        return false;
    }
}

function ChkStudent()
{
	vWinCal2 = window.open("lib/checkstudent.asp?sid="+form1.sid.value,"iframe_query", "width=200,height=50,status=no,resizable=no,scrollbars=no");
	vWinCal2.opener = form1;
}

function delete_record()
{
    if (confirm("您確定要取消該筆報名嗎?"))
    {
        form1.validate.value="cancel";
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
		<TD WIDTH="5" height="20"></TD><TD class="T3" valign="center"><%if category="L" then response.write "報名外語學習講座資料"  else response.write "報名處方課程"  end if%></TD>
	</TR>
	<TR >
		<TD WIDTH="5"></TD><TD class="errmsg"><%=showmessage%></TD>
	</TR>

	<TR>
		<TD></TD><TD valign="top">
			<TABLE cellSpacing=1 cellPadding=0 width="100%"  border=0   >
			<form id="form1" name="form1" method="post" action="booklectureedit.asp" >
			<input type="hidden" value="" name="validate">
			<input type="hidden" id="category" name="category" value="<%=category%>">
			<input type="hidden" id="id" name="id" value="<%=id%>">
			<input type="hidden" id="sender" name="sender" value="<%=sender%>">
			<TR><TD>
				<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD>學號：</TD>
						<TD>姓名：</TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=sid%>" maxlength="25" size="35"  name="sid" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly >
						</TD>
						<TD>
						<input type="text" value="<%=name%>" maxlength="25" size="35" name="name" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly >
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
						<TD>預約狀態：</TD>
						<TD></TD>
					</TR>
					<TR>
						<TD>
						<input type="text" value="<%=slevel%>" maxlength="10"   name="slevel" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						<TD>
						<input type="text" value="<%=department%>" maxlength="10"  name="department" class="inputtext" style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						<TD>
						<input type="text" value="<%=grade%>" maxlength="10" size="10"  name="grade" class="inputtext"  style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						
						<TD>
						<input type="text" value="<%=class1%>" maxlength="25"  name="class1" class="inputtext"  style="BORDER-RIGHT:#f0f0f0 2px solid;BORDER-TOP:#f0f0f0 2px solid;BORDER-LEFT:#f0f0f0 2px solid;BORDER-BOTTOM:#f0f0f0 2px solid" readonly>
						</TD>
						<TD>&nbsp;&nbsp;<%=replace(replace(yn,"Y","<font color=""blue"">已報名</font>"),"N","<font color=""red"">取消</font>")%></TD>
						<TD>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR>
			<TD>
			<BR>
			<input  type="button"  onclick="delete_record();" value="取消<%if category="L" then response.write "外語學習講座"  else response.write "處方課程"  end if%>" class="inputbutton"  <%if yn="N" then response.write "disabled" end if%>>
			<input  type="button" value="返回" class="inputbutton" onclick="window.location='<%=sender%>'">
			<BR><BR>
			</TD>
			</TR>
			</form>
			</TABLE>
			<font color="#FF0000"><%if category="L" then response.write "外語學習講座"  else response.write "處方課程"  end if%>更詳細資訊如下：</font>
			<TABLE cellSpacing=2 cellPadding=3 width="90%"  border=0   bgcolor="#F3EFE6" style="border: 1px solid #FF66CC; padding: 0">
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR >
						<TD class="inputlabel"><%if category="L" then response.write "外語學習講座名稱"  else response.write "處方課程名稱"  end if%>：</TD>
						<TD>
						<%=subject%>
						</TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR >
						<TD class="inputlabel">日期：</TD><TD><%=date1%>&nbsp;</TD>
						<TD class="inputlabel">時間起迄：</TD>
						
							<TD>
							<TABLE cellSpacing=0 cellPadding=0  border=0 >
							<TR>
							<TD>
								<%=stimeh%>時<%=stimem%>分
							</TD>
							<TD class="inputlabel">&nbsp;~&nbsp;</TD>
							<TD>
								<%=etimeh%>時<%=etimem%>分
							</TD>
							</TR>
							</TABLE>
						</TD>
						<TD class="inputlabel">地點：</TD>
						<TD><%=place%></TD>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR >
						<TD class="inputlabel"><%if category="L" then response.write "主講人"  else response.write "老師"  end if%>：</TD>
						<TD><%=speaker%></TD>
						<TD class="inputlabel">開始報名日期：</TD>
						<TD><%=sdate%></TD>
						<TD class="inputlabel">結束報名日期：</TD>
						<TD><%=edate%></TD>
					</TR>
					<TR>
					</TR>
				</TABLE>
			</TD></TR>
			<TR><TD>
				<TABLE cellSpacing=1 cellPadding=2  border=0 >
					<TR >
						<TD class="inputlabel" nowrap><%if category="L" then response.write "講座內容"  else response.write "課程內容"  end if%>：</TD>
					</TR>
					<TR>
					<TD><%=content%></TD>
					</TR>
				</TABLE>
			</TD></TR>
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
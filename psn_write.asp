<%
'�g�@
sql = "select a.id as bookid,a.sid,a.name,a.department,a.grade,a.class1,a.bdate,a.btime,a.timeflag,a.score,a.teachername,b.* from boo_book_T_M a  "
sql = sql & " inner join boo_write b on a.id=b.tid  where a.item='�g�@'  and  a.signin is not null  "
sql = sql & " and a.sid='"&sid&"'  order by bdate desc "
rs.Open sql,msconn,adOpenStatic,adLockReadonly

%>
<TABLE cellSpacing=1 cellPadding=2 border=0 width="100%"   name="folder_4" id="folder_4"  style="display:<%if forderid="4" then%>block<%else%>none<%end if%>">
<TR>
	<TD>
	<% if rs.EOF then
			response.write "<font   class=""norecord"" >�S���ŦX���󪺼g�@�t�Ӭ������</font>"
	else
		while not rs.EOF 
			
	%>
			<font color="#FF0000">�w������G<%=rs("bdate")%>&nbsp;&nbsp;�n�E�Ѯv�G<%=rs("teachername")%>&nbsp;</font>
			<TABLE cellSpacing=1 cellPadding=2  border=0 width="90%"   bgcolor="#F3EFE6" style="border: 1px solid #FF66CC; padding: 0">
				<TR><TD>
					<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD colspan="2">�g�@�D�D�G</TD>
					</TR>
					<TR >
						<TD width="10"></TD><TD><%=rs("subject")%>&nbsp;</TD>
					</TR>
					<TR class="inputlabel">
						<TD colspan="2">�g�@���D�G</TD>
					</TR>
					<TR >
						<TD width="10"></TD><TD><%=rs("content")%>&nbsp;</TD>
					</TR>
					
					<TR class="inputlabel">
						<TD  colspan="2">�Ѯv�^�X�G</TD>
					</TR>
					<TR>
						<TD ></TD>
						<TD><%if rs("feedback")<>"" then response.write  rs("feedback") else response.write "�L" end if%>&nbsp;</TD>
					</TR>

				</TABLE>



				</TD></TR>
				</TABLE><BR>

	<%
			rs.MoveNext
		wend
	end if
	rs.close
	%>
	
	</TD>
</TR>
</TABLE>
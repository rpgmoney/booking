<%
'�f�y




sql = "select a.id as bookid,a.sid,a.name,a.department,a.grade,a.class1,a.bdate,a.btime,a.timeflag,a.score,a.teachername,a.orallevel,a.oralset,a.topic,b.* from boo_book_T_M a  "
sql = sql & " inner join boo_op b on a.id=b.tid  where a.item='�f�y'  and  a.signin is not null  "
sql = sql & " and a.sid='"&sid&"'  order by bdate desc "
rs.Open sql,msconn,adOpenStatic,adLockReadonly

%>
<TABLE cellSpacing=1 cellPadding=2 border=0 width="100%"   name="folder_3" id="folder_3"  style="display:<%if forderid="3" then%>block<%else%>none<%end if%>">
<TR>
	<TD>
	<% if rs.EOF then
			response.write "<font   class=""norecord"" >�S���ŦX���󪺤f���m�߬������</font>"
	else
		while not rs.EOF 
			
	%>
			<font color="#FF0000">�w������G<%=rs("bdate")%>&nbsp;&nbsp;�n�E�Ѯv�G<%=rs("teachername")%>&nbsp;</font>
			<TABLE cellSpacing=1 cellPadding=2  border=0 width="90%"   bgcolor="#F3EFE6" style="border: 1px solid #FF66CC; padding: 0">
			<TR><TD>
				<TABLE cellSpacing=2 cellPadding=3  border=0  class=normal >
				<TR>
				<TD class="inputlabel">�f�y�żơG</TD><TD><%=rs("orallevel")%></TD>
				<TD class="inputlabel">�f�y�t�C�G</TD><TD><%=rs("oralset")%></TD>
				</TR>
				</TABLE>
				<TABLE cellSpacing=2 cellPadding=3  border=0  class=normal >
				<TR>
				<TD class="inputlabel">�f�y�D�ءG</TD><TD><%=rs("topic")%></TD>
				</TR>
				</TABLE>
				<TABLE cellSpacing=2 cellPadding=3  border=0 >
					<TR >
						<TD class="inputlabel">�o��(pronunciation)�G</TD><TD><%=rs("pronunciation")%></TD>
						<TD class="inputlabel">�y�Q��(fluency)�G</TD><TD><%=rs("fluency")%></TD>
						<TD class="inputlabel">��r(vocabulary)�G</TD><TD><%=rs("vocabulary")%></TD>
					</TR>
				</TABLE>	
			
				<TABLE cellSpacing=2 cellPadding=3  border=0 class=normal >
					<TR>
						<TD class="inputlabel">��k(grammar)�G</TD><TD><%=rs("grammar")%></TD>
						<TD class="inputlabel">��X(overall)�G</TD><TD><%=rs("overall")%></TD>
					</TR>
				</TABLE>
				<TABLE cellSpacing=2 cellPadding=3  border=0  class=normal >
				<TR>
				<TD class="inputlabel">1�B�����r�J Related Vocabulary�G</TD>
				</TR>
				<TR>
				<TD><%=rs("related")%></TD>
				</TR>
				</TABLE>
				<TABLE cellSpacing=2 cellPadding=3  border=0  class=normal >
				<TR>
				<TD class="inputlabel">2�B���y�]�u�����I�^Comment(s)�G</TD>
				</TR>
				<TR>
				<TD><%=rs("idea")%></TD>
				</TR>
				</TABLE>
				<TABLE cellSpacing=2 cellPadding=3  border=0  class=normal >
				<TR>
				<TD class="inputlabel">3�B�ݧﵽ���BImprovement(s) Needed�G</TD>
				</TR>
				<TR>
				<TD><%=rs("comment")%></TD>
				</TR>
				</TABLE>
				<TABLE cellSpacing=2 cellPadding=3  border=0  class=normal >
				<TR>
				<TD class="inputlabel">4�B��LOthers�G</TD>
				</TR>
				<TR>
				<TD><%=rs("others")%></TD>
				</TR>
				</TABLE>
			</TD></TR>
			</TABLE>
			<BR>

	<%
			rs.MoveNext
		wend
	end if
	rs.close
	%>
	
	</TD>
</TR>
</TABLE>
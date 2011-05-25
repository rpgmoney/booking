<%
'口語




sql = "select a.id as bookid,a.sid,a.name,a.department,a.grade,a.class1,a.bdate,a.btime,a.timeflag,a.score,a.teachername,a.orallevel,a.oralset,a.topic,b.* from boo_book_T_M a  "
sql = sql & " inner join boo_op b on a.id=b.tid  where a.item='口語'  and  a.signin is not null  "
sql = sql & " and a.sid='"&sid&"'  order by bdate desc "
rs.Open sql,msconn,adOpenStatic,adLockReadonly

%>
<TABLE cellSpacing=1 cellPadding=2 border=0 width="100%"   name="folder_3" id="folder_3"  style="display:<%if forderid="3" then%>block<%else%>none<%end if%>">
<TR>
	<TD>
	<% if rs.EOF then
			response.write "<font   class=""norecord"" >沒有符合條件的口說練習紀錄顯示</font>"
	else
		while not rs.EOF 
			
	%>
			<font color="#FF0000">預約日期：<%=rs("bdate")%>&nbsp;&nbsp;駐診老師：<%=rs("teachername")%>&nbsp;</font>
			<TABLE cellSpacing=1 cellPadding=2  border=0 width="90%"   bgcolor="#F3EFE6" style="border: 1px solid #FF66CC; padding: 0">
			<TR><TD>
				<TABLE cellSpacing=2 cellPadding=3  border=0  class=normal >
				<TR>
				<TD class="inputlabel">口語級數：</TD><TD><%=rs("orallevel")%></TD>
				<TD class="inputlabel">口語系列：</TD><TD><%=rs("oralset")%></TD>
				</TR>
				</TABLE>
				<TABLE cellSpacing=2 cellPadding=3  border=0  class=normal >
				<TR>
				<TD class="inputlabel">口語題目：</TD><TD><%=rs("topic")%></TD>
				</TR>
				</TABLE>
				<TABLE cellSpacing=2 cellPadding=3  border=0 >
					<TR >
						<TD class="inputlabel">發音(pronunciation)：</TD><TD><%=rs("pronunciation")%></TD>
						<TD class="inputlabel">流利度(fluency)：</TD><TD><%=rs("fluency")%></TD>
						<TD class="inputlabel">單字(vocabulary)：</TD><TD><%=rs("vocabulary")%></TD>
					</TR>
				</TABLE>	
			
				<TABLE cellSpacing=2 cellPadding=3  border=0 class=normal >
					<TR>
						<TD class="inputlabel">文法(grammar)：</TD><TD><%=rs("grammar")%></TD>
						<TD class="inputlabel">綜合(overall)：</TD><TD><%=rs("overall")%></TD>
					</TR>
				</TABLE>
				<TABLE cellSpacing=2 cellPadding=3  border=0  class=normal >
				<TR>
				<TD class="inputlabel">1、相關字彙 Related Vocabulary：</TD>
				</TR>
				<TR>
				<TD><%=rs("related")%></TD>
				</TR>
				</TABLE>
				<TABLE cellSpacing=2 cellPadding=3  border=0  class=normal >
				<TR>
				<TD class="inputlabel">2、評語（優／缺點）Comment(s)：</TD>
				</TR>
				<TR>
				<TD><%=rs("idea")%></TD>
				</TR>
				</TABLE>
				<TABLE cellSpacing=2 cellPadding=3  border=0  class=normal >
				<TR>
				<TD class="inputlabel">3、待改善之處Improvement(s) Needed：</TD>
				</TR>
				<TR>
				<TD><%=rs("comment")%></TD>
				</TR>
				</TABLE>
				<TABLE cellSpacing=2 cellPadding=3  border=0  class=normal >
				<TR>
				<TD class="inputlabel">4、其他Others：</TD>
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
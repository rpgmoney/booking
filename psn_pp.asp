<%
'簡報
sql = "select a.id as bookid,a.sid,a.name,a.department,a.grade,a.class1,a.bdate,a.btime,a.timeflag,a.score,a.teachername,a.briefing,b.* from boo_book_T_M a  "
sql = sql & " inner join boo_pp b on a.id=b.tid  where a.item='簡報'  and  a.signin is not null  "
sql = sql & " and a.sid='"&sid&"'  order by bdate desc "
rs.Open sql,msconn,adOpenStatic,adLockReadonly

%>
<TABLE cellSpacing=1 cellPadding=2 border=0 width="100%"   name="folder_6" id="folder_6"  style="display:<%if forderid="6" then%>block<%else%>none<%end if%>">
<TR>
	<TD>
	<% if rs.EOF then
			response.write "<font   class=""norecord"" >沒有符合條件的簡報練習紀錄資料顯示</font>"
	else
		while not rs.EOF 
			
	%>
			<font color="#FF0000">預約日期：<%=rs("bdate")%>&nbsp;&nbsp;駐診老師：<%=rs("teachername")%>&nbsp;</font>
			<TABLE cellSpacing=1 cellPadding=2  border=0 width="90%"   bgcolor="#F3EFE6" style="border: 1px solid #FF66CC; padding: 0">
				<TR><TD>
					<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD colspan="2">簡報題目：</TD>
					</TR>
					<TR >
						<TD width="10"></TD><TD><%=rs("briefing")%>&nbsp;</TD>
					</TR>
					
					<TR class="inputlabel">
						<TD  colspan="2">老師回饋：</TD>
					</TR>
					<TR>
						<TD ></TD>
						<TD><%if rs("feedback")<>"" then response.write  rs("feedback") else response.write "無" end if%>&nbsp;</TD>
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
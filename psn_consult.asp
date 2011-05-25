<%
'諮商
sql = "select a.id as bookid,a.sid,a.name,a.department,a.grade,a.class1,a.bdate,a.btime,a.timeflag,a.score,a.teachername,a.consult,b.* from boo_book_T_M a  "
sql = sql & " inner join boo_consult b on a.id=b.tid  where a.item='諮商'  and  a.signin is not null  "
sql = sql & " and a.sid='"&sid&"'  order by bdate desc "
rs.Open sql,msconn,adOpenStatic,adLockReadonly

%>
<TABLE cellSpacing=1 cellPadding=2 border=0 width="100%"   name="folder_2" id="folder_2"  style="display:<%if forderid="2" then%>block<%else%>none<%end if%>">
<TR>
	<TD>
	<% if rs.EOF then
			response.write "<font   class=""norecord"" >沒有符合條件的咨商紀錄顯示</font>"
	else
		while not rs.EOF 
			
	%>
			<font color="#FF0000">預約日期：<%=rs("bdate")%>&nbsp;&nbsp;駐診老師：<%=rs("teachername")%>&nbsp;</font>
			<TABLE cellSpacing=1 cellPadding=2  border=0 width="90%"   bgcolor="#F3EFE6" style="border: 1px solid #FF66CC; padding: 0">
				<TR><TD>
					<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD colspan="2">諮商主題：</TD>
					</TR>
					<TR >
						<TD width="10"></TD><TD><%=ifnull(rs("consult"),"無")%>&nbsp;</TD>
					</TR>
					<TR class="inputlabel">
						<TD colspan="2">諮商內容：</TD>
					</TR>
					<TR >
						<TD width="10"></TD><TD><%=rs("content")%>&nbsp;</TD>
					</TR>
					
					<TR class="inputlabel">
						<TD  colspan="2">老師回饋：</TD>
					</TR>
					<TR>
						<TD ></TD>
						<TD><%if rs("feedback")<>"" then response.write  rs("feedback") else response.write "無" end if%>&nbsp;</TD>
					</TR>
					<TR class="inputlabel">
						<TD  colspan="2">優點/缺點Strength(s)/Weakness(es)：</TD>
					</TR>
					<TR>
						<TD ></TD>
						<TD><%if rs("strength")<>"" then response.write  rs("strength") else response.write "無" end if%>&nbsp;</TD>
					</TR>
					<TR class="inputlabel">
						<TD  colspan="2">須改善之處Tmprovement(s) Needed：</TD>
					</TR>
					<TR>
						<TD ></TD>
						<TD><%if rs("needed")<>"" then response.write  rs("needed") else response.write "無" end if%>&nbsp;</TD>
					</TR>
					<TR class="inputlabel">
						<TD  colspan="2">預期成效Anticipated Effects：</TD>
					</TR>
					<TR>
						<TD ></TD>
						<TD><%if rs("effect")<>"" then response.write  rs("effect") else response.write "無" end if%>&nbsp;</TD>
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
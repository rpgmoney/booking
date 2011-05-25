<%
'詩歌
sql = "select a.id as bookid,a.sid,a.name,a.department,a.grade,a.class1,a.bdate,a.btime,a.timeflag,a.score,a.teachername,a.briefing,b.* from boo_book_T_M a  "
sql = sql & " inner join boo_crkp b on a.id=b.tid  where a.item='詩歌'  and  a.signin is not null  "
sql = sql & " and a.sid='"&sid&"'  order by bdate desc "
rs.Open sql,msconn,adOpenStatic,adLockReadonly

%>
<TABLE cellSpacing=1 cellPadding=2 border=0 width="100%"   name="folder_5" id="folder_5"  style="display:<%if forderid="5" then%>block<%else%>none<%end if%>">
<TR>
	<TD>
	<% if rs.EOF then
			response.write "<font   class=""norecord"" >沒有符合條件的詩歌饒舌練習紀錄顯示</font>"
	else
		while not rs.EOF 
			
	%>
			<font color="#FF0000">預約日期：<%=rs("bdate")%>&nbsp;&nbsp;駐診老師：<%=rs("teachername")%>&nbsp;</font>
			<TABLE cellSpacing=1 cellPadding=2  border=0 width="90%"   bgcolor="#F3EFE6" style="border: 1px solid #FF66CC; padding: 0">
				<TR><TD>
					<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD colspan="2">
							<TABLE cellSpacing=1 cellPadding=2  border=0 >
							<TR class="inputlabel">
								<TD>詩歌：</TD><TD><input type="checkbox"  value="Y"  name="chant" <%if  rs("chant")="Y" then response.write "checked" end if %> class="inputtext"  ></TD>
								<TD>饒舌：</TD><TD><input type="checkbox"  value="Y"  name="rhyme"  <%if  rs("rhyme")="Y" then response.write "checked" end if %> class="inputtext" ></TD>
								<TD>歌曲：</TD><TD><input type="checkbox"  value="Y"  name="song"  <%if  rs("song")="Y" then response.write "checked" end if %> class="inputtext" ></TD>
							</TR>
							</TABLE>
						</TD>
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
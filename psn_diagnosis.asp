<%
set rsLoad = server.CreateObject("adodb.recordset")
'診斷
sql = "select a.id as bookid,a.sid,a.name,a.department,a.grade,a.class1,a.bdate,a.btime,a.timeflag,a.score,a.teachername,a.score,b.* from boo_book_T_M a  "
sql = sql & " inner join boo_diagnosis b on a.id=b.tid  where a.item='診斷'  and  a.signin is not null  "
sql = sql & " and a.sid='"&sid&"'  order by bdate desc "
rs.Open sql,msconn,adOpenStatic,adLockReadonly

%>
<TABLE cellSpacing=1 cellPadding=2 border=0 width="100%"   name="folder_1" id="folder_1"  style="display:<%if forderid="1" then%>block<%else%>none<%end if%>">
<TR>
	<TD>
	<% if rs.EOF then
			response.write "<font   class=""norecord"" >沒有符合條件的診斷紀錄顯示</font>"
	else
		while not rs.EOF 
			
	%>
			<font color="#FF0000">預約日期：<%=rs("bdate")%>&nbsp;&nbsp;駐診老師：<%=rs("teachername")%>&nbsp;回診日期：<%=rs("backdate")%></font>
			<TABLE cellSpacing=1 cellPadding=2  border=0 width="90%"   bgcolor="#F3EFE6" style="border: 1px solid #FF66CC; padding: 0">
				<TR><TD>
					<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD colspan="2">診斷內容：</TD>
					</TR>
					<TR >
						<TD width="10"></TD><TD><%=rs("content")%>&nbsp;</TD>
					</TR>
					<TR class="inputlabel">
						<TD colspan="2">優點/缺點Strength(s)/Weakness(es)：</TD>
					</TR>
					<TR >
						<TD width="10"></TD><TD><%=rs("strength")%>&nbsp;</TD>
					</TR>
					<TR class="inputlabel">
						<TD colspan="2">須改善之處 Improvement(s) Needed：</TD>
					</TR>
					<TR >
						<TD width="10"></TD><TD><%=rs("needed")%>&nbsp;</TD>
					</TR>
					<TR class="inputlabel">
						<TD  colspan="2">建議（Recommendation）：</TD>
					</TR>
					<TR>
						<TD ></TD>
						<TD>
							<TABLE cellSpacing=1 cellPadding=2  border=0 >
							<TR><TD  class="inputlabel">口語：</TD>
							<TD><%=rs("optime")%>&nbsp;</TD>
							<TD>次(一次為上下二節50分)</TD>
							<TD  class="inputlabel">已預約：</TD>
							<TD><%=rs("optime_b")%>&nbsp;</TD>
							<TD>次</TD>
							<TD  class="inputlabel">已完成：</TD>
							<TD><%=rs("optime_c")%>&nbsp;</TD>
							<TD>次</TD>
							</TR>
							<TR>
							<TD class="inputlabel">簡報：</TD>
							<TD><%=rs("pptime")%>&nbsp;</TD>
							
							<TD>次(一次為上下二節50分)</TD>
							<TD  class="inputlabel">已預約：</TD>
							<TD><%=rs("pptime_b")%>&nbsp;</TD>
							<TD>次</TD>
							<TD  class="inputlabel">已完成：</TD>
							<TD><%=rs("pptime_c")%>&nbsp;</TD>
							<TD>次</TD>
							</TR>
							<TR><TD class="inputlabel">詩歌：</TD>
							<TD><%=rs("crkptime")%>&nbsp;</TD>
							
							<TD>次(一次為上下二節50分)</TD>
							<TD  class="inputlabel">已預約：</TD>
							<TD><%=rs("crkptime_b")%>&nbsp;</TD>
							<TD>次</TD>
							<TD  class="inputlabel">已完成：</TD>
							<TD><%=rs("crkptime_c")%>&nbsp;</TD>
							<TD>次</TD>
							</TR>
							<TR>
							<TD class="inputlabel">寫作：</TD>
							<TD><%=rs("writetime")%>&nbsp;</TD>
							
							<TD>次(一次為上下二節50分)</TD>
							<TD  class="inputlabel">已預約：</TD>
							<TD><%=rs("writetime_b")%>&nbsp;</TD>
							<TD>次</TD>
							<TD  class="inputlabel">已完成：</TD>
							<TD><%=rs("writetime_c")%>&nbsp;</TD>
							<TD>次</TD>
							</TR>
							<TR><TD class="inputlabel">閱讀：</TD>
							<TD><%=rs("readtime")%>&nbsp;</TD>
							
							<TD>次(一次為上下二節50分)</TD>
							<TD  class="inputlabel">已預約：</TD>
							<TD><%=rs("readtime_b")%>&nbsp;</TD>
							<TD>次</TD>
							<TD  class="inputlabel">已完成：</TD>
							<TD><%=rs("readtime_c")%>&nbsp;</TD>
							<TD>次</TD>
							</TR>

							<TR>
							<TD class="inputlabel" colspan="9">自學軟體：</TD>
							</TR>
							<TR>
							<TD  colspan="9">
								<TABLE cellSpacing=1 cellPadding=2  border=0 >
									<%
									'軟體
									
									sql = "select a.*,b.times,b.times_b,b.times_c  from "
									sql = sql & " ( "
									sql = sql & " select * from boo_software where yn='Y' and category='S' "
									sql = sql & " ) a left join "
									sql = sql & " ( "
									sql = sql & " select  sid,times,times_b,times_c  from boo_diagnosis_softwore  where  tid='"&rs("bookid") &"' and  category='S' "
									sql = sql & " ) b on a.id=b.sid  order by floor,software"

									'response.write sql
									rsLoad.Open sql,msconn,adOpenStatic,adLockReadonly
									Str=""
									i = 0
									while not rsLoad.EOF
										i= i +1
									%>
										<TR><TD></TD><TD>&nbsp;<% =rsLoad("floor")%>&nbsp;-&nbsp;<% =rsLoad("software")%>&nbsp;&nbsp;</TD><TD width="30"></TD>
										<TD><%=rsLoad("times")%></TD>
											<TD>&nbsp;&nbsp;分鐘&nbsp;&nbsp;</TD>
											<TD  class="inputlabel">已預約：</TD>
											<TD>&nbsp;&nbsp;<%=rsLoad("times_b")%>&nbsp;&nbsp;</TD>
											<TD  class="inputlabel">已完成：</TD>
											<TD>&nbsp;&nbsp;<%=rsLoad("times_c")%>&nbsp;&nbsp;</TD>
										
										</TR>
									<%
										rsLoad.MoveNext 
									wend
									rsLoad.close
									
									response.write Str
									%>
								</TABLE>									
							</TD>
							</TR>
							<TR>
							<TD class="inputlabel" colspan="9">補充教材：</TD>
							</TR>
							<TR>
							<TD  colspan="9">
								<TABLE cellSpacing=1 cellPadding=2  border=0 >
									<%
									'補充教材：
									
									sql = "select a.*,b.times,b.times_b,b.times_c  from "
									sql = sql & " ( "
									sql = sql & " select * from boo_software where yn='Y' and category='T' "
									sql = sql & " ) a left join "
									sql = sql & " ( "
									sql = sql & " select  sid,times,times_b,times_c  from boo_diagnosis_softwore  where  tid='"&rs("bookid") &"' and  category='T' "
									sql = sql & " ) b on a.id=b.sid  order by floor,software"

									'response.write sql
									rsLoad.Open sql,msconn,adOpenStatic,adLockReadonly
									Str=""
									i = 0
									while not rsLoad.EOF
										i= i +1
									%>
										<TR><TD></TD><TD>&nbsp;<% =rsLoad("software")%>&nbsp;&nbsp;</TD><TD width="30"></TD>
										<TD><%=rsLoad("times")%></TD>
											<TD>&nbsp;&nbsp;分鐘&nbsp;&nbsp;</TD>
											<TD  class="inputlabel">已預約：</TD>
											<TD>&nbsp;&nbsp;<%=rsLoad("times_b")%>&nbsp;&nbsp;</TD>
											<TD  class="inputlabel">已完成：</TD>
											<TD>&nbsp;&nbsp;<%=rsLoad("times_c")%>&nbsp;&nbsp;</TD>
										
										</TR>
									<%
										rsLoad.MoveNext 
									wend
									rsLoad.close
									
									response.write Str
									%>
								</TABLE>									
							</TD>
							</TR>
							</TABLE>
							
					<TR class="inputlabel">
						<TD  colspan="2">預期成效Anticipated Effects：</TD>
					</TR>
					<TR>
						<TD ></TD>
						<TD><%if rs("effect")<>"" then response.write  rs("effect") else response.write "無" end if%>&nbsp;</TD>
					</TR>	
					<TR class="inputlabel">
						<TD  colspan="2">備註說明：</TD>
					</TR>
					<TR>
						<TD ></TD>
						<TD><%if rs("note")<>"" then response.write  rs("note") else response.write "無" end if%>&nbsp;</TD>
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
<%
set rsLoad = server.CreateObject("adodb.recordset")
'�E�_
sql = "select a.id as bookid,a.sid,a.name,a.department,a.grade,a.class1,a.bdate,a.btime,a.timeflag,a.score,a.teachername,a.score,b.* from boo_book_T_M a  "
sql = sql & " inner join boo_diagnosis b on a.id=b.tid  where a.item='�E�_'  and  a.signin is not null  "
sql = sql & " and a.sid='"&sid&"'  order by bdate desc "
rs.Open sql,msconn,adOpenStatic,adLockReadonly

%>
<TABLE cellSpacing=1 cellPadding=2 border=0 width="100%"   name="folder_1" id="folder_1"  style="display:<%if forderid="1" then%>block<%else%>none<%end if%>">
<TR>
	<TD>
	<% if rs.EOF then
			response.write "<font   class=""norecord"" >�S���ŦX���󪺶E�_�������</font>"
	else
		while not rs.EOF 
			
	%>
			<font color="#FF0000">�w������G<%=rs("bdate")%>&nbsp;&nbsp;�n�E�Ѯv�G<%=rs("teachername")%>&nbsp;�^�E����G<%=rs("backdate")%></font>
			<TABLE cellSpacing=1 cellPadding=2  border=0 width="90%"   bgcolor="#F3EFE6" style="border: 1px solid #FF66CC; padding: 0">
				<TR><TD>
					<TABLE cellSpacing=0 cellPadding=0  border=0 >
					<TR class="inputlabel">
						<TD colspan="2">�E�_���e�G</TD>
					</TR>
					<TR >
						<TD width="10"></TD><TD><%=rs("content")%>&nbsp;</TD>
					</TR>
					<TR class="inputlabel">
						<TD colspan="2">�u�I/���IStrength(s)/Weakness(es)�G</TD>
					</TR>
					<TR >
						<TD width="10"></TD><TD><%=rs("strength")%>&nbsp;</TD>
					</TR>
					<TR class="inputlabel">
						<TD colspan="2">���ﵽ���B Improvement(s) Needed�G</TD>
					</TR>
					<TR >
						<TD width="10"></TD><TD><%=rs("needed")%>&nbsp;</TD>
					</TR>
					<TR class="inputlabel">
						<TD  colspan="2">��ĳ�]Recommendation�^�G</TD>
					</TR>
					<TR>
						<TD ></TD>
						<TD>
							<TABLE cellSpacing=1 cellPadding=2  border=0 >
							<TR><TD  class="inputlabel">�f�y�G</TD>
							<TD><%=rs("optime")%>&nbsp;</TD>
							<TD>��(�@�����W�U�G�`50��)</TD>
							<TD  class="inputlabel">�w�w���G</TD>
							<TD><%=rs("optime_b")%>&nbsp;</TD>
							<TD>��</TD>
							<TD  class="inputlabel">�w�����G</TD>
							<TD><%=rs("optime_c")%>&nbsp;</TD>
							<TD>��</TD>
							</TR>
							<TR>
							<TD class="inputlabel">²���G</TD>
							<TD><%=rs("pptime")%>&nbsp;</TD>
							
							<TD>��(�@�����W�U�G�`50��)</TD>
							<TD  class="inputlabel">�w�w���G</TD>
							<TD><%=rs("pptime_b")%>&nbsp;</TD>
							<TD>��</TD>
							<TD  class="inputlabel">�w�����G</TD>
							<TD><%=rs("pptime_c")%>&nbsp;</TD>
							<TD>��</TD>
							</TR>
							<TR><TD class="inputlabel">�ֺq�G</TD>
							<TD><%=rs("crkptime")%>&nbsp;</TD>
							
							<TD>��(�@�����W�U�G�`50��)</TD>
							<TD  class="inputlabel">�w�w���G</TD>
							<TD><%=rs("crkptime_b")%>&nbsp;</TD>
							<TD>��</TD>
							<TD  class="inputlabel">�w�����G</TD>
							<TD><%=rs("crkptime_c")%>&nbsp;</TD>
							<TD>��</TD>
							</TR>
							<TR>
							<TD class="inputlabel">�g�@�G</TD>
							<TD><%=rs("writetime")%>&nbsp;</TD>
							
							<TD>��(�@�����W�U�G�`50��)</TD>
							<TD  class="inputlabel">�w�w���G</TD>
							<TD><%=rs("writetime_b")%>&nbsp;</TD>
							<TD>��</TD>
							<TD  class="inputlabel">�w�����G</TD>
							<TD><%=rs("writetime_c")%>&nbsp;</TD>
							<TD>��</TD>
							</TR>
							<TR><TD class="inputlabel">�\Ū�G</TD>
							<TD><%=rs("readtime")%>&nbsp;</TD>
							
							<TD>��(�@�����W�U�G�`50��)</TD>
							<TD  class="inputlabel">�w�w���G</TD>
							<TD><%=rs("readtime_b")%>&nbsp;</TD>
							<TD>��</TD>
							<TD  class="inputlabel">�w�����G</TD>
							<TD><%=rs("readtime_c")%>&nbsp;</TD>
							<TD>��</TD>
							</TR>

							<TR>
							<TD class="inputlabel" colspan="9">�۾ǳn��G</TD>
							</TR>
							<TR>
							<TD  colspan="9">
								<TABLE cellSpacing=1 cellPadding=2  border=0 >
									<%
									'�n��
									
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
											<TD>&nbsp;&nbsp;����&nbsp;&nbsp;</TD>
											<TD  class="inputlabel">�w�w���G</TD>
											<TD>&nbsp;&nbsp;<%=rsLoad("times_b")%>&nbsp;&nbsp;</TD>
											<TD  class="inputlabel">�w�����G</TD>
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
							<TD class="inputlabel" colspan="9">�ɥR�Ч��G</TD>
							</TR>
							<TR>
							<TD  colspan="9">
								<TABLE cellSpacing=1 cellPadding=2  border=0 >
									<%
									'�ɥR�Ч��G
									
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
											<TD>&nbsp;&nbsp;����&nbsp;&nbsp;</TD>
											<TD  class="inputlabel">�w�w���G</TD>
											<TD>&nbsp;&nbsp;<%=rsLoad("times_b")%>&nbsp;&nbsp;</TD>
											<TD  class="inputlabel">�w�����G</TD>
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
						<TD  colspan="2">�w������Anticipated Effects�G</TD>
					</TR>
					<TR>
						<TD ></TD>
						<TD><%if rs("effect")<>"" then response.write  rs("effect") else response.write "�L" end if%>&nbsp;</TD>
					</TR>	
					<TR class="inputlabel">
						<TD  colspan="2">�Ƶ������G</TD>
					</TR>
					<TR>
						<TD ></TD>
						<TD><%if rs("note")<>"" then response.write  rs("note") else response.write "�L" end if%>&nbsp;</TD>
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
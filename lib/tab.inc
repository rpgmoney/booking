<%
if forderid<>"" then
%>
<SCRIPT LANGUAGE=javascript FOR=window EVENT=onload>
<!--
sel_tab("<%=forderid%>");
//alert("<%=sid%>");
//-->
</SCRIPT>
<%
end if
%>
<table cellspacing="0" cellpadding="1" border="0" width="100%" align="center">
<tr height="5"><td></td></tr>
<tr valign="bottom" height="20">
<td width="5" nowrap></td>
<td width="100" class="tabactive" align="center" nowrap id="tab_1" name="tab_1" language="javascript" onclick="sel_tab('1');">�E�_����</td>
<td width="1" nowrap></td>
<td width="100" class="tabinactive" align="center" nowrap id="tab_2" name="tab_2" language="javascript" onclick="sel_tab('2')">�t�Ӭ���</td>
<td width="1" nowrap></td>
<td width="100" class="tabinactive" align="center" nowrap id="tab_3" name="tab_3" language="javascript" onclick="sel_tab('3')">�f���m�߬���</td>
<td width="1" nowrap></td>
<td width="100" class="tabinactive" align="center" nowrap id="tab_4" name="tab_4" language="javascript" onclick="sel_tab('4')">�g�@�t�Ӭ���</td>
<td width="1" nowrap></td>
<td width="100" class="tabinactive" align="center" nowrap id="tab_5" name="tab_5" language="javascript" onclick="sel_tab('5')">�ֺq�Ǧެ���</td>
<td width="1" nowrap></td>
<td width="100" class="tabinactive" align="center" nowrap id="tab_6" name="tab_6" language="javascript" onclick="sel_tab('6')">²������</td>
<td width="1" nowrap></td>
<td width="100" class="tabinactive" align="center" nowrap id="tab_7" name="tab_7" language="javascript" onclick="sel_tab('7')">�\Ū�ޥ�����</td>
<td width="1" nowrap></td>

<td width="90%"></td>
</tr>
<tr height="1">
<td bgcolor="silver"></td>
<td bgcolor="white" id="tab_1b" name="tab_1b"></td>
<td bgcolor="silver"></td>
<td bgcolor="silver" id="tab_2b" name="tab_2b"></td>
<td bgcolor="silver"></td>
<td bgcolor="silver" id="tab_3b" name="tab_3b"></td>
<td bgcolor="silver"></td>
<td bgcolor="silver" id="tab_4b" name="tab_4b"></td>
<td bgcolor="silver"></td>
<td bgcolor="silver" id="tab_5b" name="tab_5b"></td>
<td bgcolor="silver"></td>
<td bgcolor="silver" id="tab_6b" name="tab_6b"></td>
<td bgcolor="silver"></td>
<td bgcolor="silver" id="tab_7b" name="tab_7b"></td>
<td bgcolor="silver"></td>

<td bgcolor="silver"></td>
<td bgcolor="silver"></td>
</tr>
</table>
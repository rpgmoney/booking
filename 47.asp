<!-- #INCLUDE virtual="/include/lib/conn/msconn.asp" -->
<%
function getguid()
	dim sql,rslib
	getguid=""
	sql="select newid() as guid"
	set rslib=server.createobject("adodb.recordset")
	rslib.open sql,msconn,adOpenStatic,adLockReadOnly
	if err.number=0 then
		if not rslib.eof then
			getguid=rslib("guid").value
		end if
	end if
	rslib.close
	set rslib=nothing
end function



sid = 99083
id = getguid()
name = "±i¥@©_"
response.Write("sid => " & sid & "<br />" )
response.Write("id => " & id & "<br />" )
response.Write("name => " & name & "<br />" )
Set rs = server.CreateObject("adodb.recordset")
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = msconn
sql = "INSERT INTO boo_profile(id, sid, name) VALUES('"& id &"','"& sid & "',  '" & name & "')"
response.Write(sql)
'response.End
cmd.CommandText = sql
cmd.Execute


%>
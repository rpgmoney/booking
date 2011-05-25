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

'scid=getguid()
response.write(getguid())
%>
<%
if session("sid")="" or session("st_status")=""  then
	response.redirect "login.asp"
end if

%>

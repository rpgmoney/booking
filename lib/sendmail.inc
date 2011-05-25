<%

function send_HTMLMail(username,mailfrom,subject,bodyText,filename)
	on Error resume next
	dim objCDOMail
	
	set objCDOMail=Server.CreateObject("CDO.Message")

	
	if Err.number <> 0 then
		Response.Write "Create Mail Object Error!"
		Response.End
	end if 
	objCDOMail.From=mailfrom
	objCDOMail.To=username
	objCDOMail.Subject=subject
	objCDOMail.HTMLBody=bodyText
	if filename<>"" then
		'objCDOMail.AttachFile Server.MapPath("file/" & filename)
		objCDOMail.AddAttachment Server.MapPath("file/" & filename)
	end if
	objCDOMail.Send
	if Err.number <> 0 then
		Response.Write "Send Mail Error!"
		Response.Write Err.description
		Response.End
	end if 
	Set objCDOMail=nothing
	on Error goto 0
end function

 

%>

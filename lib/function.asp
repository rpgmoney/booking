<%

function UpdateItemTime(vitem,vtid,vflag)
	'��s�w������,vflag=+�[�@,vflag=-��@
	'response.write "vitem=" & vitem & ":vtid=" & vtid & ":vflag=" &  vflag
	if  vflag="1" then

		if vitem="�f�y" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  optime_b=optime_b+1 where tid  in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="²��" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  pptime_b=pptime_b+1 where tid  in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="�ֺq" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  crkptime_b=crkptime_b+1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="�g�@" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  writetime_b=writetime_b+1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="�\Ū" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  readtime_b=readtime_b+1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		end if
	elseif vflag="2" then
		if vitem="�f�y" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  optime_b=optime_b-1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="²��" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  pptime_b=pptime_b-1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="�ֺq" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  crkptime_b=crkptime_b-1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="�g�@" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  writetime_b=writetime_b-1 where tid  in ('"&vtid&"')"
			msconn.Execute updatesql
		elseif  vitem="�\Ū" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  readtime_b=readtime_b-1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		end if

	end if
	'response.write "updatesql=" & updatesql

end function

function UpdateItemTime_C(vitem,vtid,vflag)
	'��s�w������,vflag=+�[�@,vflag=-��@
	'response.write "vitem=" & vitem & ":vtid=" & vtid & ":vflag=" &  vflag
	if  vflag="1" then

		if vitem="�f�y" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  optime_c=optime_c+1 where tid  in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="²��" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  pptime_c=pptime_c+1 where tid  in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="�ֺq" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  crkptime_c=crkptime_c+1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="�g�@" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  writetime_c=writetime_c+1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="�\Ū" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  readtime_c=readtime_c+1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		end if
	elseif vflag="2" then
		if vitem="�f�y" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  optime_c=optime_c-1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="²��" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  pptime_c=pptime_c-1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="�ֺq" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  crkptime_c=crkptime_c-1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="�g�@" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  writetime_c=writetime_c-1 where tid  in ('"&vtid&"')"
			msconn.Execute updatesql
		elseif  vitem="�\Ū" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  readtime_c=readtime_c-1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		end if

	end if
	'response.write "updatesql=" & updatesql

end function


%>
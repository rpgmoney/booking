<%

function UpdateItemTime(vitem,vtid,vflag)
	'更新預約次數,vflag=+加一,vflag=-減一
	'response.write "vitem=" & vitem & ":vtid=" & vtid & ":vflag=" &  vflag
	if  vflag="1" then

		if vitem="口語" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  optime_b=optime_b+1 where tid  in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="簡報" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  pptime_b=pptime_b+1 where tid  in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="詩歌" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  crkptime_b=crkptime_b+1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="寫作" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  writetime_b=writetime_b+1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="閱讀" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  readtime_b=readtime_b+1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		end if
	elseif vflag="2" then
		if vitem="口語" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  optime_b=optime_b-1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="簡報" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  pptime_b=pptime_b-1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="詩歌" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  crkptime_b=crkptime_b-1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="寫作" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  writetime_b=writetime_b-1 where tid  in ('"&vtid&"')"
			msconn.Execute updatesql
		elseif  vitem="閱讀" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  readtime_b=readtime_b-1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		end if

	end if
	'response.write "updatesql=" & updatesql

end function

function UpdateItemTime_C(vitem,vtid,vflag)
	'更新預約次數,vflag=+加一,vflag=-減一
	'response.write "vitem=" & vitem & ":vtid=" & vtid & ":vflag=" &  vflag
	if  vflag="1" then

		if vitem="口語" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  optime_c=optime_c+1 where tid  in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="簡報" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  pptime_c=pptime_c+1 where tid  in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="詩歌" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  crkptime_c=crkptime_c+1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="寫作" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  writetime_c=writetime_c+1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="閱讀" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  readtime_c=readtime_c+1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		end if
	elseif vflag="2" then
		if vitem="口語" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  optime_c=optime_c-1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="簡報" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  pptime_c=pptime_c-1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="詩歌" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  crkptime_c=crkptime_c-1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		elseif  vitem="寫作" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  writetime_c=writetime_c-1 where tid  in ('"&vtid&"')"
			msconn.Execute updatesql
		elseif  vitem="閱讀" and  vtid<>"" then
			updatesql = "update  boo_diagnosis set  readtime_c=readtime_c-1 where tid in ('"&vtid&"') "
			msconn.Execute updatesql
		end if

	end if
	'response.write "updatesql=" & updatesql

end function


%>
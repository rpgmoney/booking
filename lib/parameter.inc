<%

'�^�ˤj��X����h�����n�̳B����
par_score=240

'�B���Ҥ��^�E����Y�W�L�X�ѫh�ݭ��s�E�_

par_extinct_day =-45

'�Ǧ~�Ǵ�
par_yms = "992"



function YmsOption(StartNum,EndNum,chk)
	if chk<>"" then
		chk=Cstr(chk)
	end if	
	for i=StartNum to EndNum 
		for j=1 to 2
			if Cstr (i&j) = chk then
				YmsOption=YmsOption & "<option value=""" &  i&j & """  selected>" & i&j  & "</option>"
			else
				YmsOption=YmsOption & "<option value=""" & i&j  & """>" & i&j  & "</option>"
			end if
		Next
	Next
end function
%>
<%

'英檢大於幾分後則不須要依處方籤
par_score=240

'處方籤之回診日期若超過幾天則需重新診斷

par_extinct_day =-45

'學年學期
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
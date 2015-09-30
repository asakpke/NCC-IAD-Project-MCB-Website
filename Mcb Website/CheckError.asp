<%
	if Err.number <> 0 then
		'Response.Write "<hr>Error # " & Err.number & "<br>" & Err.description 
		select case Err.number 
			case -2147217873 'string value exits in db
			     '2147217873<BR>
			     '2147217873<BR>
				'Response.Write "This NIC # Already exists"
				Session("Msg") = "The Record Already exist."
			case -2147217900 'numeric value exits in db
				Session("Msg") = "The Record Already exist."
			'80040e2f
			'-2147217873
			case else
				Session("Msg") = "Unknown Error."
		end select
		'Response.Write "<hr>"
		'Session("Msg") =  Session("Msg") &  "<br>Following is a detail<br>" & "Error Number = " & Err.number & "<br>" & "Error Description = " & Err.description
		'&  "<br>" & 
		'-2147217900
	end if
%>
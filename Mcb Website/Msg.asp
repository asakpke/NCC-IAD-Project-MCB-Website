<%
	if Len( Session( "Msg" ) ) > 0 then
		%>
		<BODY background="b.jpg">
		<%
		Response.Write "<HR>"
		Response.Write "<B>Message = </B>" & Session("Msg")
		Response.Write "<HR>"
		Session( "Msg" ) = ""
		%>
		</BODY>
		<%
	end if
%>
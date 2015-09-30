<%
	'dim sName
	'sName = "Aamir"
	Response.Write "<HR>"
	Response.Write "<B>UserID = </B>" & Session("UserID") & "<BR>"
	Response.Write "<B>UserType = </B>" & Session("UserType") & "<BR>"
	Response.Write "<B>Msg = </B>" & Session("Msg")
	Response.Write "<HR>"
'if len(session("Msg")) > 0 then
	'Response.Write "<BR>isLogin = " & Session("isLogin")
	'Response.Write  "<BR><B>Message</B> = " &  session("Msg")
	session("Msg") = ""
'end if
%>
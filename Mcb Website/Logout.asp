<%
	Session("UserID") = 0
	Session("UserType") = 0
	Session("AccID") = 0
	Session("Amount") = 0
	Session("Msg") = ""
	Response.Redirect "default.asp"
%>
<%@ Language=VBScript %>
<%
if Session("UserID") = 0 then
	Session("Msg") = "You r not login"
	Response.Redirect "default.asp"
end if
%>
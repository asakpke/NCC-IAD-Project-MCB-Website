<%@ Language=VBScript %>
<%
if Session("UserType") <> 2 then
	Session("Msg") = "This page is only for Shop. You r not valid user for this page"
	Response.Redirect "default.asp"
end if
%>
<%@ Language=VBScript %>
<%
if Session("UserType") <> 4 then
	Session("Msg") = "This page is only for Administrators. You r not valid user for this page"
	Response.Redirect "default.asp"
end if
%>
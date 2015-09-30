<%@ Language=VBScript %>
<%
if Session("UserType") <> 1 then
	Session("Msg") = "This page is only for Account holders. You r not valid user for this page"
	Response.Redirect "default.asp"
end if
%>
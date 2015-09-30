<%@ Language=VBScript %>
<%
if Session("UserType") <> 3 then
	Session("Msg") = "This page is only for Billing Company. You r not valid user for his page"
	Response.Redirect "default.asp"
end if
%>
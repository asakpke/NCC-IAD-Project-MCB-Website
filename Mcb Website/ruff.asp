<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<P>
<%
	'Response.Write Request.Form.Count
	Response.Write "<br>Request.Form(""text1"") = " & Request.Form("text1")
		Response.Write "<br>Request.Form(""text2"") = " & Request.Form("text2")

	'Response.Write "<br>Request.Form(2) = " & Request.Form(2)
	Response.Write "<br>Request.TotalBytes = " & Request.TotalBytes 
%>
</P>
</BODY>
</HTML>

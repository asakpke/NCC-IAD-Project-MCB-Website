<%@ Language=VBScript %>
<!-- #include file="ADOvbs.inc" -->
<%
	'Option Explicit
	set oRs = server.CreateObject("ADODB.recordset")	
	dim sSQL
	'sSQL="SELECT " _
		'& "User.LoginName FROM UserType INNER JOIN [User] ON UserType.ID = " _
		'& "User.UserTypeID ORDER BY User.ID;"
		
	oRs.Open "User", "DSN=dsnMCB",adOpenDynamic ,adLockPessimistic 

	Response.Write "<hr><h1>Password Changing</h1>"
	Response.Write "<br><TABLE BORDER=1>"
	Response.Write "<TR><TH>Old Pwd</TH><TH>New Pwd</TH></TR><TR><TD>"
	While Not oRs.EOF
	'	Response.Write oRs("Type")& "&nbsp;</TD><TD>"	
		Dim Obj
		Set Obj = CreateObject("pjtMyDll6.clsMyDll6")
		Dim Str
		Str = oRs("Password")
		Response.Write Str & "</TD><TD>"
		Str = Obj.Encript(CStr(Str))
		oRs("Password") = CStr(Str)
		Ors.Update 
		Response.Write Str & "</TD></TR><TR><TD>"
		oRs.MoveNext 
	Wend
	Response.Write "</TD></TR></TABLE>"
	Response.Write "<hr><TABLE BORDER=1>"
	Response.Write "<TR><TH>User Type</TH><TH>ID</TH><TH>Name</TH><TH>" _
		& "Address</TH><TH>Phone</TH><TH>Loin Name</TH><TH>Password</TH><TR><TD>"
	oRs.MoveFirst 
	Response.Write oRs.GetString(,,"</TD><TD>","</TD></TR><TR><TD>","&nbsp;")
	Response.Write "</TD></TR></TABLE>"

	oRs.Close
%>
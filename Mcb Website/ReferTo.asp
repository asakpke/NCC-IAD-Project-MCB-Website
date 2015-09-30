<!-- #include file="AuthenticateAccHolder.asp" -->
<!-- #include file="ADOvbs.inc" -->
<%
on error resume next
'Response.Write Request.Form("txtName")
	'set oRsMax = server.CreateObject("ADODB.recordset")
	'strSQL = "SELECT MAX([ID]) As MaxID FROM Customer;"
	'Response.Write "H1"
	'oRsMax.Open strSQL, "DSN=dsnAsjad"
	'Response.Write oRsMax("MaxID")
	'Response.Write "H2"
	dim oRs
	set oRs = server.CreateObject("ADODB.recordset")
	'Response.Write "H1"
	oRs.Open "ReferTo","DSN=dsnMCB",adOpenDynamic ,adLockPessimistic   
	'Response.Write "H2"
	'Response.Write "H"
	'Response.Write ors("name")
	oRs.AddNew 
	oRs("UserID") = Session("UserID")
	'oRs("ID")= oRsMax("MaxID") + 1
	oRs("Name")= request.Form("txtName")
	oRs("NIC")= Request.Form("txtNIC")
	oRs("Date")= Date
	'oRs("FName")= Request.Form("txtFName")
	'oRs("LoginName")= Request.Form("txtLoginName")
	'oRs("Password")= Request.Form("txtPwd")
	'Response.Write "<SCRIPT language=VBScript>Msgbox 1</SCRIPT>"
	'Response.Write 1
	oRs.Update
	%>
	<!-- #include file="CheckError.asp" -->
	<%
	if len(Session("Msg")) > 0 then
		'Session("Msg") = Session("Msg") & " i.e NIC #
		Response.Redirect "frmReferTo.asp"
	else
		%>
		<BODY background="b.jpg">
		<%
		Response.Write "<HR>New Record is Added<HR>"
	end if
	oRs.Close 
	%>

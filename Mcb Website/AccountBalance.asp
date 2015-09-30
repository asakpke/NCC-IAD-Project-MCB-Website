<!-- #include file="AuthenticateUser.asp" -->
<!-- #include file="ADOvbs.inc" -->
<!-- # include file="NavigatingAccHolder.htm" -->
<body background="b.jpg">
<%
	dim oCon
	set ocon = server.CreateObject("ADODB.Connection")
	oCon.Open "DSN=dsnMCB"
	
	Dim oCmd
	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = oCon
	oCmd.CommandText = "qFindBalance"
	oCmd.CommandType = adCmdStoredProc
	
	dim oParm
	set oParm = server.CreateObject("ADODB.Parameter")
	oParm.Name ="AID"
	oParm.Type=adInteger
	oParm.Direction = adParamInput
	'oParm.Value = 103
	oCmd.Parameters.Append oParm	
	dim oRsBal
	set oRsBal = server.CreateObject("ADODB.recordset")	

	dim oRsAcc
	set oRsAcc = server.CreateObject("ADODB.recordset")	
	dim sSql
	sSql="select * from Account where UserID=" & Session("UserID")
	oRsAcc.Open sSql,oCon
	'Response.Write "<BR>1. EOF = " & oRsAcc.EOF
	'Response.Write "<BR>1. BOF = " & oRsAcc.BOF
	
	'Remove this because each customer must have 1 bank account
	if oRsAcc.EOF or oRsAcc.BOF then
		Response.Write "No Account" 
	end if
	%>
	<font face="arial" color="green"><h2>
	Account Holder &rarr; <FONT size=7>A</FONT>ccount Balance</h2></font>
	<hr>
	<%
	'Response.Write "<H1>Balance of All Account</H1>"
	Response.Write "<TABLE BORDER=1>"
	Response.Write "<TR><TH>Account #</TH><TH>Balance</TH></TR><TR><TD>"
	while not oRsAcc.EOF 
		oParm.Value = oRsAcc("ID")
		set oRsBal = oCmd.Execute
		'esponse.Write "<BR>1. EOF = " & oRsBal.EOF
		'esponse.Write "<BR>1. BOF = " & oRsBal.BOF   
		if not oRsBal.EOF then ' or not oRsBal.BOF  
			'Response.Write "<BR>Session('UserID') = " & Session("UserID") & "<BR>"
			'Response.Write "<HR>Account # " & oRsAcc("ID")
			'Response.Write "<TABLE BORDER=1>"
			'Response.Write "<TR><TH>Balance</TH></TR><TR><TD>"
			Response.Write oRsAcc("ID") & "</TD><TD>"
			Response.Write oRsBal.GetString(,,"</TD></TR>", "<TR><TD>","&nbsp;")
			''Response.Write oRs.GetString(,,"</TD><TD>","</TD></TR><TR><TD>","&nbsp;")
			'Response.Write "</TD></TR></TABLE>"
	
	'	while not oRsBal.EOF
	'		Response.Write oRsBal("ID")& "</TD><TD>"
	'		'Response.Write oRs("Balance") & "</TD><TD>"
	'		Response.Write oRsBal("Balance") & "</TD></TR><TR><TD>"
	'		oRsBal.MoveNext 
	'	wend
		'Response.Write "</TD></TR></TABLE>"
		
			oRsAcc.MoveNext 
		end if
		oRsBal.Close 
	wend 
	Response.Write "</TD></TR></TABLE>"
	oRsAcc.Close 
	oCon.Close 
%>
</body>
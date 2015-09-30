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
	
	dim oRsTranCharges
	set oRsTranCharges = server.CreateObject("ADODB.recordset")
	oRsTranCharges.Open "Settings", oCon
	dim nTranCharges
	nTranCharges = oRsTranCharges("TransactionCharges")
	oRsTranCharges.Close
	
	dim nAmount
	nAmount = 0
	if Session("AccID") <> 0 then
		oParm.Value = Session("AccID")
		nAmount = Session("Amount")
	end if
	
	nAmount = nAmount + nTranCharges
		
	Session("AccID") = 0
	Session("Amount") = 0
	
	oCmd.Parameters.Append oParm	
	
	dim oRsBal
	set oRsBal = server.CreateObject("ADODB.recordset")			
	set oRsBal = oCmd.Execute
	
	if CCur(oRsBal("Balance")) >= CCur(nAmount)  then 
		Session( "Msg" ) = ""
	else
		Session( "Msg" ) = "There is only Rs " & CCur( oRsBal("Balance")) _
				& " in your account balance while given amount is Rs " _
				&  CCur( nAmount ) & " included with Transaction charges = " & nTranCharges
	end if	
	oRsBal.Close 
	oCon.Close 
%>
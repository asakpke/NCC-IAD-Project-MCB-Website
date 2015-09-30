<!-- #include file="AuthenticateUser.asp" -->
<!-- #include file="ADOvbs.inc" -->
<%
	if CCur(Request.Form("txtAmount")) < 1 then
		Session( "Msg" ) = "Amount should not be < 1"
		Response.Redirect "frmTransfer.asp"
	end if
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
	oParm.Value = Request.Form("lstAcc")
	oCmd.Parameters.Append oParm	
	
	dim oRsBal
	set oRsBal = server.CreateObject("ADODB.recordset")			
	set oRsBal = oCmd.Execute
	'sdfsdfsdfsdfsd
	'Response.Write oRsBal("Balance") & " >= " & Request.Form("txtAmount") & " = "
	'Response.Write ( cint( oRsBal("Balance") ) >= cint( Request.Form("txtAmount") ) )
	'Response.End 
	'dfsdfsdfsdfsdfsf
	dim oRsTranCharges
	set oRsTranCharges = server.CreateObject("ADODB.recordset")
	oRsTranCharges.Open "Settings", oCon
	dim nTranCharges
	nTranCharges = oRsTranCharges("TransactionCharges")
	oRsTranCharges.Close 
	
	if CCur( oRsBal("Balance")) >= (CCur( Request.Form("txtAmount") )+ nTranCharges) then 
		dim oRsTran
		
		set oRsTran = server.CreateObject("ADODB.recordset")
		oRsTran.Open "Transaction", oCon, adOpenDynamic ,adLockPessimistic
		orsTran.AddNew   
		oRsTran("FromAccountID")= Request.Form("lstAcc") 'Request.Form("txtFromAccID")
		oRsTran("ToAccountID")=Request.Form("txtToAccID")	
		oRsTran("TranscationTypeID")= 1 '--> transfer
		oRsTran("Date")= date
		oRsTran("Amount")= Request.Form("txtAmount")	
on error resume next
		oRsTran.Update  
		%>
		<!-- #include file="CheckError.asp" -->
		<%
		oRsTran.Close
		if len(Session("Msg")) > 0 then
			Session("Msg") = "Transfer to account # does not exists"
			Response.Redirect "frmTransfer.asp"
		end if
		
		'for Tran Charges
		oRsTran.Open "Transaction", oCon, adOpenDynamic ,adLockPessimistic
		orsTran.AddNew   
		oRsTran("FromAccountID")= Request.Form("lstAcc") 'Request.Form("txtFromAccID")
		oRsTran("ToAccountID")= 38 ' a bank deduction account #	
		oRsTran("TranscationTypeID")= 5 '--> bank deduction
		oRsTran("Date")= date
		oRsTran("Amount")= nTranCharges	
		oRsTran.Update  
		%>
		<!-- #include file="CheckError.asp" -->
		<%
		oRsTran.Close
		if len(Session("Msg")) > 0 then
			'Session("Msg") = ""
			Response.Redirect "frmTransfer.asp"
		end if
		%>
		<BODY background="b.jpg">
		<%
		Response.Write "<HR>Transcation is complete<HR>"
	else
		Session( "Msg" ) = "There is only Rs " & CCur( oRsBal("Balance")) _
			& " in balance while amount transfering is Rs " _
			&  CCur( Request.Form("txtAmount") ) & " + Transaction charges = " & nTranCharges
		Response.Redirect "frmTransfer.asp"
	end if	
	oRsBal.Close 
	oCon.Close 
%>

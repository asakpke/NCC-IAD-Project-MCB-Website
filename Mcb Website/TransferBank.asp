<!-- #include file="AuthenticateUser.asp" -->
<!-- #include file="ADOvbs.inc" -->
<%
	if Request.Form("lstAcc") = "" then
		session("Msg") = "Invalid Account #. You must select an Account # for money Transfer"
		Response.Redirect "frmTransferBank.asp"
	end if
	
	if Request.Form("lstBank") = "" then
		session("Msg") = "Invalid Bank. You must select a Bank for money Transfer"
		Response.Redirect "frmTransferBank.asp"
	end if
	
	if Request.Form("txtToAccID") = "" then
		session("Msg") = "Invalid Account #"
		Response.Redirect "frmTransferBank.asp"
	end if
	
	if Request.Form("txtAmount") = "" then
		session("Msg") = "Invalid Amount"
		Response.Redirect "frmTransferBank.asp"
	end if
%>

<%
	if CCur(Request.Form("txtAmount")) < 1 then
		Session( "Msg" ) = "Amount should not be < 1"
		Response.Redirect "frmTransferBank.asp"
	end if
	
	'ssssssssssssssssssssss
	Session("AccID") = Request.Form("lstAcc")
	Session("Amount") = Request.Form("txtAmount")
	%>
	<!-- #include file="ValidateBalance.asp" -->
	<%
	if Len( Session( "Msg" ) ) > 0 then
		Response.Redirect "frmTransferBank.asp"
	end if
	'eeeeeeeeeeeeeeeeeeeeee
	'dim oCon
	'set ocon = server.CreateObject("ADODB.Connection")
	oCon.Open "DSN=dsnMCB"
	oCon.BeginTrans
	
	's
	dim oRs
	set oRs = server.CreateObject("ADODB.recordset")
	sSQL = "SELECT * FROM Account WHERE UserID = " _
			& CLng(Request.Form("lstBank")) & ";"
	oRs.Open  sSQL, oCon
	
	dim nComAccID
	nBankAccID = oRs("ID")
	oRs.Close
	'e
		
	'dim oRs
	'set oRs = server.CreateObject("ADODB.recordset")
	oRs.Open "Transaction", oCon, adOpenDynamic ,adLockPessimistic 
	 
	
	ors.AddNew   
	oRs("FromAccountID")= Request.Form("lstAcc") 'Request.Form("txtFromAccID")
	oRs("ToAccountID")= nBankAccID 'Request.Form("txtToAccID")	
	oRs("TranscationTypeID")= 4 '--> transfer to other bank
	oRs("Date")= date
	oRs("Amount")= Request.Form("txtAmount")	
	oRs.Update
	
	oRs.MoveLast 
	dim nTranID
	nTranID = oRs("ID")  
	
	oRs.Close
	
	'for Tran Charges
		oRs.Open "Transaction", oCon, adOpenDynamic ,adLockPessimistic
		ors.AddNew   
		oRs("FromAccountID")= Request.Form("lstAcc") 'Request.Form("txtFromAccID")
		oRs("ToAccountID")= 38 ' a bank deduction account #	
		oRs("TranscationTypeID")= 5 '--> bank deduction
		oRs("Date")= date
		oRs("Amount")= nTranCharges	
		oRs.Update  
		%>
		<!-- #include file="CheckError.asp" -->
		<%
		oRs.Close
		if len(Session("Msg")) > 0 then
			'Session("Msg") = ""
			Response.Redirect "frmTransferBank.asp"
		end if
	'end- for Tran Charges
	
	's
	'Dim oCmd
	'Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = oCon
	oCmd.CommandText = "INSERT INTO TransferBank(TranID, BankID, AccID) " & _
			"VALUES(" & nTranID & "," & Request.Form("lstBank") & "," &  Request.Form("txtToAccID") & ");"
	'oCmd.CommandText = "INSERT INTO Billing (TransactionID, BillNumber) " & _
								'"VALUES(52,999);"
	oCmd.CommandType = adCmdText
	'ocmd.Prepared = false '=true is compiled ver.
	oCmd.Execute	
	'e
	
	oCon.CommitTrans 
	
	oCon.Close
	%>
	<!-- # include file="NavigatingAccHolder.htm" -->
	<BODY background=b.jpg>
	<%
	Response.Write "<HR>Transcation is complete<HR>"
%>

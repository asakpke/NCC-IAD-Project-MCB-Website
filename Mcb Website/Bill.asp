<!-- #include file="AuthenticateAccHolder.asp" -->
<!-- #include file="ADOvbs.inc" -->
<%
	if CCur(Request.Form("txtAmount")) < 1 then ' CCur = Change Currency
		Session( "Msg" ) = "Amount should not be < 1"
		Response.Redirect "frmBill.asp"
	end if
	
	Session("AccID") = Request.Form("lstFromAccID")
	Session("Amount") = Request.Form("txtAmount")
%>
<!-- #include file="ValidateBalance.asp" -->
<%
	if Len( Session( "Msg" ) ) > 0 then
		Response.Redirect "frmBill.asp"
	end if

	dim Cn
	set Cn = server.CreateObject("ADODB.Connection")
	Cn.Open "DSN=dsnMCB"
	Cn.BeginTrans 
	
	dim oRsCom
	set oRsCom = server.CreateObject("ADODB.recordset")
	sSQL = "SELECT * FROM Account WHERE UserID = " _
			& Request.Form("lstComID") & ";"
	oRsCom.Open  sSQL, Cn
	
	dim nComAccID
	nComAccID = oRsCom("ID")
	oRsCom.Close 
	
	dim ORsTran
	set oRsTran = server.CreateObject("ADODB.recordset")
	oRsTran.Open "Transaction", Cn,adOpenDynamic ,adLockPessimistic 
	
	oRsTran.AddNew   
	oRsTran("FromAccountID")= Request.Form("lstFromAccID")
	oRsTran("ToAccountID")= nComAccID	
	oRsTran("TranscationTypeID")= 2 '"--> bill payment
	oRsTran("Date")= date
	oRsTran("Amount")= Request.Form("txtAmount")	
	oRsTran.Update
	oRsTran.MoveLast 
	dim nTranID
	nTranID = oRsTran("ID")
	ORsTran.Close
	
on error resume next
	'for Tran Charges
		oRsTran.Open "Transaction", Cn, adOpenDynamic ,adLockPessimistic
		orsTran.AddNew   
		oRsTran("FromAccountID")= Request.Form("lstFromAccID") 'Request.Form("txtFromAccID")
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
			Session("Msg") = Session("Msg") & " ."
			Response.Redirect "frmBill.asp"
		end if
	'end- for Tran Charges
	'''''''''''''''''''''''''''''''''
	'Dim oCmd
	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Cn
	oCmd.CommandText = "INSERT INTO Billing (TransactionID, BillNumber) " & _
			"VALUES(" & nTranID & "," &  clng(Request.Form("txtBillNo")) & ");"
	'oCmd.CommandText = "INSERT INTO Billing (TransactionID, BillNumber) " & _
								'"VALUES(52,999);"
	oCmd.CommandType = adCmdText
	
	'dim oParm
	'set oparm = server.CreateObject("ADODB.Parameter")
	'set oParm = oCmd.CreateParameter("TranID",adInteger ,adParamInput,,nTranID)
	'ocmd.Parameters.Append oParm
	'set oParm = oCmd.CreateParameter("BillNo",adInteger ,adParamInput,, _
	'									cint(Request.Form("txtBillNo")))
	'ocmd.Parameters.Append oParm
	'ocmd.Prepared = false '=true is compiled ver.
	oCmd.Execute
	'start-err
	%>
	<!-- #include file="CheckError.asp" -->
	<%
	if len(Session("Msg")) > 0 then
		Session("Msg") = Session("Msg") & " i.e Bill #"
		Response.Redirect "frmBill.asp"
	else
		Cn.CommitTrans 
		Cn.Close
		%>
		<BODY background="b.jpg">
		<%
		Response.Write "<HR><H2>Bill payment is completed</H2><HR>"
	end if
	'end-err
	'''''''''''''''''''''''''''''''
	'dim ORsBilling
	'set oRsBilling = server.CreateObject("ADODB.recordset")
	'oRsBilling.Open "Billing", Cn,adOpenDynamic ,adLockPessimistic    
	'oRsBilling.AddNew
	'oRsBilling("TransactionID")= nTranID
	'oRsBilling("BillNumber") = cint(Request.Form("txtBillNo"))
	'ORsBilling.Update 
	'ORsBilling.Close 	
%>

<!-- #include file="AuthenticateAccHolder.asp" -->
<!-- #include file="ADOvbs.inc" -->
<%
on error resume next
	set oCn = server.CreateObject("ADODB.Connection")
	set oRs = server.CreateObject("ADODB.recordset")

	oRs.Open "Account", "DSN=dsnMCB",adOpenDynamic ,adLockPessimistic 	
	oRs.AddNew 
	oRs("UserID")= Session("UserID")
	oRs("BranchID")= Request.Form("lstBID")
	oRs("AccountTypeID")=request.form("lstAccType")	
	oRs("AccountLoginName")=request.form("txtPC")	
	oRs("Password")=request.form("txtPwd")			
	oRs.Update
	'start-err
	%>
	<!-- #include file="CheckError.asp" -->
	<%
	if len(Session("Msg")) > 0 then
		'Session("Msg") = Session("Msg") & " i.e NIC #
		Response.Redirect "frmNewAccount.asp"
	end if
	'-end-err
	oRs.MoveLast
	dim nNewAccID
	nNewAccID = oRs("ID")
	oRs.Close
	
	oRs.Open "Deposit", "DSN=dsnMCB",adOpenDynamic ,adLockPessimistic 	
	oRs.AddNew
	oRs("AccountID")= nNewAccID
	oRs("BranchID")= 1 'Request.Form("txtBID")
	oRs("Date")= Date
	oRs("Amount")= 0 'request.form("txtPC")		
	oRs.Update
	'start-err
		%>
		<!-- #include file="CheckError.asp" -->
		<%
		if len(Session("Msg")) > 0 then
			'Session("Msg") = Session("Msg") & " i.e NIC #
			Response.Redirect "frmNewAccount.asp"
		end if
	'-end-err
	oRs.Close
	
	oRs.Open "Withdraw", "DSN=dsnMCB",adOpenDynamic ,adLockPessimistic 	
	oRs.AddNew
	oRs("AccountID")= nNewAccID
	oRs("BranchID")= 1 'Request.Form("txtBID")
	oRs("Date")= Date
	oRs("Amount")= 0 'request.form("txtPC")		
	oRs.Update
	oRs.Close		
	'start-err
		%>
		<!-- #include file="CheckError.asp" -->
		<%
		if len(Session("Msg")) > 0 then
			'Session("Msg") = Session("Msg") & " i.e NIC #
			Response.Redirect "frmNewAccount.asp"
		end if
	'-end-err
	oRs.Open "Transaction", "DSN=dsnMCB",adOpenDynamic ,adLockPessimistic 	
	oRs.AddNew
	oRs("FromAccountID")= nNewAccID
	oRs("ToAccountID")= nNewAccID 'Request.Form("txtBID")
	oRs("TranscationTypeID")= 1 'transfer
	oRs("Date")= Date
	oRs("Amount")= 0 'request.form("txtPC")		
	oRs.Update
	'start-err
		%>
		<!-- #include file="CheckError.asp" -->
		<%
		if len(Session("Msg")) > 0 then
			'Session("Msg") = Session("Msg") & " i.e NIC #
			Response.Redirect "frmNewAccount.asp"
		end if
	'-end-err
	oRs.Close	
	%>
	<BODY background="b.jpg">
	<%
	Response.Write "New Account # " & nNewAccID & " is Created"
%>
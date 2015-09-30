<!-- #include file="AuthenticateAdmin.asp" -->
<!-- #include file="ADOvbs.inc" -->
<%
	if CCur(Request.Form("txtAmount")) < 1 then
		Session( "Msg" ) = "Amount should not be < 1"
		Response.Redirect "AdminFrmDeposit.asp"
	end if
	dim oCon
	set ocon = server.CreateObject("ADODB.Connection")
	oCon.Open "DSN=dsnMCB"
	
	dim oRsEmpBID
	set oRsEmpBID = server.CreateObject("ADODB.recordset")
	sSQL = "SELECT * FROM Employee WHERE UserID = " _
			& Session("UserID")
	oRsEmpBID.Open  sSQL, oCon
	
	dim nBID
	nBID = oRsEmpBID("BranchID")
	oRsEmpBID.Close 
	
		dim oRsTran	
		set oRsTran = server.CreateObject("ADODB.recordset")
		oRsTran.Open "Deposit", oCon, adOpenDynamic ,adLockPessimistic
		orsTran.AddNew   
		oRsTran("AccountID")= Request.Form("txtAccID") 'Request.Form("txtFromAccID")
		oRsTran("BranchID")= nBID
		oRsTran("EmpID")= Session("UserID")
		oRsTran("Date")= date
		oRsTran("Amount")= Request.Form("txtAmount")	
on error resume next
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

	oCon.Close 
%>



<!-- #include file="AuthenticateAdmin.asp" -->
<!-- #include file="ADOvbs.inc" -->
<%
	if CCur(Request.Form("txtAmount")) < 1 then
		Session( "Msg" ) = "Amount should not be < 1"
		Response.Redirect "AdminFrmWithdraw.asp"
	end if
	
	dim oCon
	set ocon = server.CreateObject("ADODB.Connection")
	oCon.Open "DSN=dsnMCB"
	
	dim oRsState	
	set oRsState = server.CreateObject("ADODB.recordset")
	Dim sSQL
	sSQL = "SELECT * FROM	Account WHERE ID = " & Request.Form("txtAccID")
	oRsState.Open sSQL, oCon
	dim nUserID
	nUserID = 0
	if Not oRsState.EOF then
		nUserID = oRsState("UserID")
		oRsState.Close
		sSQL = "SELECT * FROM User WHERE ID = " & nUserID
		oRsState.Open sSQL, oCon
		dim nState
		nState = 0
		if Not oRsState.EOF then
			nState = oRsState("UserStateID")
			if nState <> 3 then
				Session("Msg") = "Your account is not set for offline"
				Response.Redirect "AdminFrmWithdraw.asp"
			end if
			oRsState.Close
		else
			oRsState.Close
		end if
	else
		oRsState.Close
	end if
	 
	's-balancd
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
	oParm.Value = Request.Form("txtAccID")
	oCmd.Parameters.Append oParm	
	
	dim oRsBal
	set oRsBal = server.CreateObject("ADODB.recordset")			
	set oRsBal = oCmd.Execute
	
	if CCur( oRsBal("Balance")) >= CCur( Request.Form("txtAmount")) then 
		's-dep
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
		oRsTran.Open "Withdraw", oCon, adOpenDynamic ,adLockPessimistic
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
			Response.Redirect "AdminFrmWithdraw.asp"
		end if
		'e-dep
		%>
		<BODY background="b.jpg">
		<%
		Response.Write "<HR>Transcation is complete<HR>"
	else
		Session( "Msg" ) = "There is only Rs " & CCur( oRsBal("Balance")) _
			& " in balance while amount transfering is Rs " _
			&  CCur( Request.Form("txtAmount") )
		Response.Redirect "AdminFrmWithdraw.asp"
	end if	
	oRsBal.Close 
	oCon.Close 
	'e-balance
	oCon.Close 
%>



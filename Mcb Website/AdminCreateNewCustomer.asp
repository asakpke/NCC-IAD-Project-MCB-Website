<!-- #include file="AuthenticateAdmin.asp" -->
<!-- #include file="ADOvbs.inc" -->
<%
'<!-- #include file="AuthenticateAccHolder.asp" -->
'Response.Write Request.Form("txtName")
	if Request.Form("txtName") = "" Or Request.Form("txtNIC") = "" _ 
	Or Request.Form("txtLoginName") = "" Or Request.Form("txtPwd") = "" _
	Or Request.Form("lstBranchID") = "" Or Request.Form("lstAccTypeID") = "" _
	Or Request.Form("txtPinCode") = "" Or Request.Form("txtAccPwd") = "" then
		session("Msg") = "You must fill the fields with bold lables"
		Response.Redirect "frmNewCustomer.asp"
	end if
	
	dim Cn
	set Cn = server.CreateObject("ADODB.Connection")
	Cn.Open "DSN=dsnMCB"
	Cn.BeginTrans 
	
		dim oRs
		set oRs = server.CreateObject("ADODB.recordset")
		oRs.Open "User", Cn,adOpenDynamic ,adLockPessimistic 
		oRs.AddNew
		oRs("UserTypeID") = 1 'for Account Holder
		oRS("UserStateID") = 3 ' for both off/on line
		'oRs("ID")= oRsMax("MaxID") + 1 'It is Auto# now
		oRs("Name")= request.Form("txtName")
		's
		if Request.Form("txtAddr") <> "" then
			oRs("Address")= Request.Form("txtAddr")
		end if
		'e
		if Request.Form("txtPhone") <> "" then
			oRs("Phone")= Request.Form("txtPhone")
		end if
		
		oRs("LoginName")= Request.Form("txtLoginName")
		
		Dim Obj
		Set Obj = CreateObject("pjtMyDll6.clsMyDll6")
		Dim sStr
		sStr = Request.Form("txtPwd")
		sStr = Obj.Encript(CStr(sStr))
		
		oRs("Password")= CStr(sStr)
on error resume next
		oRs.Update
		'start-err
		%>
		<!-- #include file="CheckError.asp" -->
		<%
		if len(Session("Msg")) > 0 then
			'Session("Msg") = Session("Msg") & " i.e NIC #
			Response.Redirect "AdminFrmCreateNewCustomer.asp"
		end if
	'-end-err
		oRs.MoveLast 
		dim nUserID
		nUserID = oRs("ID")
		'Response.Write oRs("ID")& " = "  &nUserID
		'Response.End 
		oRs.Close
		'sssssssssssssssssssssssss
		'dim oRs
		'set oRs = server.CreateObject("ADODB.recordset")
		oRs.Open "AccountHolder", Cn,adOpenDynamic ,adLockPessimistic 
		oRs.AddNew
		oRs("ID") =nUserID
		oRs("NIC")= Request.Form("txtNIC") 
		if Request.Form("txtFName") <> "" then
			oRs("FatherName")= Request.Form("txtFName")
		end if

		oRs.Update
		'start-err
		%>
		<!-- #include file="CheckError.asp" -->
		<%
'on error resume
		if len(Session("Msg")) > 0 then
			'Session("Msg") = Session("Msg") & " i.e NIC #
			Response.Redirect "AdminFrmCreateNewCustomer.asp"
			'Response.Redirect "frmNewCustomer.asp"
		end if
	'-end-err
		oRs.Close 
		'oRs("ID")= oRsMax("MaxID") + 1 'It is Auto# now
		'oRs("Name")= request.Form("txtName")
		's
		'*************************
		'dim oRsAcc
		'set oRsAcc = server.CreateObject("ADODB.recordset")
		oRs.Open "Account", Cn,adOpenDynamic ,adLockPessimistic 
		oRs.AddNew 
		oRs("UserID") = nUserID
		oRs("BranchID")= request.Form("lstBranchID")
		oRs("AccountTypeID")= Request.Form("lstAccTypeID")
		oRs("AccountLoginName")= Request.Form("txtPinCode")
		oRs("Password")= Request.Form("txtAccPwd")
		oRs.Update
		'start-err
		%>
		<!-- #include file="CheckError.asp" -->
		<%
		if len(Session("Msg")) > 0 then
			'Session("Msg") = Session("Msg") & " i.e NIC #
			'Response.Redirect "frmNewCustomer.asp"
			Response.Redirect "AdminFrmCreateNewCustomer.asp"
		end if
	'-end-err
		oRs.MoveLast 
		dim nAccID		
		nNewAccID = oRs("ID")
		oRs.Close
		
		'*****************************************************
		oRs.Open "Deposit", Cn,adOpenDynamic ,adLockPessimistic 	
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
			Response.Redirect "AdminFrmCreateNewCustomer.asp"
			'Response.Redirect "frmNewCustomer.asp"
		end if
	'-end-err
		oRs.Close
	
		oRs.Open "Withdraw", Cn,adOpenDynamic ,adLockPessimistic 	
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
			Response.Redirect "AdminFrmCreateNewCustomer.asp"
			'Response.Redirect "frmNewCustomer.asp"
		end if
	'-end-err
		oRs.Close		
	
		oRs.Open "Transaction", Cn,adOpenDynamic ,adLockPessimistic 	
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
			Response.Redirect "AdminFrmCreateNewCustomer.asp"
			'Response.Redirect "frmNewCustomer.asp"
		end if
	'-end-err
		oRs.Close	
		
		Cn.CommitTrans 
		'start-err
		%>
		<!-- #include file="CheckError.asp" -->
		<%
		if len(Session("Msg")) > 0 then
			'Session("Msg") = Session("Msg") & " i.e NIC #
			'Response.Redirect "frmNewCustomer.asp"
			Response.Redirect "AdminFrmCreateNewCustomer.asp"
		end if
	'-end-err
		Cn.Close 
		'*****************************************************
		'Response.Write "<BR>BID" &  request.Form("lstBranchID")
		'Response.Write "<BR>ATypeID" &  Request.Form("lstAccTypeID")
		%>
		<BODY background="b.jpg">
		<%
		Response.Write "<HR>" & request.Form("txtName")
		Response.Write "'s Record is Added<HR>"
		Response.Write "Your Account # is " & nNewAccID
%>

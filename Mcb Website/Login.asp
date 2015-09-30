<%@ Language=VBScript %>
<%
	if  Request.Form("txtName")  = "" or _
			Request.Form("txtPwd") = "" then
			
			Session("Msg") = "Fill all field"
			Response.Redirect "default.asp"
			Response.End
	end if

	set oRs = server.CreateObject("ADODB.recordset")
	strSQL = "SELECT * FROM User WHERE LoginName='" & Request.Form("txtName") & "'"
	oRs.Open strsql, "DSN=dsnMCB"
	
	if not oRs.EOF and not ors.BOF then
		if oRs("UserStateID") = 0 then
			session("Msg") = "Your are not allowed to use this account due to any some reason. Contact with MCB to solve this problem."
			Response.Redirect "default.asp"
		end if
		Dim Obj 'As pjtMyDll.clsMyDll
		Set Obj = CreateObject("pjtMyDll6.clsMyDll6")
		Dim StrDB
		StrDB = oRs("Password")
		StrDB = Obj.Decript(CStr(StrDB))
		Dim StrFrm
		StrFrm = Request.Form("txtPwd")
		StrFrm = Obj.Decript(CStr(StrFrm))
		if StrDB = StrFrm then
			'dim nCustTypeID 'check public
			Session("UserID") = oRs("ID")
			Session("UserType") = oRs("UserTypeID")
			'Response.Write "isLogin = " & Session("islogin")
		else
			session("Msg") = "Invalid Password"
			Session("UserID") = 0
			Session("UserType")=0
		end if
	else
			session("Msg") = "Invalid Login name"
			Session("UserID") = 0
			Session("UserType")=0
	end if
	
	oRs.Close
	
	select case Session("UserType")
		case 0 '
			'session("Msg") = "Invalid Login name/Password"
			Response.Redirect "default.asp"
		case 1 'AccountHolder
			Response.Redirect "AccountHolder.asp"
		case 2
			'Response.Write "Shop"
			Response.Redirect "Shop.asp"
		case 3
			Response.Redirect "BillingComp.asp"
			'Response.Write "Billing Company"
		case 4
			Response.Redirect "Admin.asp"
			'Response.Write "Admin"
		case 5
			Response.Write "Welcom Bank"
	end select
%>
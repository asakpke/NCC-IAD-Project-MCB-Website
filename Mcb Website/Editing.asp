<!-- #include file="AuthenticateAccHolder.asp" -->
<!-- #include file="ADOvbs.inc" -->
<%
	set oRs = server.CreateObject("ADODB.recordset")
	Dim s
	s="select * from User where id=" & Session("UserID")
	oRs.Open s, "DSN=dsnMCB",adOpenDynamic ,adLockPessimistic 
	if not ors.EOF or  not ors.BOF then 
		dim isUpdate
		isUpdate = false
		if Request.Form("txtPwd") <> "" then
			Dim Obj 'As pjtMyDll.clsMyDll
			Set Obj = CreateObject("pjtMyDll6.clsMyDll6")
			Dim St
			St = Request.Form("txtPwd")
			St = Obj.Encript(CStr(St))
			set Obj = Nothing
			oRs("Password")= St'Request.Form("txtpwd")
			isUpdate = true
		end if
		if Request.Form("txtAddr") <> "" then
			ors("Address")= Request.Form("txtAddr")
			isUpdate = true
		end if
		if Request.Form("txtphone") <> "" then
			ors("Phone")=request.form("txtPhone")	
			isUpdate = true
		End if
		
		'if isUpdate = true then
		ors.Update
		'end if
	end if
	
	if isUpdate = true then
		%>
		<!-- # include file="NavigatingAccHolder.htm" -->
		<BODY background="b.jpg">
		<%
		Response.Write "<HR>Record is update<HR>"
	else
		session("Msg") = "Record is not update Please enter values"
		Response.Redirect "frmEditing.asp"
	end if
	ors.Close 
	
%>


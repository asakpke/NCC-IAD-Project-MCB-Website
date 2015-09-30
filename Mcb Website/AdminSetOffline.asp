<!-- #include file="AuthenticateAdmin.asp" -->
<!-- #include file="ADOvbs.inc" -->
<%
	dim oCon
	set ocon = server.CreateObject("ADODB.Connection")
	oCon.Open "DSN=dsnMCB"
	
	dim oRsState	
	set oRsState = server.CreateObject("ADODB.recordset")
	Dim sSQL
	sSQL = "SELECT * FROM	User WHERE ID = " & Request.Form("txtAccID")
	oRsState.Open sSQL, oCon, adOpenDynamic ,adLockPessimistic
	
	if Not oRsState.EOF then
		oRsState("UserStateID") = Request.Form("txtState") 'for both
		
		's-err
on error resume next
		oRsState.Update
		%>
		<!-- #include file="CheckError.asp" -->
		<%
		oRsState.Close
		if len(Session("Msg")) > 0 then
			'Session("Msg") = ""
			Response.Redirect "AdminFrmSetOffline.asp"
		end if
		%>
		<BODY background="b.jpg">
		<%
		'e-dep
		Response.Write "<HR>Account state is set<HR>"
		'e-err
		oRsState.Close
	end if
%>
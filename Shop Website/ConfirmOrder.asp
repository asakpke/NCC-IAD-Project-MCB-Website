<%@ Language=VBScript %>
<!-- #include file="ADOvbs.inc" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
</HEAD>
<BODY>
<hr>
<h1>Welcome Back to Shop's Website</h1>
<hr>
<%
'47 for caller/sender 's URL							|MCB's Addr
'Response.Write "<hr>L = " & 47 & ", " & Request.ServerVariables.Item(47)  
if Request.ServerVariables(47) = "http://localhost/mcb/Purshase.asp" then
	'Response.Write "<br>txtConfirm = " & Request.Form("txtConfirm")
	'Response.Write "<br>txtOrderID = " & Request.Form("txtOrderID")
	'Response.Write "<br>txtTranID = " & Request.Form("txtTranID")
	's-edit & confirm order
	if Request.Form("txtConfirm") = 1 then
		dim oRs
		set oRs = server.CreateObject("ADODB.recordset")	
		dim strSQL
		strSQL = "SELECT BankTranID,State FROM OrderMaster WHERE ID=" _
							& CLng(Request.Form("txtOrderID"))
		oRs.Open strSQL, "dsnShop",adOpenDynamic ,adLockPessimistic 	
		if Not oRs.EOF then 
			oRs("BankTranID")= Request.Form("txtTranID")
			oRs("State")= True
			oRs.Update
		end if
		oRs.Close
		Response.Write "<br>Your Order is confirmed"
	'e-edit & confirm order
	else 
	'delete the order
		'DELETE OrderMaster.ID FROM OrderMaster WHERE (((OrderMaster.ID)=51));
		dim Cn
		set Cn = server.CreateObject("ADODB.Connection")
		Cn.Open "DSN=dsnShop"
		Cn.BeginTrans 
		Set oCmd = Server.CreateObject("ADODB.Command")
		oCmd.ActiveConnection = Cn
		oCmd.CommandText = "DELETE OrderMaster.ID FROM OrderMaster " _
			& "WHERE OrderMaster.ID = " &  CLng(Request.Form("txtOrderID"))
		oCmd.CommandType = adCmdText
		oCmd.Execute
		Cn.CommitTrans 
		Response.Write "<br>Your Order is <b>not</b> confirmed"
	end if
else
	Response.Write "<br>Wrong URL"
end if
%>

</BODY>
</HTML>
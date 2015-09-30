<%@ Language=VBScript %>
<!-- #include file="ADOvbs.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<META content="text/html; charset=unicode" http-equiv=Content-Type>
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub frmMCB_onsubmit
	if not IsNumeric(frmmcb.txtUserAccID.value) then
		msgbox "Enter Numeric Value"
		frmMCB.action =""
		window.history.back 
	end if
End Sub

-->
</SCRIPT>
</HEAD>
<BODY>
<%
'Option Explicit
dim TotalRs
TotalRs = 0
Response.Write "<h1>Order Detail</h1>"
's
	dim oCn
	set oCn = server.CreateObject("ADODB.Connection")
	oCn.Open "DSN=dsnShop"
	
	'start-Insert Into Order
	'dim strSQL
	'strSQL = "INSERT INTO OrderMaster ( CustName, [Date] ) " _
	'& "SELECT '" & Request.Form("txtName") & "', #" & date & "#;"
	'oCn.Execute strSQL 
	'end-Insert Into Order
	's-new order
	dim oRs
	set oRs = server.CreateObject("ADODB.recordset")	
	
	oRs.Open "OrderMaster", oCn,adOpenDynamic ,adLockPessimistic 	
	oRs.AddNew 
	oRs("CustName")= Request.Form("txtName")
	oRs("CustAddr")= Request.Form("txtAddr")
	oRs("Date")= Date
	oRs.Update
	
	oRs.MoveLast
	dim nOrderID
	nOrderID = oRs("ID")
	oRs.Close
	'e-new order
	
	Dim s
	s="SELECT * FROM Item"
	oRs.Open s, oCn
	
	's-order detail
	dim oRsDetail
	set oRsDetail = server.CreateObject("ADODB.recordset")	
	
	oRsDetail.Open "OrderDetail", oCn,adOpenDynamic ,adLockPessimistic 	
	'e-order detail
	''''''
	Response.Write "<TABLE BORDER=0>"
	Response.Write "<TR><TH>ID</TH><TH>Item</TH><TH>Price</TH></TR><TR><TD>"
	'Response.Write oRs.GetString(,,"</TD><TD>","</TD></TR><TR><TD>","&nbsp;")
	'Response.Write "</TD></TR></TABLE>"
	while not oRs.EOF
		'dim strID
		'strID = oRs("ID")
		'Response.Write  "<br>" & Request.Form(x)
		'Response.Write "<br>" & Request.Form(CStr(oRs("ID")))
		'Response.Write  Request.Form("2")
		if Request.Form(CStr(oRs("ID"))) = "on" then
			oRsDetail.AddNew 
			oRsDetail("OrderID") = nOrderID
			oRsDetail("ItemID") = oRs("ID")
			oRsDetail("Price") = oRs("Price")
			 
			Response.Write oRs("ID")& "</TD><TD>"
			Response.Write oRs("Name") & "</TD><TD>"
			Response.Write oRs("Price") & "</TD></TR><TR><TD>"
			TotalRs = TotalRs + oRs("Price")
		end if
		oRs.MoveNext 
	wend
	Response.Write "</TD></TR></TABLE>"
		'ors.Update
	'end if
	oRsDetail.Update 
	oRsDetail.Close 
	oRs.Close 
	'e
	oCn.Close 
	
Response.Write "<h3>Total Bill = " & TotalRs & "</h3>"
if TotalRs <= 0 then
	Response.End 
end if

	
'start - ruff del it after...
'Response.Write "<br>txtName = " & Request.Form("txtName")
'Response.Write "<br>txtAddr = " & Request.Form("txtAddr")
'end - ruff del it after...
%>
<form action="/mcb/frmPurshase.asp" id=frmMCB name=frmMCB method="post"><br>
<INPUT id=txtShopAccID name=txtShopAccID type=hidden value=35>
<INPUT id=txtTotal name=txtTotal type=hidden value=<%=TotalRs%>>
<INPUT id=txtShopURL name=txtShopURL type=hidden value="/shop/ConfirmOrder.asp">
<INPUT id=txtOrderID name=txtOrderID type=hidden value=<%=nOrderID%>>

Enter Your MCB Account Number <INPUT id=txtUserAccID name=txtUserAccID maxlength=9><br><br>
<INPUT id=submit1 name=submit1 type=submit value="Go to MCB">
<br></form>
</BODY>
</HTML>

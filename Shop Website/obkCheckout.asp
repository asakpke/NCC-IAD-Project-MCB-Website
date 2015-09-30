<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<META content="text/html; charset=unicode" http-equiv=Content-Type>
<META content="MSHTML 5.00.2614.3500" name=GENERATOR>
</HEAD>
<BODY>
<%
dim TotalRs
TotalRs = 0
Response.Write "<h1>Order Detail</h1>"
'Response.Write "<b><u>Item Name&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Rs</u></b>"
if Request.Form("chkItem1") = "on" then
	Response.Write "<br>Item1, Rs = " & Request.Form("Rs1")
	TotalRs = TotalRs + Request.Form("Rs1")
end if
if Request.Form("chkItem2") = "on" then
	Response.Write "<br>Item2, Rs = " & Request.Form("Rs2")
	TotalRs = TotalRs + Request.Form("Rs2")
end if
if Request.Form("chkItem3") = "on" then
	Response.Write "<br>Item3, Rs = " & Request.Form("Rs3")
	TotalRs = TotalRs + Request.Form("Rs3")
end if
if Request.Form("chkItem4") = "on" then
	Response.Write "<br>Item4, Rs = " & Request.Form("Rs4")
	TotalRs = TotalRs + Request.Form("Rs4")
end if
if Request.Form("chkItem5") = "on" then
	Response.Write "<br>Item5, Rs = " & Request.Form("Rs5")
	TotalRs = TotalRs + Request.Form("Rs5")
end if
Response.Write "<h3>Total Bill = " & TotalRs & "</h3>"
if TotalRs <= 0 then
	Response.End 
end if
dim nOrderID
nOrderID = 1
%>
<form action="http://aamir/aamir/frmPurshase.asp" id=frmMCB name=frmMCB method="post"><br>
txtShopAccID <INPUT name=txtShopAccID type=text value=35><br>
txtTotal <INPUT name=txtTotal type=text value=<%=TotalRs%>><br>
txtShopURL <INPUT name=txtShopURL type=text value="http://aamir/shop/ConfirmOrder.asp"><br>
txtOrderID <INPUT type="text" name=txtOrderID value=<%=nOrderID%>><br>
Enter Your MCB Account Number <INPUT id=txtUserAccID name=txtUserAccID><br>
<INPUT id=submit1 name=submit1 type=submit value=MCB>
<br></form>
</BODY>
</HTML>

<!-- #include file="AuthenticateShop.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY background="b.jpg">
<font face="arial" color="green"><h2>
Shop &rarr; <FONT size=7>T</FONT>ranscation list</h2>
</font>
<hr>
<%
	set oRs = server.CreateObject("ADODB.recordset")	
	Dim strSQL
	strSQL="SELECT * FROM Account WHERE UserID=" & Session("UserID")
	oRs.Open strSQL, "DSN=dsnMCB"
	''''''
	dim nShopAccID
	if Not oRs.EOF then
		nShopAccID = oRs("ID")
		oRs.Close 
		
		'strSQL = "SELECT * FROM Transaction WHERE ToAccountID=" _
								'& nShopAccID
		's-sql
'		sSQL="SELECT User.Name, Transaction.FromAccountID," _
'	& "Transaction.Date, "_
'	& "Transaction.Amount FROM User INNER JOIN Account "_
'	& "ON User.ID = Account.UserID) "_
'	& "INNER JOIN "_
'	& "ON "_
'	& "Account.ID = Transaction.FromAccountID) ON Account_1.ID = "_
'	& "Transaction.ToAccountID) ON User_1.ID = Account_1.UserID "_
'	& "WHERE (((Transaction.ToAccountID)= " & Request.Form("lstAccID") _
'	& ")) ORDER BY Transaction.Date;"
		'e-sql
's
strSQL = "SELECT Transaction.ID, User.Name, Transaction.FromAccountID, Transaction.Date, Transaction.Amount" _
& " FROM [User] INNER JOIN (Account INNER JOIN [Transaction] ON Account.ID = Transaction.FromAccountID) ON User.ID = Account.UserID" _
& " WHERE (Transaction.TranscationTypeID)=3 AND (Transaction.ToAccountID)=" & nShopAccID _
& " ORDER BY Transaction.Date DESC;"
'e
	oRs.Open strSQL, "DSN=dsnMCB"
		'Response.Write "<HR><H3>List of Transaction</H3><HR>"
		Response.Write "<TABLE BORDER=1>"
		Response.Write "<TR><TH>Transaction #</TH><TH>Name</TH><TH>Accounnt #</TH><TH>Date</TH><TH>Amount</TH></TR><TR><TD>"
		Response.Write oRs.GetString(,,"</TD><TD>","</TD></TR><TR><TD>","&nbsp;")
		Response.Write "</TD></TR></TABLE>"	
	end if
	
	oRs.Close 
	
%>

</BODY>
</HTML>

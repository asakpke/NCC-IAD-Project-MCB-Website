<!-- #include file="AuthenticateAccHolder.asp" -->
<%
	if Request.Form("lstAccID") = "" then
		session("Msg") = "Invalid Account #. You must select an Account # to see bank statment"
		Response.Redirect "frmSelectAccountBS.asp"
	end if
	'Response.Write "<br>acc#" & Request.Form("lstAccID")
	'Response.Write "<br>isnum(acc)" & IsNumeric(Request.Form("lstAccID"))
%>
<!-- #include file="ADOvbs.inc" -->
<body background="b.jpg">
<font face="arial" color="green"><h2>
Account Holder &rarr; <FONT size=7>B</FONT>ank Statment</h2></font>
<hr>
<!-- # include file="NavigatingAccHolder.htm" -->
<%	
	dim oCon
	set ocon = server.CreateObject("ADODB.Connection")
	oCon.Open "DSN=dsnMCB"
	
	'dim oRsAcc
	'set oRsAcc = server.CreateObject("ADODB.recordset")	
	'Dim sSQL
	'sSQL = "SELECT * FROM	Account WHERE UserID = " & Session("UserID")
	'oRsAcc.Open sSQL, oCon
	'Response.Write "1"
	'Response.Write oRsAcc(0)
	'Response.Write "2"
	
	'Response.Write "Session('UserID') = " & Session("UserID") & "<BR>"
	
	dim oRs
	set oRs = server.CreateObject("ADODB.recordset")	
	'Dim sSQL
	
	'while not oRsAcc.EOF 
		'Deposite
		Response.Write "<H1>Account # " & Request.Form("lstAccID") & "</H1>"
		Response.Write "Bank Statement between " & Request.Form("txtStartDt") _
			& " <B>To</B> " & Request.Form("txtEndDt") & " dates" & "<br>"
	
		
		'Transaction In Amount
		'sSQL="SELECT * FROM	Transaction WHERE ToAccountID=" & oRsAcc("ID")'& Session("CustID")
		'
		'sSQL="SELECT Transaction.FromAccountID, Transaction.ToAccountID, " _
'& "TranscationType.Type, Transaction.Date, Transaction.Amount " _
'& "FROM TranscationType INNER JOIN [Transaction] ON TranscationType.ID = " _
'& "Transaction.TranscationTypeID WHERE (((Transaction.ToAccountID)=" _
'& Request.Form("lstAccID") & "));"

'sssssssssss
	sSQL="SELECT User.Name, Transaction.FromAccountID, User_1.Name, " _
& "Transaction.ToAccountID, TranscationType.Type, Transaction.Date, "_
& "Transaction.Amount FROM [User] AS User_1 INNER JOIN (Account AS Account_1 "_
& "INNER JOIN (([User] INNER JOIN Account ON User.ID = Account.UserID) "_
& "INNER JOIN (TranscationType INNER JOIN [Transaction] ON "_
& "TranscationType.ID = Transaction.TranscationTypeID) ON "_
& "Account.ID = Transaction.FromAccountID) ON Account_1.ID = "_
& "Transaction.ToAccountID) ON User_1.ID = Account_1.UserID "_
& "WHERE Transaction.ToAccountID = " & Request.Form("lstAccID") & " " _
& "AND Transaction.Date >= #" & CDate(Request.Form("txtStartDt")) & "# " _
& "And Transaction.Date <= #" & CDate(Request.Form("txtEndDt")) & "# " _
& "ORDER BY Transaction.Date;"

'eeeeeeeeee
		'
		oRs.Open sSQL, oCon
		Response.Write "<BR><B>Online</B> Transaction (Debit/Credit)"
		Response.Write "<TABLE BORDER=1>"
		Response.Write "<TR><TH>From</TH><TH>Account #</TH><TH>To</TH><TH>Account #</TH><TH>Transcation Type</TH><TH>Date</TH><TH>Amount</TH><TH>Debit/Credit</TH></TR><TR><TD><FONT color=blue>"
	'	'dim curTotal
	'	curTotal = 0
		dim curTotal
		dim curBalance
		curTotal = 0
		curBalance = 0
		if not oRs.EOF then
			Response.Write oRs.GetString(,,"</TD><TD><FONT color=blue>","</TD><TD><FONT color=blue>Dr</TD></TR><TR><TD><FONT color=blue>","&nbsp;")
'	'		sSQL="SELECT User.Name, Transaction.FromAccountID, User_1.Name, " _
'	'		& "Transaction.ToAccountID, TranscationType.Type, Transaction.Date, "_
'	'		& "Transaction.Amount
'	'		Response.Write "<TR><TH>From</TH><TH>Account #</TH><TH>To</TH><TH>Account #</TH><TH>Transcation Type</TH><TH>Date</TH><TH>Amount</TH><TH>Debit/Credit</TH></TR><TR><TD>
	'	'	Response.Write oRs("Name") & "</TD><TD>" & oRs("FromAccountID") & "</TD><TD>" & oRs("User_1.Name") & "</TD></TR><TR><TD>"
	'	'	curTotal = curTotal + CCur(oRs("Amount"))
	'	'	oRs.MoveNext 		
			oRs.MoveFirst 
			while not oRs.EOF 
				curTotal = curTotal + CCur(oRs("Amount"))
				oRs.MoveNext 
			wend
			curBalance = curBalance + curTotal
		end if
		'Response.Write "</TD></TR></TABLE><HR>"	
		
		
		oRs.Close
		'ssssssssssssss
		'calc total amount
'	sSQL="SELECT Sum(Transaction.Amount) AS Total " _
'& "FROM [Transaction] GROUP BY Transaction.ToAccountID " _
'& "HAVING (((Transaction.ToAccountID)=" & Request.Form("lstAccID") & "));"
'oRs.Open sSQL, oCon
'	'if not oRs.EOF then
	Response.Write "&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD><TD><B>Total</B></TD><TD><B>"
	Response.Write curTotal'oRs("Total")
	Response.Write "</B></TD><TD>&nbsp;</TD></TR><TR><TD><FONT color=red>"
	'end if
		'Response.Write "</TD></TR></TABLE><HR>"	
'	oRs.Close
		'eeeeeeeeeeeee  
		
		'Transaction out Amount
		'sSQL="SELECT * FROM	Transaction WHERE FromAccountID=" & oRsAcc("ID")'& Session("CustID")
		'sSQL="SELECT Transaction.FromAccountID, Transaction.ToAccountID, " _
'& "TranscationType.Type, Transaction.Date, Transaction.Amount " _
'& "FROM TranscationType INNER JOIN [Transaction] ON TranscationType.ID = " _
'& "Transaction.TranscationTypeID WHERE (((Transaction.FromAccountID)=" _
'& Request.Form("lstAccID") & "));"

'sssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssss
	sSQL="SELECT User.Name, Transaction.FromAccountID, User_1.Name, " _
& "Transaction.ToAccountID, TranscationType.Type, Transaction.Date, "_
& "Transaction.Amount FROM [User] AS User_1 INNER JOIN (Account AS Account_1 "_
& "INNER JOIN (([User] INNER JOIN Account ON User.ID = Account.UserID) "_
& "INNER JOIN (TranscationType INNER JOIN [Transaction] ON "_
& "TranscationType.ID = Transaction.TranscationTypeID) ON "_
& "Account.ID = Transaction.FromAccountID) ON Account_1.ID = "_
& "Transaction.ToAccountID) ON User_1.ID = Account_1.UserID "_
& "WHERE Transaction.FromAccountID = " & Request.Form("lstAccID") & " " _
& "AND Transaction.Date >= #" & Request.Form("txtStartDt") & "# " _
& "And Transaction.Date <= #" & Request.Form("txtEndDt") & "# " _
& "ORDER BY Transaction.Date;"
'eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee

		oRs.Open sSQL, oCon
		'Response.Write "<BR>Transaction Out Amount"
		'Response.Write "<TABLE BORDER=1>"
		'Response.Write "<TR><TH>FromAccountID</TH><TH>ToAccountID</TH><TH>TranscationTypeID</TH><TH>ID</TH><TH>Date</TH><TH>Amount</TH></TR><TR><TD>"
		'Response.Write "---------=======--------</TD>/TR><TR><TD>"
		
		'Response.Write "&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD></TR><TR><TD>"
		'Response.Write "&nbsp;</TD></TR><TR><TD>"
		Response.Write "<FONT color=red>"
		if not oRs.EOF then
			Response.Write oRs.GetString(,,"</TD><TD><FONT color=red>","</TD><TD><FONT color=red>Cr</TD></TR><TR><TD><FONT color=red>","&nbsp;")
			
			curTotal = 0
			oRs.MoveFirst 
			while not oRs.EOF 
				curTotal = curTotal + CCur(oRs("Amount"))
				oRs.MoveNext 
			wend
			curBalance = curBalance - curTotal
		end if
		
		
		oRs.Close
		'sssssss
		'total cr
'	sSQL="SELECT Sum(Transaction.Amount) AS Total " _
'& "FROM [Transaction] GROUP BY Transaction.FromAccountID " _
'& "HAVING (((Transaction.FromAccountID)=" & Request.Form("lstAccID") & "));"
'		oRs.Open sSQL, oCon
	'if not oRs.EOF then
	Response.Write "&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD><TD><B>Total</B></TD><TD><B>"
	Response.Write curTotal'oRs("Total")
	Response.Write "</TD><TD>&nbsp;</TD></TR><TABLE>"
	'end if
		'Response.Write "</TD></TR></TABLE><HR>"	
'	oRs.Close
		'eeeeee
		'Response.Write "</TD></TR></TABLE>"	
		
		'Deposit
		'sSQL="SELECT * FROM	Deposit WHERE AccountID=" & oRsAcc("ID")'& Session("CustID")
		sSQL="SELECT Branch.Title, Deposit.Date, Deposit.Amount " _
& "FROM Branch INNER JOIN Deposit ON Branch.ID = Deposit.BranchID " _
& "WHERE Deposit.AccountID = " & Request.Form("lstAccID") & " " _
& "AND Deposit.Date >= #" & Request.Form("txtStartDt") & "# " _
& "And Deposit.Date <= #" & Request.Form("txtEndDt") & "# "

		oRs.Open sSQL, oCon
		Response.Write "<BR><B>Offline</B> Transaction (Debit/Credit)"
		'Response.Write "<BR>Offline Amounts "
		Response.Write "<TABLE BORDER=1>"
		Response.Write "<TR><TH>Branch ID</TH><TH>Date</TH><TH>Amount</TH><TH>Debit/Credit</TH></TR><TR><TD><FONT color=blue>"
		curTotal = 0
		if not oRs.EOF then
			Response.Write oRs.GetString(,,"</TD><TD><FONT color=blue>","</TD><TD><FONT color=blue>Dr</TD></TR><TR><TD><FONT color=blue>","&nbsp;")
			oRs.MoveFirst 
			while not oRs.EOF 
				curTotal = curTotal + CCur(oRs("Amount"))
				oRs.MoveNext 
			wend
			curBalance = curBalance + curTotal
		end if
		'Response.Write "</TD></TR></TABLE><HR>"	
		oRs.Close
		
		'ssssssss
		'total deposit 
'		sSQL="SELECT Sum(Deposit.Amount) AS Total FROM Deposit " _
'& "GROUP BY Deposit.AccountID HAVING (((Deposit.AccountID)=" _
'& Request.Form("lstAccID") & "));"
		
'		oRs.Open sSQL, oCon
		
		Response.Write "&nbsp;</TD><TD><B>Total</B></TD><TD><B>"
		Response.Write curTotal'oRs("Total")
		Response.Write "</B></TD><TD>&nbsp;</TD></TR><TR><TD><FONT color=red>"
		
'		oRs.Close 
		'eeeeee
		
		'Withdraw		
		'sSQL="SELECT * FROM	Withdraw WHERE AccountID=" & oRsAcc("ID")'& Session("CustID")
		sSQL="SELECT Branch.Title, Withdraw.Date, Withdraw.Amount " _
& "FROM Branch INNER JOIN Withdraw ON Branch.ID = Withdraw.BranchID " _
& "WHERE Withdraw.AccountID = " & Request.Form("lstAccID") & " " _
& "AND Withdraw.Date >= #" & Request.Form("txtStartDt") & "# " _
& "And Withdraw.Date <= #" & Request.Form("txtEndDt") & "# "

		oRs.Open sSQL, oCon
		'Response.Write "<BR>Withdraw Amounts"
		'Response.Write "<TABLE BORDER=1>"
		'Response.Write "<TR><TH>Account ID</TH><TH>Branch ID</TH><TH>ID</TH><TH>Date</TH><TH>Amount</TH></TR><TR><TD>"
		'Response.Write "-</TD><TD>-</TD><TD>-</TD><TD>-</TD></TR><TR><TD>"
		'Response.Write "</TD>/TR><TR><TD>"
		Response.Write "</TD><TD></TD><TD></TD><TD></TD></TR><TR><TD><FONT color=red>"
		'Response.Write "&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD></TR><TR><TD>"
		curTotal = 0
		if not oRs.EOF then
			Response.Write oRs.GetString(,,"</TD><TD><FONT color=red>","</TD><TD><FONT color=red>Cr</TD></TR><TR><TD><FONT color=red>","&nbsp;")
			
			oRs.MoveFirst 
			while not oRs.EOF 
				curTotal = curTotal + CCur(oRs("Amount"))
				oRs.MoveNext 
			wend
			curBalance = curBalance - curTotal
		end if
		'Response.Write "</TD></TR></TABLE>"	
		oRs.Close
		
		'ssssssssssss
'		sSQL="SELECT Sum(Withdraw.Amount) AS Total FROM Withdraw " _
'& "GROUP BY Withdraw.AccountID " _
'& "HAVING (((Withdraw.AccountID)= " & Request.Form("lstAccID") &"));"
'		oRs.Open sSQL, oCon
	'if not oRs.EOF then
		Response.Write "&nbsp;</TD><TD><B>Total</B></TD><TD><B>"
		Response.Write curTotal'oRs("Total")
		Response.Write "</TD><TD>&nbsp;</TD></TR><TABLE>"
	'end if
		'Response.Write "</TD></TR></TABLE><HR>"	
'		oRs.Close
		'eeeeeeeeeeee
		
		'ssssssssssssssss
'		Dim oCmd
'		Set oCmd = Server.CreateObject("ADODB.Command")
'		oCmd.ActiveConnection = oCon
'		oCmd.CommandText = "qFindBalance"
'		oCmd.CommandType = adCmdStoredProc
	
'		dim oParm
'		set oParm = server.CreateObject("ADODB.Parameter")
'		oParm.Name ="AID"
'		oParm.Type=adInteger
'		oParm.Direction = adParamInput
'		oParm.Value = Request.Form("lstAccID")
'		oCmd.Parameters.Append oParm
		'eeeeeeeeeeeeeeee  
		
		'sssssssss
'		set oRs = oCmd.Execute
		'esponse.Write "<BR>1. EOF = " & oRsBal.EOF
		'esponse.Write "<BR>1. BOF = " & oRsBal.BOF   
'		if not oRs.EOF then ' or not oRsBal.BOF  
			
			Response.Write "<br><H2>Account Balance = " & curBalance & "</H2>"
			'Response.Write "<BR>Session('UserID') = " & Session("UserID") & "<BR>"
			'Response.Write "Account Balance"
			'Response.Write "<TABLE BORDER=1>"
			'Response.Write "<TR><TH>Balance</TH></TR><TR><TD>"
			'Response.Write oRs.GetString(,,"</TD><TD>","</TD></TR><TR><TD>","&nbsp;")
			'Response.Write "</TD></TR></TABLE>"
'		end if
'		oRs.Close 		
		'eeeeeeeeee
		
		'oRsAcc.MoveNext 
	'wend
	
	'RsAcc.Close 
	oCon.Close 
%>
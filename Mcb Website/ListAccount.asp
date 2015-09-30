<!-- #include file="AuthenticateAccHolder.asp" -->
<body background="b.jpg">
<!-- # include file="NavigatingAccHolder.htm" -->
<%
	set oRs = server.CreateObject("ADODB.recordset")	
	Dim strSQL
	's="select * from Account where UserID=" & Session("UserID")
	's-q
strSQL = "SELECT Account.ID, Branch.Title, AccountType.Title, Account.AccountLoginName  " & _
"FROM AccountType INNER JOIN (Branch INNER JOIN Account ON " & _
"Branch.ID =Account.BranchID) ON " & _
"AccountType.ID = Account.AccountTypeID " & _
"WHERE Account.UserID=" & Session("UserID")
	'e-q
	oRs.Open strSQL, "DSN=dsnMCB"
	''''''
	%>
	<font face="arial" color="green"><h2>
	Account Holder &rarr; <FONT size=7>A</FONT>ccount List</h2></font>
	<hr>
	<%
	'Response.Write "<HR><H3>List of Account</H3><HR>"
	
	Response.Write "<TABLE BORDER=1>"
	Response.Write "<TR><TH>Account #</TH><TH>Branch</TH><TH>Account Type</TH><TH>Account Login Name</TH></TR><TR><TD>"
	Response.Write oRs.GetString(,,"</TD><TD>","</TD></TR><TR><TD>","&nbsp;")
	Response.Write "</TD></TR></TABLE>"
	
'	while not ors.EOF
'		Response.Write oRs("ID")& "</TD><TD>"
'		Response.Write oRs("BranchID") & "</TD><TD>"
'		Response.Write oRs("AccountTypeID") & "</TD></TR><TR><TD>"
'		oRs.MoveNext 
'	wend
'	Response.Write "</TD></TR></TABLE>"
		'ors.Update
	'end if
	oRs.Close 
%>
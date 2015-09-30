<%@ Language=VBScript %>
<body background="b.jpg">
<!-- # include file="NavigatingAccHolder.htm" -->
<font face="arial" color="green"><h2>
Account Holder &rarr; <FONT size=7>R</FONT>efered Persons</h2></font>
<hr>
<%
	dim oRs
	set oRs = server.CreateObject("ADODB.recordset")
	'Response.Write "H1"
	dim sSQL
	sSQL = "SELECT Name, NIC, [Date] FROM	ReferTo WHERE UserID = " & Session("UserID") _
		& " ORDER BY Name"
	oRs.Open sSQL,"DSN=dsnMCB"

	Response.Write "Refered Person Detail"
	Response.Write "<TABLE BORDER=1>"
	'UserID	ID	Name	NIC
	Response.Write "<TR><TH><B>Name</B></TH><TH><B>NIC #</B></TH><TH><B>Date</B></TH></TR><TR><TD>"
	if not oRs.EOF then
		Response.Write oRs.GetString(,,"</TD><TD>","</TD></TR><TR><TD>","&nbsp;")
	end if
	Response.Write "</TD></TR></TABLE>"	
	oRs.Close
%>
<hr>
<body>
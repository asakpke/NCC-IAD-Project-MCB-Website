<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<%
	set oRs = server.CreateObject("ADODB.recordset")	
	Dim s
	s="SELECT OrderMaster.CustName, OrderMaster.CustAddr, OrderMaster.Date " _
	& "FROM OrderMaster " _
	& "WHERE (((OrderMaster.State)=False))" _
	& "ORDER BY OrderMaster.Date DESC;"

	oRs.Open s, "DSN=dsnShop"
	
	Response.Write "<HR><H1>Not Confirmed Order List</H1>"
	Response.Write "<TABLE BORDER=1>"
	Response.Write "<TR><TH>Customer</TH><TH>Address</TH><TH>Date</TH></TR><TR><TD>"
	Response.Write oRs.GetString(,,"</TD><TD>","</TD></TR><TR><TD>","&nbsp;")
	Response.Write "</TD></TR></TABLE>"
	
	oRs.Close 
%>

</BODY>
</HTML>

<HTML>
<HEAD>
</HEAD>
<BODY background="b.jpg">
<font face="arial" color="green"><h1>
Products List</h1>
</font>
<hr>
<FORM ACTION="checkout.asp" id=frmPro name=frmPro METHOD="post">
<UL>
  <LI><INPUT id=chkItem1 name=chkItem1 type=checkbox>Item 1, Rs.1000</LI><INPUT name=Rs1 type=hidden value=1000>
  <LI><INPUT id=chkItem2 name=chkItem2 type=checkbox>Item 2, Rs.3000</LI><INPUT name=Rs2 type=hidden value=3000>
  <LI><INPUT id=chkItem3 name=chkItem3 type=checkbox>Item 3, Rs.5000</LI><INPUT name=Rs3 type=hidden value=5000>
  <LI><INPUT id=chkItem4 name=chkItem4 type=checkbox>Item 4, Rs.2000</LI><INPUT name=Rs4 type=hidden value=2000>
  <LI><INPUT id=chkItem5 name=chkItem5 type=checkbox>Item 5, Rs.4000</LI><INPUT name=Rs5 type=hidden value=4000>
</UL>
  <!--
  <%
  's
	set oRs = server.CreateObject("ADODB.recordset")	
	Dim s
	s="SELECT * FROM Item"
	oRs.Open s, "DSN=dsnShop"
	''''''
	Response.Write "<TABLE BORDER=>"
	Response.Write "<TR><TH>ID</TH><TH>Item</TH><TH>Price</TH></TR><TR><TD>"
	'Response.Write oRs.GetString(,,"</TD><TD>","</TD></TR><TR><TD>","&nbsp;")
	'Response.Write "</TD></TR></TABLE>"
	
	while not ors.EOF
		Response.Write "<INPUT name=""" & oRs("ID") & """ type=checkbox>"
		Response.Write oRs("ID")& "</TD><TD>"
		Response.Write oRs("Name") & "</TD><TD>"
		Response.Write oRs("Price") & "</TD></TR><TR><TD>"
		oRs.MoveNext 
	wend
	Response.Write "</TD></TR></TABLE>"
		'ors.Update
	'end if
	oRs.Close 
	'e
%>
-->
<INPUT id=submit1 name=submit1 type=submit value="Check Out">&nbsp;<INPUT id=reset1 name=reset1 type=reset value=Reset>
</FORM>
<hr>
</BODY>
</HTML>
<HTML>
<HEAD>
</HEAD>
<BODY background="b.jpg">
<font face="arial" color="green"><h1>
Products List</h1>
</font>
<hr>
<FORM ACTION="checkout.asp" id=frmPro name=frmPro METHOD="post">
 <%
  's
	set oRs = server.CreateObject("ADODB.recordset")	
	Dim s
	s="SELECT * FROM Item"
	oRs.Open s, "DSN=dsnShop"
	''''''
	Response.Write "<TABLE BORDER=0>"
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
<hr>
<table>
	<tr><td>Enter your Name<td><INPUT name=txtName maxlength=20 ></td>
	<tr><td>Enter your Address<td><INPUT name=txtAddr maxlength=50 style="HEIGHT: 22px; WIDTH: 343px"></td></tr>
</table>
<INPUT id=submit1 name=submit1 type=submit value="Check Out">&nbsp;<INPUT id=reset1 name=reset1 type=reset value=Reset>
</FORM>
<hr>
</BODY>
</HTML>
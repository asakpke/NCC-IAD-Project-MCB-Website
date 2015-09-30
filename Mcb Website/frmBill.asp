<!-- #include file="AuthenticateAccHolder.asp" -->	
<!-- #include file="Msg.asp" -->

<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub FORM1_onsubmit
	if not IsNumeric(form1.lstFromAccID.value) or _
			not IsNumeric(form1.txtBillNo.value) or _
			not IsNumeric(form1.txtAmount.value) then
		msgbox "Enter Numeric Value"
		form1.action =""
	end if
	
	if form1.lstComID.value = "" then
		msgbox "Your must select a billing company"
		form1.action =""
	end if
End Sub

-->
</SCRIPT>
</HEAD>
<BODY background="b.jpg">
<font face="arial" color="green"><h2>
Account Holder &rarr; <FONT size=7>B</FONT>ill Payment</h2></font>
<hr>
<!-- # include file="NavigatingAccHolder.htm" -->
<!--'<h1>Bill payments options</h1>-->
<FORM action="Bill.asp" id=FORM1 method=post name=FORM1>
<table>
<Tr><td>Select your Account #</td><td><SELECT NAME=lstFromAccID SIZE=1>
<%
	dim Cn
	set Cn = server.CreateObject("ADODB.Connection")
	Cn.Open "DSN=dsnMCB"
	
	dim oRsAcc
	set oRsAcc = server.CreateObject("ADODB.recordset")	
	Dim sSQL
	sSQL = "SELECT * FROM	Account WHERE UserID = " & Session("UserID")
	oRsAcc.Open sSQL, Cn

		while not oRsAcc.EOF  
		Response.Write "<OPTION VALUE=" & oRsAcc("ID") & ">"
		Response.Write oRsAcc("ID") & "</OPTION>"	
		oRsAcc.MoveNext 
	wend
	oRsAcc.Close	 
%>
</SELECT>
		</td></tr>

	<tr><td>Select Billing Company</td><td><SELECT NAME=lstComID SIZE=1>
<%
	dim oRsCom
	set oRsCom = server.CreateObject("ADODB.recordset")	
	'Dim sSQL
	'sSQL = "SELECT * FROM	BillingCompany WHERE ID = " & Session("UserID")
	oRsCom.Open "BillingCompany", Cn

	while not oRsCom.EOF 
		Response.Write "<OPTION VALUE=" & oRsCom("ID") & ">"
		Response.Write oRsCom("Title") & "</OPTION>"	
		oRsCom.MoveNext 
	wend
	oRsCom.Close	 
%>		
</SELECT>
</tr></td>

<tr><td>Enter Bill #</td><td><INPUT id="txtBillNo" name="txtBillNo" maxlength=9></td></tr>
<tr><td>Amout</td><td><INPUT id="txtAmount" name="txtAmount"></td></tr>
</table>
<%
	dim oRsTranCharges
	set oRsTranCharges = server.CreateObject("ADODB.recordset")
	oRsTranCharges.Open "Settings", Cn
	dim nTranCharges
	nTranCharges = oRsTranCharges("TransactionCharges")
	oRsTranCharges.Close 
%>
<P><B>Note</B> Transaction charges = Rs <%=nTranCharges%> </P>
<INPUT id=submit name=submit type=submit value=Submit> <INPUT id=reset name=reset type=reset value=Reset></P>
</FORM>

</BODY>
</HTML>


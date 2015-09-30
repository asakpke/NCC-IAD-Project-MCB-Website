<!-- #include file="AuthenticateAccHolder.asp" -->
<!-- #include file="Msg.asp" -->
<!-- # include file="NavigatingAccHolder.htm" -->

<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub FORM1_onsubmit
	if not IsNumeric(form1.txtToAccID.value ) or _
			not IsNumeric(form1.txtAmount.value ) or _
			not IsNumeric(form1.lstAcc.value)  then
		msgbox "Enter Numeric Value"
		form1.action =""
	end if
End Sub

-->
</SCRIPT>
</HEAD>
<BODY background="b.jpg">
<font face="arial" color="green"><h2>
Account Holder &rarr; <FONT size=7>M</FONT>oney Transfer</h2></font>
<hr>
<FORM action="Transfer.asp" id=FORM1 method=post name=FORM1>
<P>&nbsp;</P>
<table>
	<tr><td>Select your Account #</td><td><SELECT NAME=lstAcc SIZE=1>
<%
	dim oCn
	set oCn = server.CreateObject("ADODB.Connection")
	oCn.Open "DSN=dsnMCB"
	
	dim oRsAcc
	set oRsAcc = server.CreateObject("ADODB.recordset")	
	Dim sSQL
	
	sSQL = "SELECT * FROM	Account WHERE UserID = " & Session("UserID")
	oRsAcc.Open sSQL, oCn

	while not oRsAcc.EOF 
		Response.Write "<OPTION VALUE=" & oRsAcc("ID") & ">"
		Response.Write oRsAcc("ID") & "</OPTION>"	
		oRsAcc.MoveNext 
	wend
	oRsAcc.Close	 
%>	
</SELECT>
		</td></tr>
	<tr><td>ToAccountID</TD><td><INPUT id="txtToAccID" name="txtToAccID" maxlength=9></td></tr>
	<tr><td>Amount</td><td><INPUT id="txtAmount" name="txtAmount" maxlength=14></td></tr>
</table>
<%
	dim oRsTranCharges
	set oRsTranCharges = server.CreateObject("ADODB.recordset")
	oRsTranCharges.Open "Settings", oCn
	dim nTranCharges
	nTranCharges = oRsTranCharges("TransactionCharges")
	oRsTranCharges.Close 
%>
<P><B>Note</B> Transaction charges = Rs <%=nTranCharges%> </P>
<P>&nbsp;<INPUT id=submit name=submit type=submit value=Submit style="LEFT: 13px; TOP: 182px">&nbsp;<INPUT id=reset name=reset type=reset value=Reset></P>
</FORM>

</BODY>
</HTML>


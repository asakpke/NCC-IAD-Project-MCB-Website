<!-- #include file="AuthenticateAccHolder.asp" -->
<!-- #include file="Msg.asp" -->
<%
	dim oRsAcc
	set oRsAcc = server.CreateObject("ADODB.recordset")	
	Dim sSQL
	sSQL = "SELECT * FROM	Account WHERE UserID = " & Session("UserID")
	oRsAcc.Open sSQL, "DSN=dsnMCB"
%>		
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub FORM1_onsubmit
	if Not IsNumeric(form1.lstAccID.value) or Not _
			IsDate(form1.txtStartDt.value) or not _
			IsDate(form1.txtEndDt.value) then
		msgbox "Enter valid Value"
		form1.action=""
	end if
End Sub

-->
</SCRIPT>
</HEAD>
<BODY background="b.jpg">
<font face="arial" color="green"><h2>
Account Holder &rarr; <FONT size=7>S</FONT>elect Account</h2></font>
<hr>
<!-- # include file="NavigatingAccHolder.htm" -->
<FORM action="BankStatementDate.asp" id=FORM1 method=post name=FORM1>
<P>Select an Account # 
<SELECT NAME=lstAccID SIZE=1>
<%
	while not oRsAcc.EOF 
		Response.Write "<OPTION VALUE=" & oRsAcc("ID") & ">"
		Response.Write oRsAcc("ID") & "</OPTION>"	
		oRsAcc.MoveNext 
	wend
	oRsAcc.Close	 
%>	
</SELECT>
</P>
<P>Enter start date <INPUT id=txtStartDt name=txtStartDt type=text maxlength=10> (Date Formate = MM/DD/YYYY)</P>
<P>Enter end date <INPUT id=txtEndDt name=txtEndDt type=text maxlength=10></P>
<P><INPUT id=submit name=submit type=submit value=Submit>
<INPUT id=reset name=reset type=reset value=Reset></P>
</FORM>

</BODY>
</HTML>
<!-- #include file="AuthenticateAdmin.asp" -->
<!-- #include file="Msg.asp" -->
<!-- # include file="NavigatingMain.htm" -->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub frmNewCust_onsubmit
	dim strErr 
	strErr=""
	if frmNewCust.txtPhone.value <> "" and not IsNumeric(frmNewCust.txtPhone.value) then
		strErr = strErr & "Enter valid Phone number. "
	end if
'	
'	if form1.txtAddr.value  = "" then
'		strErr = strErr & "please enter the correct address. "
'	end if
	if len(strErr) > 0 then
		frmNewCust.action = ""
		MsgBox strErr
	end if
End Sub

-->
</SCRIPT>
</HEAD>
<BODY  background="b.jpg">
<font face="arial" color="green"><h2>
MCB &rarr; <FONT size=7>N</FONT>ew Customer</h2></font>
<hr>
You must fill the fields with bold text.
<FORM action=AdminCreateNewCustomer.asp id=frmNewCust name=frmNewCust METHOD=post>
<TABLE>
	<TR><TD><STRONG>Name</STRONG></TD><TD><INPUT id=txtName name=txtName maxlength=20></TD></TR>
	<TR><TD><STRONG>NIC #</STRONG></TD><TD><INPUT id=txtNIC name=txtNIC maxlength=20></TD></TR>
	<TR><TD>Father Name</TD><TD><INPUT id=txtFName name=txtFName maxlength=20></TD></TR>
	<TR><TD>Address</TD><TD><INPUT id=txtAddr name=txtAddr maxlength=30></TD></TR>
	<TR><TD>Phone</TD><TD><INPUT id=txtPhone name=txtPhone maxlength=9></TD></TR>
	<TR><TD><STRONG>LoginName</STRONG></TD><TD><INPUT id=txtLoginName name=txtLoginName maxlength=15></TD></TR>
	<TR><TD><STRONG>Password</STRONG></TD><TD><INPUT id=txtPwd name=txtPwd type=password maxlength=15></TD></TR>

<!--
<P>Select Branch: 
<SELECT NAME=BranchName SIZE=1>
<OPTION VALUE="ak">Attock</OPTION>
</SELECT>
</P>
-->

	<TR><TD>Select <STRONG>Branch</STRONG></TD><TD><SELECT NAME=lstBranchID SIZE=1>
<%
	dim Cn
	set Cn = server.CreateObject("ADODB.Connection")
	Cn.Open "DSN=dsnMCB"
	
	dim oRs
	set oRs = server.CreateObject("ADODB.recordset")	
	Dim sSQL
	'sSQL = "SELECT * FROM	Branch"
	oRs.Open "Branch", Cn

	while not oRs.EOF  
		Response.Write "<OPTION VALUE=" & oRs("ID") & ">" & oRs("Title") _
		& "</OPTION>"	
		'Response.Write oRsAcc("Title") & "</OPTION>"	
		oRs.MoveNext 
	wend
	oRs.Close	 
%> 
</SELECT></TD></TR>

	<TR><TD>Select the <STRONG>Account Type</STRONG></TD><TD><SELECT NAME=lstAccTypeID SIZE=1>
<%
'dim Cn
	'set Cn = server.CreateObject("ADODB.Connection")
	'Cn.Open "DSN=dsnMCB"
	
	'dim oRsAcc
	'set oRsAcc = server.CreateObject("ADODB.recordset")	
	'Dim sSQL
	'sSQL = "SELECT * FROM	AccountType"
	oRs.Open "AccountType", Cn

	while not oRs.EOF  
		Response.Write "<OPTION VALUE=" & oRs("ID") & ">"
		Response.Write oRs("Title") & "</OPTION>"	
		oRs.MoveNext 
	wend
	oRs.Close	 
%>
</select></TD></TR>

	<TR><TD><STRONG>Login Name of Account</STRONG></TD><TD><INPUT id=txtPinCode name=txtPinCode maxlength=9></TD></TR>
	<TR><TD><STRONG>Password of Account</STRONG></TD><TD><INPUT id=txtAccPwd name=txtAccPwd type=password maxlength=15></TD></TR>
	<TR><TD></TABLE>
<P><INPUT id=cmdSubmit name=cmdSubmit type=submit value=Submit> <INPUT id=cmdReset name=cmdReset type=reset value=Reset></P></TD></TR>
<hr>
</FORM>
</BODY>
</HTML>

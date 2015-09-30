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
	dim strErr
	strErr = ""
	
	
	
	if form1.txtPwd.value  = "" or form1.txtPC.value = "" then
		strErr = strErr & "Fill field. "
	end if
	
	if len(strErr) > 0 then
		form1.action = ""
		MsgBox strErr
	end if
End Sub

-->
</SCRIPT>
</HEAD>
<BODY background="b.jpg">
<font face="arial" color="green"><h2>
Account Holder &rarr; <FONT size=7>N</FONT>ew Account</h2></font>
<hr>


<FORM action="NewAccount.asp" id=FORM1 method=post name=FORM1>
	<table>
<tr><td>Select Branch:</td><td><SELECT NAME=lstBID SIZE=1>
<%
	dim Cn
	set Cn = server.CreateObject("ADODB.Connection")
	Cn.Open "DSN=dsnMCB"
	
	dim oRsAcc
	set oRsAcc = server.CreateObject("ADODB.recordset")	
	Dim sSQL
	sSQL = "SELECT * FROM	Branch"
	oRsAcc.Open sSQL, Cn

		while not oRsAcc.EOF  
		Response.Write "<OPTION VALUE=" & oRsAcc("ID") & ">"
		Response.Write oRsAcc("Title") & "</OPTION>"	
		oRsAcc.MoveNext 
	wend
	oRsAcc.Close	 
%>
</SELECT>


<tr><td>Account Login Name</td><td><INPUT id=txtPC name=txtPC maxlength=9></td></tr>
<tr><td>Password</td><td><INPUT type=password id=txtPwd name=txtPwd maxlength=15></td></tr>

<tr><td>Select the Account Type:</td><td><SELECT NAME=lstAccType SIZE=1>
<%
'dim Cn
	set Cn = server.CreateObject("ADODB.Connection")
	Cn.Open "DSN=dsnMCB"
	
	'dim oRsAcc
	set oRsAcc = server.CreateObject("ADODB.recordset")	
	'Dim sSQL
	sSQL = "SELECT * FROM	AccountType"
	oRsAcc.Open sSQL, Cn

		while not oRsAcc.EOF  
		Response.Write "<OPTION VALUE=" & oRsAcc("ID") & ">"
		Response.Write oRsAcc("Title") & "</OPTION>"	
		oRsAcc.MoveNext 
	wend
	oRsAcc.Close	 

%>
</select>
</tr></td>
</table>
<P><INPUT id=submit name=submit type=submit value=Submit></P>
</FORM>
</BODY>
</HTML>
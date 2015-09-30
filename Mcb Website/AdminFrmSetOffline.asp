<!-- #include file="AuthenticateAdmin.asp" -->
<!-- #include file="Msg.asp" -->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub form1_onsubmit
	if 	not IsNumeric(form1.txtAccID.value) and form1.txtState.value >=0 and _
			form1.txtState.value <=3 then
		msgbox "Enter valid Numeric Value"
		form1.action =""
	end if
End Sub

-->
</SCRIPT>
</HEAD>
<BODY background="b.jpg">
<font face="arial" color="green"><h2>
MCB &rarr; <FONT size=7>S</FONT>et Account state</h2></font>
<hr>
<FORM action="AdminSetOffline.asp" method="post" id=form1 name=form1><BR>
<TABLE>
	<tr><td>Enter UserID<td><INPUT id=txtAccID name=txtAccID maxlength=9></tr>
	<tr><td>Enter State<td><INPUT id=txtState name=txtState maxlength=1></tr>
</TABLE>
state 0 = off user, 1 = only online, 2 = only offline, 3 = both<br><br>
<INPUT id=submit1 name=submit1 type=submit value=Submit>&nbsp;<INPUT id=reset1 name=reset1 type=reset value=Reset>
<br></FORM>
</BODY>
</HTML>
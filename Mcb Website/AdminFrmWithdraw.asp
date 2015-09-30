<!-- #include file="AuthenticateAdmin.asp" -->
<!-- #include file="Msg.asp" -->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub form1_onsubmit
	if 	not IsNumeric(form1.txtAccID.value) or _
			not IsNumeric(form1.txtAmount.value) then
		msgbox "Enter Numeric Value"
		form1.action =""
	end if
End Sub

-->
</SCRIPT>
</HEAD background="b.jpg">
<BODY  background="b.jpg">
<font face="arial" color="green"><h2>
MCB &rarr; <FONT size=7>W</FONT>ithdraw of money</h2></font>
<hr>
<FORM action="AdminWithdraw.asp" method="post" id=form1 name=form1><BR>
<TABLE>
	<tr><td>Enter Account #<td><INPUT id=txtAccID name=txtAccID maxlength=9></tr>
	<tr><td>Enter Amount<td><INPUT id=txtAmount name=txtAmount maxlength=14></tr>
</TABLE>
<INPUT id=submit1 name=submit1 type=submit value=Submit>&nbsp;<INPUT id=reset1 name=reset1 type=reset value=Reset>
<br></FORM>
</BODY>
</HTML>
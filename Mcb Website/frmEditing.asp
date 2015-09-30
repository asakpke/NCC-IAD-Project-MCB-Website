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
'	
	if form1.txtPhone.value  <> "" And not IsNumeric(form1.txtPhone.value)then
		strErr = strErr & "Enter valid Phone number. "
	end if
'	
'	if form1.txtAddr.value  = "" then
'		strErr = strErr & "please enter the correct address. "
'	end if
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
Account Holder &rarr; <FONT size=7>E</FONT>diting Options</h2></font>
<hr>
<DIV align=center><FONT color=blue size=5>
<P>
<FORM action="Editing.asp" id=FORM1 method=post name=FORM1>Password<INPUT id=txtPwd name=txtPwd type=password maxlength=10></P>
<P>Address&nbsp;&nbsp; <INPUT id=txtAddr name=txtAddr type=text maxlength=30></P>
<P>Phone&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <INPUT id=txtPhone name=txtphone 
type=text maxlength=9></P>
<P>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<INPUT id=cmdsubmit name=cmdsubmit style="COLOR: crimson; FONT-STYLE: italic" type=submit value=Submit>&nbsp;&nbsp;<INPUT id=cmdreset name=cmdreset style="COLOR: fuchsia; FONT-STYLE: italic; HEIGHT: 25px; WIDTH: 67px" type=reset value=Reset></P></FORM>
<P>&nbsp;</P> 
</FONT></DIV>
</BODY>
</HTML>


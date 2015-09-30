<!-- #include file="AuthenticateAccHolder.asp" -->
<!-- #include file="Msg.asp" -->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub frmRefTo_onsubmit
	dim strErr
	strErr = ""
	
	'msgbox "NIC type = " & vartype(frmRefTo.txtNIC.value)
	
	if frmRefTo.txtName.value  = "" or _
			frmRefTo.txtNIC.value  = "" then
		strErr = strErr & "Fill all field. "
	end if
	if IsNumeric(frmrefto.txtName.value) then
		strErr = strErr & "Enter valid name. "
	end if
	'msgbox "frmRefTo.checkbox1.status = " & frmRefTo.checkbox1.status 
	if frmRefTo.checkbox1.status <> true then
		strErr = strErr & "Your should tick the check box. "
	end if
	
	if len(strErr) > 0 then
		frmRefTo.action =""
		MsgBox strErr
	end if
End Sub

-->
</SCRIPT>
</HEAD>
<BODY background="b.jpg">
<!-- # include file="NavigatingAccHolder.htm" -->
<font face="arial" color="green"><h2>
Account Holder &amp;rarr <FONT size=7>R</FONT>efer To Person</h2></font>
<hr>
<FORM action="ReferTo.asp" id=frmRefTo method=post name=frmRefTo>
<P>Name <INPUT id=txtName name=txtName maxLength=20></P>
<P>NIC # <INPUT id=txtNIC name=txtNIC maxLength=20>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<STRONG><EM>Format of NIC #</EM></STRONG> (37102-1759907-5)</P>
<P>
<INPUT id=checkbox1 name=checkbox1 type=checkbox value ="">          
             
      Are you agree to refered the&nbsp;above person.If 
the person that you are refering make anykind of illegal work.So you are resonible for 
that&nbsp;person.</P>
<P><INPUT id=cmdSubmit name=cmdSubmit type=submit value=Submit></P>
</FORM>
<hr>
</BODY>
</HTML>
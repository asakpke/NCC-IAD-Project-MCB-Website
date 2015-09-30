<%@ Language=VBScript %>
<%'<!-- #include file="UserInfo.asp" -->%>
<!-- #include file="Msg.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<frameset rows="5%,,5%" border=0>
	<frameset cols="50%,50%">
		<frame src="top1.html" scrolling="no">
		<frame src="top2.html" scrolling="no">
	</frameset>
	<frameset cols="17%,,17%">
		<!--	<frame src="content.html"  name="cont" scrolling="no">		-->
		<frame src="content.html"  name="cont">
		<frame src="default.html" name="b">
		<frame src="external.html" scrolling="no">
	</frameset>   
	<frame src="bottom.html" scrolling="no">
</frameset>
<META content="text/html; charset=unicode" http-equiv=Content-Type>
<META content="MSHTML 5.00.2614.3500" name=GENERATOR>

<!--
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
function cmdSubmit_onclick() {
	//prompt( "sdfjkskdfj" );
	if( document.frmLogin.txtName.value == "" 
			|| document.frmLogin.txtPwd.value  == "")
		{
		//Alert("You must enter user name and password");
		return false;
		//window.frames.
		}
}

//
</SCRIPT>
-->
</HEAD>
<!--
<BODY>
<%'="<SCRIPT LANGUAGE=VBScript>MsgBox " & chr(34) & "Abc" & chr(34) & "'chr(34)=quot</Script>"%>
<%
'Response.Write "Aamir Said! ""God will helps us"""
%>
<H1 align=center>Welcome to MCB Website</H1>
<FORM action="Login.asp" method=post id=frmLogin name=frmLogin>
<P align=right>&nbsp;</P>
<P align=right>Name <INPUT id=txtName name=txtName></P>
<P align=right>Password <INPUT id=txtPwd name=txtPwd type=password></P>
<P align=right><INPUT id=cmdSubmit name=cmdSubmit type=submit value=Login LANGUAGE=javascript onclick="return cmdSubmit_onclick()">&nbsp;
<INPUT id=cmdReset name=cmdReset type=reset value=Reset></P>
</FORM>
<P align=right>
<A href="frmNewCustomer.asp">Add New Customer</A> </P>
</BODY>
-->
</HTML>








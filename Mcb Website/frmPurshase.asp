<%@ Language=VBScript %>

<!-- #include file="Msg.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

//function button1_onclick() {
	//alert( window.history.back  );
	//alert(  )
	//window.history.back(1);
//}

//-->
</SCRIPT>
</HEAD>
<BODY background="b.jpg">
<%
'	dim L
'	for L=1 to Request.ServerVariables.Count 
'		Response.Write "<hr>L = " & L & ", " & Request.ServerVariables.Item(L)  
'	next
'	Response.End 
'Response.Write "<hr>L = " & 47 & ", " & Request.ServerVariables.Item(47)  
if Request.Form("txtTotal") < 1 then
	Response.Write "<br>Total amount is 0"
	Response.End 
end if
'Response.Write "<br>txtShopAccID = " & Request.Form("txtShopAccID")
'Response.Write "<br>txtTotal = " & Request.Form("txtTotal")
'Response.Write "<br>txtUserAccID = " & Request.Form("txtUserAccID")
%>
<h2>Welcome to MCB's Online Shopping Service</h2>
<FORM action="Purshase.asp" method="post"><BR>
<INPUT id=txtShopAccID name=txtShopAccID type=hidden value=<%=Request.Form("txtShopAccID")%>>
<INPUT id=txtTotal name=txtTotal type=hidden value=<%=Request.Form("txtTotal")%>>
<INPUT id=txtUserAccID name=txtUserAccID type=hidden  value=<%=Request.Form("txtUserAccID")%>>
<INPUT name=txtShopURL type=hidden  value=<%=Request.Form("txtShopURL")%>>
<INPUT name=txtOrderID  type=hidden value=<%=Request.Form("txtOrderID")%>>
<TABLE>
	<tr><td>Enter Account Login Name<td><INPUT id=txtAccPin name=txtAccPin></tr>
	<tr><td>Enter Password<td><INPUT id=txtAccPwd name=txtAccPwd type=password></tr>
</TABLE>
<%
	dim oRsTranCharges
	set oRsTranCharges = server.CreateObject("ADODB.recordset")
	oRsTranCharges.Open "Settings", "dsnMCB"
	dim nTranCharges
	nTranCharges = oRsTranCharges("TransactionCharges")
	oRsTranCharges.Close 
%>
<P><B>Note</B> Transaction charges = Rs <%=nTranCharges%> </P>
<INPUT id=submit1 name=submit1 type=submit value=Submit>&nbsp;<INPUT id=reset1 name=reset1 type=reset value=Reset>
<br></FORM>
</BODY>
</HTML>
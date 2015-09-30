<!-- #include file="AuthenticateComp.asp" -->
<!-- #include file="ADOvbs.inc" -->
<!-- # include file="NavigatingBillingComp.htm" -->
<%'<!-- #include file="UserInfo.asp" -->%>
<BODY background="b.jpg">
<font face="arial" color="green"><h2>
Billing Company &rarr; <FONT size=7>P</FONT>aid Bill Detail</h2></font>
<hr>
<%
	'Response.Write "CompID = " & Session("UserID")
	'Response.Write "Dt = " & Request.Form("txtDate")
	'Response.End 
'checking date
if isdate(Request.Form("txtDate")) then
'	Response.Write "Dt = " & Request.Form("txtDate")
'	Response.Write "<BR>Day = " & day(Request.Form("txtDate"))
'	Response.Write "<BR>Mon = " & Month(Request.Form("txtDate"))
'	Response.Write "<BR>Year = " & Year(Request.Form("txtDate"))
	'SSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSS
	dim oCon
	set ocon = server.CreateObject("ADODB.Connection")
	oCon.Open "DSN=dsnMCB"
	
	Dim oCmd
	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = oCon
	oCmd.CommandText = "qBillMon"
	oCmd.CommandType = adCmdStoredProc
	
	'SSSSSSSSS
	dim oParm
	set oparm = server.CreateObject("ADODB.Parameter")
	set oParm = oCmd.CreateParameter("ComID",adInteger ,adParamInput,,Session("UserID"))
	ocmd.Parameters.Append oParm
	set oParm = oCmd.CreateParameter("Dt",adDate,adParamInput,,Request.Form("txtDate"))
	oCmd.Parameters.Append oParm
	
	dim oRsBill
	set oRsBill = server.CreateObject("ADODB.recordset")	
	set oRsBill = oCmd.Execute
	'EEEEEEEEE
	if not oRsBill.EOF then
		'SSSSSSSSSSS
		Response.Write "<BR>Bill Details of Date " & Request.Form("txtDate")
		Response.Write "<TABLE BORDER=1>"
		Response.Write "<TR><TH>Customer Name</TH><TH>Customer Account #</TH><TH>Bill #</TH><TH>Date</TH><TH>Bill Amount</TH></TR><TR><TD>"
		Response.Write oRsBill.GetString(,,"</TD><TD>","</TD></TR><TR><TD>","&nbsp;")
		Response.Write "</TD></TR></TABLE><HR>"
		'EEEEEEEEEEE
	else
		Response.Write "No Record"	
	end if
	
	oRsBill.Close 
	oCon.Close 
	'EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
else
	Response.Write "No a valid date"
end if

'checking checkbox
'if Request.Form("chkDate")= "" then
'	Response.Write "0 = " & Request.Form("chkDate")
'else
'	Response.Write "1 = " & Request.Form("chkDate")
'end if
%>
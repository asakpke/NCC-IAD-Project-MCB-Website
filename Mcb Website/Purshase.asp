<%@ Language=VBScript %>
<!-- #include file="ADOvbs.inc" -->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
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
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onload
	'frmPur.submit
End Sub

-->
</SCRIPT>
</HEAD>
<BODY background="b.jpg">
<%
	dim nConfirm, nTranID
	nConfirm = 0
	nTranID = 0

'<comment these line>
	'Response.Write "<br>1.txtShopAccID = " & Request.Form("txtShopAccID")
	'Response.Write "<br>2.txtTotal = " & Request.Form("txtTotal")
	'Response.Write "<br>3.txtUserAccID = " & Request.Form("txtUserAccID")
	'Response.Write "<br>4.txtAccPin = " & Request.Form("txtAccPin")
	'Response.Write "<br>5.txtAccPwd = " & Request.Form("txtAccPwd")
	'Response.Write "<br>6.txtOrderID = " & Request.Form("txtOrderID")
	'Response.Write "<br>7.txtShopURL  = " & Request.Form("txtShopURL")
	'Response.End
'</comment these line>
'extra check
if Request.Form("txtTotal") < 1 then
	Response.Write "<br>Total amount is 0"
	'Response.End
else
'start-confirm pin/pwd
	dim oConPurchase
	set oConPurchase = server.CreateObject("ADODB.Connection")
	oConPurchase.Open "DSN=dsnMCB"
	set oRs = server.CreateObject("ADODB.recordset")
	strSQL = "SELECT * FROM Account WHERE ID=" & Request.Form("txtUserAccID")
	oRs.Open strSQL, oConPurchase

	if not oRs.EOF and not ors.BOF then
		'<comment these line>
			'Response.Write "<br>8.Acc = ok"
			'Response.End
		'</comment these line>
		if Request.Form("txtAccPin") = CStr(oRs("AccountLoginName")) then 'note CStr(Pin=Number)
			'<comment these line>
				'Response.Write "<br>9.pin = ok"
				'Response.End
			'</comment these line>
			if Request.Form("txtAccPwd") = oRs("Password") then
				'<comment these line>
					'Response.Write "<br>10.Pwd = ok"
					'Response.End
				'</comment these line>
				'start - pwd ok code
				Session("AccID") = Request.Form("txtUserAccID")
				Session("Amount") = Request.Form("txtTotal")

				'<comment these line>
					'Response.Write "<br>11.Before ValidateBalance.asp"
					'Response.End
				'</comment these line>

				%>
				<!-- #include file="ValidateBalance.asp" -->
				<%

				'<comment these line>
					'Response.Write "<br>12.After ValidateBalance.asp"
					'Response.End
				'</comment these line>

				if Len( Session( "Msg" ) ) > 0 then

					'<comment these line>
						'Response.Write "<br>Msg > 0"
						'Response.End
					'</comment these line>

					Response.Write "<hr>" & Session("Msg") & "<hr>"
					Session("Msg")=""
					'Response.Redirect "frmPurshase.asp"
					'Response.End
				else

					'<comment these line>
						'Response.Write "<br>Msg < 0 means all ok"
						'Response.End
					'</comment these line>

					'it means that acc,pin/pwd/bal all is ok so
					nConfirm = 1

					'start-Insert Into Transaction
					dim ORsTran
					set oRsTran = server.CreateObject("ADODB.recordset")
					oRsTran.Open "Transaction", oConPurchase,adOpenDynamic ,adLockPessimistic

					oRsTran.AddNew
					oRsTran("FromAccountID")= Request.Form("txtUserAccID")
					oRsTran("ToAccountID")= Request.Form("txtShopAccID")
					oRsTran("TranscationTypeID")= 3 '"--> Online Purchase
					oRsTran("Date")= date
					oRsTran("Amount")= Request.Form("txtTotal")
on error resume next
					oRsTran.Update

					'<comment these line>
						'Response.Write "<br>Before DB CheckError.asp"
						'Response.End
					'</comment these line>

					%>
					<!-- #include file="CheckError.asp" -->
					<%

					'<comment these line>
						'Response.Write "<br>After DB CheckError.asp"
						'Response.End
					'</comment these line>

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'i think this line should not be included, so leave it as comment
					'<un-comment these lines>
						'oRsTran.Close
					'</un-comment these lines>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

					if len(Session("Msg")) > 0 then

						'<comment these line>
							'Response.Write "<br>After DB Err Msg"
							'Response.End
						'</comment these line>

						Session("Msg") = Session("Msg") & " ."

						'<un-comment thers lines>
							Response.Redirect "frmPurshase.asp"
						'</un-comment these lines>
					end if

					'<comment these line>
						'Response.Write "<br>No DB Err Msg"
						'Response.End
					'</comment these line>

					'<un-comment these lines>
						oRsTran.MoveLast
					'</un-comment these lines>

					nTranID = oRsTran("ID")

					'<un-comment these lines>
						oRsTran.Close
					'</un-comment these lines>

					'end-Insert Into Transaction

					'for Tran Charges
					'<un-comment these lines>
						oRsTran.Open "Transaction", oConPurchase, adOpenDynamic ,adLockPessimistic
						orsTran.AddNew
						oRsTran("FromAccountID")= Request.Form("txtUserAccID") 'Request.Form("txtFromAccID")
						oRsTran("ToAccountID")= 38 ' a bank deduction account #
						oRsTran("TranscationTypeID")= 5 '--> bank deduction
						oRsTran("Date")= date
						oRsTran("Amount")= nTranCharges
						oRsTran.Update
					'</un-comment these lines>

					%>
					<!-- #include file="CheckError.asp" -->
					<%

					'<un-comment these lines>
						oRsTran.Close
					'</un-comment these lines>

					if len(Session("Msg")) > 0 then

						'<comment these line>
							'Response.Write "<br>DB update Err Msg"
							'Response.Write "<br>Err = " & Session("Msg")
							'Response.End
						'</comment these line>

						Session("Msg") = Session("Msg") & " ."
						Response.Redirect "frmPurshase.asp"
					end if

					'<comment these line>
						'Response.Write "<br>No DB update Err Msg"
						'Response.End
					'</comment these line>

					'end- for Tran Charges
					'end - pwd ok code
				end if
			else
				Response.Write "<br>False Password<br>"
				Response.Write "<br>If you want to re-enter it click back button on toolbar<br>"
				'Response.End
			end if
		else
			Response.Write "<br>False Account Login Name<br>"
			Response.Write "<br>If you want to re-enter it click back button on toolbar<br>"
			'Response.End
		end if
	else
		Response.Write "<br>False Account Number<br>"
		Response.Write "<br>If you want to re-enter it click back button twice on toolbar<br>"
		'Response.End
	end if
	oRs.Close
end if
'end-confirm pin/pwd
'Response.Write "<br>All ok"
'Response.End
%>
<FORM id=frmPur name=frmPu action="<%=Request.Form("txtShopURL")%>" method="post"><br>
<INPUT type=hidden name=txtOrderID value=<%=Request.Form("txtOrderID")%>>
<INPUT type=hidden name=txtConfirm value=<%=nConfirm%>>
<INPUT type=hidden name=txtTranID value=<%=nTranID%>>
<hr>Now click "Go to Shop" button to confirm order<hr>
<INPUT type="submit" value="Go to Shop" name=submit1>
</FORM>
</BODY>
</HTML>
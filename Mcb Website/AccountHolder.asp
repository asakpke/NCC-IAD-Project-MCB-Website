<!-- #include file="AuthenticateAccHolder.asp" -->
<%'<!-- #include file="UserInfo.asp" -->%>
<!-- #include file="Msg.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
</HEAD>
<BODY background="b.jpg">
<%
'Response.Write "<br>nCustTypeID = " & nCustTypeID
'Response.Write "<br>nCustTypeID = " & Login.nCustTypeID  
%>
<!--#include file="MainNavigating.htm"-->
<STRONG>Account Holder</STRONG><br>
User ID = <%=Session("UserID")%>
<UL style="MARGIN-LEFT: 15px">
  <LI><A href="frmReferTo.asp" target=b>Refer others</A> </LI>
  <LI><A href="ListReferedPerson.asp" target=b>Refered Person</A> </LI>
  <LI><A href="ListAccount.asp" target=b>List All Account</A> </LI>
  <LI><A href="AccountBalance.asp" target=b>Check Balance</A> </LI>
  <LI><A href="frmSelectAccountBS.asp" target=b>Bank Statement</A> </LI>
  <LI><A href="frmSelectAccountBsDate.asp" target=b>Bank Statement Between Dates</A> </LI>
  <LI><A href="frmTransfer.asp" target=b>Money Transfer</A> </LI>
  <LI><A href="frmTransferBank.asp" target=b>Money Transfer to other banks</A> </LI>
  <LI><A href="frmBill.asp" target=b>Utilitity Bill Payments</A> </LI>
  <LI><A href="frmNewAccount.asp" target=b>New Account</A> </LI>
  <LI><A href="frmEditing.asp" target=b>Editing</A> </LI></UL>
<HR>
</BODY>
</HTML>

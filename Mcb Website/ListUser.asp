<!-- #include file="AuthenticateAdmin.asp" -->
<!-- # include file="NavigatingAdmin.htm" -->
<BODY background="b.jpg">
<font face="arial" color="green"><h2>
Administration &rarr; <FONT size=7>L</FONT>ist of Users</h2></font>
<hr>
<%
	set oRs = server.CreateObject("ADODB.recordset")	
	dim sSQL
	'sSQL="SELECT UserType.Type, User.ID, User.Name, User.Address, User.Phone, " _
		'& "User.LoginName FROM UserType INNER JOIN [User] ON UserType.ID = " _
		'& "User.UserTypeID ORDER BY User.ID;"
	sSQL = "SELECT UserType.Type, UserState.State, User.ID, User.Name, User.Address, User.Phone, User.LoginName " _
			 & "FROM UserState INNER JOIN (UserType INNER JOIN [User] ON UserType.ID = " _
			 & "User.UserTypeID) ON UserState.ID = User.UserStateID " _
			 & "ORDER BY User.ID;"

		
	oRs.Open sSQL, "DSN=dsnMCB"

	'Response.Write "<HR><H3>List of User</H3><HR>"
	Response.Write "<TABLE BORDER=1>"
	Response.Write "<TR><TH>User Type</TH><TH>User State</TH><TH>ID</TH><TH>Name</TH><TH>" _
		& "Address</TH><TH>Phone</TH><TH>Loin Name</TH><TR><TD>"
	Response.Write oRs.GetString(,,"</TD><TD>","</TD></TR><TR><TD>","&nbsp;")
	Response.Write "</TD></TR></TABLE>"	
	'while not oRs.EOF
	'	Response.Write oRs("Type")& "&nbsp;</TD><TD>"
	'	Response.Write oRs("ID")& "&nbsp;</TD><TD>"
	'	Response.Write oRs("Name") & "&nbsp;</TD><TD>"
	'	Response.Write oRs("Address") & "&nbsp;</TD><TD>"
	'	Response.Write oRs("Phone") & "&nbsp;</TD><TD>"
	'	Response.Write oRs("LoginName") & "&nbsp;</TD></TR><TR><TD>"
	'	oRs.MoveNext 
	'wend
	'Response.Write "</TD></TR></TABLE>"
	
		'ors.Update
	'end if
	oRs.Close 
%>

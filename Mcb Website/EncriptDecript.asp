<%@ Language=VBScript %>
<%
	Dim OtherArr(3) 
	Randomize 
	TempString = ""
	For i=0 To 3 
		OtherArr(i) = Int((9 * Rnd)) 
		TempString = TempString & OtherArr(i)
	Next 		
	'msgbox TempString
	ThisText = "A quick brown fox jumps over the lazy dog."
	Response.Write "<hr><h1>Original Form</h1>" & ThisText
	For i=1 To Len(ThisText) 
		TempNum = Asc(Mid(ThisText,i,1)) 
		If TempNum = 13 Then 
			TempNum = 28 
		ElseIf TempNum = 10 Then 
			TempNum = 29 
		End If 
		TempChar = Chr(TempNum - OtherArr(i Mod 4)) 
		If TempChar = Chr(34) Then 
			TempChar = Chr(18) 
		End If 
		TempString = TempString & TempChar 
	Next 
	Response.write "<hr><h1>Encripted Form</h1>" & TempString
	

	ThisText = ""
	ExeString = TempString
	UnLockStr = "Execute(""Dim KeyArr(3),ThisText""&vbCrLf&""KeyArr(0) = " & Cint(Mid(ExeString,1,1)) & """&vbCrLf&""KeyArr(1) = " & Cint(Mid(ExeString,2,1)) & """&vbCrLf&""KeyArr(2) = " & Cint(Mid(ExeString,3,1)) & """&vbCrLf&""KeyArr(3) = " & Cint(Mid(ExeString,4,1)) & """&vbCrLf&""For i=5 To Len(ExeString)""&vbCrLf&""TempNum = Asc(Mid(ExeString,i,1))""&vbCrLf&""If TempNum = 18 Then""&vbCrLf&""TempNum = 34""&vbCrLf&""End If""&vbCrLf&""TempChar = Chr(TempNum + KeyArr(i Mod 4))""&vbCrLf&""If TempChar = Chr(28) Then""&vbCrLf&""TempChar = vbCr""&vbCrLf&""ElseIf TempChar = Chr(29) Then""&vbCrLf&""TempChar = vbLf""&vbCrLf&""End If""&vbCrLf&""ThisText = ThisText & TempChar""&vbCrLf&""Next"")" '& vbCrLf & "Execute(ThisText)" 
	Execute UnLockStr
	Response.write "<hr><h1>Decripted Form</h1>" & ThisText	
%>
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDll6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Encript(ThisText As String) As String
  'fnWelcome = "Welcome to My Dll"
  'ssssss
  If ThisText = "" Then
    Encript = ""
    Exit Function
  End If
  Dim OtherArr(3) As Byte
  Randomize
  Dim TempString
  TempString = ""
  Dim I As Byte
  For I = 0 To 3
    OtherArr(I) = Int((9 * Rnd))
    TempString = TempString & OtherArr(I)
  Next
  'MsgBox "Code = " & TempString
  'Dim ThisText As String
  'ThisText = "A quick brown fox jumps over the lazy dog."
 ' Response.write "<hr><h1>Original Form</h1>" & ThisText
 'MsgBox "Original Form = " & ThisText
  Dim TempNum As Byte
  Dim TempChar As String
  For I = 1 To Len(ThisText)
    TempNum = Asc(Mid(ThisText, I, 1))
    If TempNum = 13 Then
      TempNum = 28
    ElseIf TempNum = 10 Then
      TempNum = 29
    End If
    TempChar = Chr(TempNum - OtherArr(I Mod 4))
    If TempChar = Chr(34) Then
      TempChar = Chr(18)
    End If
    TempString = TempString & TempChar
  Next
  'MsgBox "Enc Form = " & TempString
  Encript = TempString
  'Response.write "<hr><h1>Encripted Form</h1>" & TempString
  'eeeeee
End Function
Public Function Decript(TempString As String) As String
  'desr
  'ThisText = ""
  Dim ExeString As String
  ExeString = TempString
  If IsNumeric(Left(TempString, 4)) = False Then
    Decript = TempString
    Exit Function
  End If
  Dim KeyArr(3) As Byte, ThisText As String
  KeyArr(0) = CInt(Mid(ExeString, 1, 1))
  KeyArr(1) = CInt(Mid(ExeString, 2, 1))
  KeyArr(2) = CInt(Mid(ExeString, 3, 1))
  KeyArr(3) = CInt(Mid(ExeString, 4, 1))
  Dim I As Integer
  Dim TempNum As Integer
  Dim TempChar As String
  For I = 5 To Len(ExeString)
    TempNum = Asc(Mid(ExeString, I, 1))
    If TempNum = 18 Then
      TempNum = 34
    End If
    TempChar = Chr(TempNum + KeyArr(I Mod 4))
    If TempChar = Chr(28) Then
      TempChar = vbCr
    ElseIf TempChar = Chr(29) Then
      TempChar = vbLf
    End If
    ThisText = ThisText & TempChar
  Next
  'UnlockStr
  'MsgBox "Decr Form = " & ThisText
  Decript = ThisText
'  Response.write "<hr><h1>Decripted Form</h1>" & ThisText
End Function


Public Function fnWelcome() As String
  fnWelcome = "Welcome to my dll"
End Function

Public Function fnWelcomeTxt(sStr As String) As String
  fnWelcomeTxt = sStr
End Function


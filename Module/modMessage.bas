Attribute VB_Name = "modMessage"
Option Explicit

Public Function MsgQtnYN(ByVal prompt As String) As Boolean
 If MsgBox(prompt, vbQuestion + vbYesNo, "Confirm") = vbYes Then
  MsgQtnYN = True
 End If
End Function

Public Function MsgWhole() As Boolean
 MsgBox "Sorry, You have to fill entire/all required information !" & vbNewLine & _
        "Please check again !", vbInformation, "Information"
End Function

Public Function MsgSaveScd(ByVal strFileName As String)
 MsgBox "Data have succeeded saved to " & strFileName & vbNewLine & _
        "Saving Suceeded", vbInformation, _
        "Information"
End Function

Public Function MsgQueryOK() As Boolean
 MsgBox "Syntax OK !", vbInformation, "Information"
End Function

Public Function MsgQueryErr() As Boolean
 MsgBox "Wrong Syntax !" & vbNewLine & _
        "Please Check Again Your SQL Syntax", vbExclamation, "Information"
End Function

Public Function MsgConScd()
 MsgBox "Connection successful" & vbNewLine & _
         "(Connected)", vbInformation, _
         "Connection Info"
End Function

Public Function MsgConAccesErr()
 MsgBox "Connection failed (Disconnected)" & vbNewLine & _
         "Please check the following information : " & vbNewLine & vbNewLine & _
         "* Make sure the Data Name is correct." & vbNewLine & _
         "* Make sure the Data Path is correct." & vbNewLine & _
         "* Make sure UserID is correct (if UserID is default, use blank UserID)." & vbNewLine & _
         "* Make sure Password is correct (if UserID is default, use blank Password)." & vbNewLine & _
         "* Please check file is exists.", _
         vbExclamation, _
         "Connection Info"
End Function

Public Function MsgConServerErr()
 MsgBox "Connection failed (Disconnected)" & vbNewLine & _
         "Please check the following information : " & vbNewLine & vbNewLine & _
         "* Make sure the Server Name is correct." & vbNewLine & _
         "* Make sure the Database Name is correct." & vbNewLine & _
         "* Make sure UserID is correct (if UserID is default, use ""sa"")." & vbNewLine & _
         "* Make sure Password is correct (if UserID is default, use blank Password)." & vbNewLine & _
         "* Please check your Network.", _
         vbExclamation, _
         "Connection Info"
End Function

Public Function MsgCantResume(ByVal strMustSet As String) As Boolean
 MsgBox "Sorry, You can't use this proccess at this time !" & vbNewLine & _
        "Please set " & strMustSet & " first.", vbInformation, "Information"
End Function

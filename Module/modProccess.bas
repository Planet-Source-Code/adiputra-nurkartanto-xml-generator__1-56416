Attribute VB_Name = "modProccess"
Option Explicit

Public frmXMLGen As frmXML

Public strQuery As String
Public intCount As Integer

Public Sub NewXMLgen()
 Set frmXMLGen = New frmXML
  
 intCount = intCount + 1
 frmXMLGen.strMyOwnNumber = "XML" & intCount
 frmXMLGen.Caption = "XML GENERATOR <-" & intCount & "->"
 frmXMLGen.Show
 
 '[Add To Collection]
 objfrmXML.SetPropertyNull
 objfrmXML.frmNumber = frmXMLGen.strMyOwnNumber
 objfrmXML.PathTempFile = PathTemp & "~convert" & frmXMLGen.strMyOwnNumber & ".tmp"
 objfrmXML.AddList
 objfrmXML.SetPropertyNull
End Sub

Public Function PathApp() As String
 Dim myPath As String
 
 myPath = App.Path
 If Right(myPath, 1) <> "\" Then
  myPath = myPath & "\"
 End If
 
 PathApp = myPath
End Function

Public Function PathTemp() As String
 Dim myPath As String
 
 myPath = PathApp & "Temp\"
 
 PathTemp = myPath
End Function

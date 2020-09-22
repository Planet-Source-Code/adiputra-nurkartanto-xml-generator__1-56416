VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Converter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public myCon As ADODB.Connection
Public myRS As ADODB.Recordset

Private mvarQueryString As String
Private mvarmyTableName As String
Private mvarDataName As String

Public Property Let DataName(ByVal vData As String)
    mvarDataName = vData
End Property

Public Property Get DataName() As String
    DataName = mvarDataName
End Property

Public Property Let myTableName(ByVal vData As String)
    mvarmyTableName = vData
End Property

Public Property Get myTableName() As String
    myTableName = mvarmyTableName
End Property

Public Property Let QueryString(ByVal vData As String)
    mvarQueryString = vData
End Property

Public Property Get QueryString() As String
    QueryString = mvarQueryString
End Property

Public Function WriteXML(ByVal FileTemp As String) As Boolean
 Dim lngRecordCount As Long
 Dim iFieldsCount As Integer
 Dim a As Integer
   
 Set myRS = New ADODB.Recordset
 myRS.Open QueryString, myCon, adOpenForwardOnly
 
 '[Write To Temporrary File]
 Open FileTemp For Output As #1
  Print #1, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & "?>"
  
  Print #1, "<!-- Created With XML GENERATOR -->"
  Print #1, "<!--   [e-Volve  Open Source]   -->"
  
  Print #1, "<" & UCase(DataName) & ">"
  
  Print #1, "   <" & myTableName & ">"
  
  If Not myRS.BOF And Not myRS.EOF Then
   lngRecordCount = 0
   Do While Not myRS.EOF
    iFieldsCount = myRS.Fields.Count
    lngRecordCount = lngRecordCount + 1
    
    Print #1, "      <RECORD_" & lngRecordCount & ">"
    
    For a = 0 To iFieldsCount - 1
     With myRS.Fields(a)
      Print #1, "           <" & .Name & ">"
      Print #1, "               " & .Value
      Print #1, "           </" & .Name & ">"
     End With
    Next
    
    Print #1, "      </RECORD_" & lngRecordCount & ">"
    
    myRS.MoveNext
   Loop
  End If
    
  Print #1, "   </" & myTableName & ">"
  Print #1, "</" & UCase(DataName) & ">"
 Close #1
 
 myRS.Close
 Set myRS = Nothing
 
 WriteXML = True
End Function

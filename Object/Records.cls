VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Records"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum OpenRecord
 OpenIsData = 0
 OpenNoData = 1
 OpenFalse = 3
 OpenClosed = 4
End Enum

Private mvarQueryString As String
Private mvarFieldName As String
Private mvarStatOpen As OpenRecord

Public myCon As ADODB.Connection
Public myRS As ADODB.Recordset
Public ListFields As RecordsFieldList

Public Property Let FieldName(ByVal vData As String)
    mvarFieldName = vData
End Property

Public Property Get FieldName() As String
    FieldName = mvarFieldName
End Property

Public Property Let StatOpen(ByVal vData As OpenRecord)
    mvarStatOpen = vData
End Property

Public Property Get StatOpen() As OpenRecord
    StatOpen = mvarStatOpen
End Property

Public Property Let QueryString(ByVal vData As String)
    mvarQueryString = vData
End Property

Public Property Get QueryString() As String
    QueryString = mvarQueryString
End Property

Public Sub myOpen()
 On Error GoTo ErrHandle

 Dim iFieldsCount As Integer
 Dim a As Integer

 myRS.Open QueryString, myCon, adOpenForwardOnly

 iFieldsCount = myRS.Fields.Count

 If iFieldsCount > 0 Then
  ListFields.Clear
  For a = 0 To iFieldsCount - 1
   ListFields.Add myRS.Fields(a).Name, myRS.Fields(a).Name
  Next
 End If

 If Not myRS.BOF And Not myRS.EOF Then
  StatOpen = OpenIsData
 ElseIf myRS.BOF And myRS.EOF Then
  StatOpen = OpenNoData
 End If

 Exit Sub

ErrHandle:
 If Err.Number = -2147217900 Then
  StatOpen = OpenFalse
  Exit Sub
 End If
End Sub

Public Sub myClose()
 myRS.Close
 StatOpen = OpenClosed
End Sub

Private Sub Class_Initialize()
 Set ListFields = New RecordsFieldList
 Set myRS = New ADODB.Recordset
 StatOpen = OpenClosed              '[Set Default Value]
End Sub

Private Sub Class_Terminate()
 Set ListFields = Nothing
 Set myRS = Nothing
End Sub

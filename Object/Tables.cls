VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public myCon As ADODB.Connection
Public myCatalog As ADOX.Catalog

Private mvarTableName As String
Public ListTables As TablesList

Public Property Let TableName(ByVal vData As String)
    mvarTableName = vData
End Property

Public Property Get TableName() As String
    TableName = mvarTableName
End Property

Public Function GetTableName() As Boolean
 Dim a As Integer
 Dim iTableCount As Integer
  
 myCatalog.ActiveConnection = myCon
 iTableCount = myCatalog.Tables.Count
 
 If iTableCount > 0 Then
  ListTables.Clear
  For a = 0 To iTableCount - 1
   If myCatalog.Tables(a).Type = "TABLE" Then
    iTableCount = iTableCount + 1
    ListTables.Add myCatalog.Tables(a).Name, myCatalog.Tables(a).Name
   End If
  Next
  GetTableName = True
 End If
End Function

Private Sub Class_Initialize()
 Set myCatalog = New ADOX.Catalog
 Set ListTables = New TablesList
End Sub

Private Sub Class_Terminate()
 Set ListTables = Nothing
 Set myCatalog = Nothing
End Sub

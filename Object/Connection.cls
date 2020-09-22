VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Connection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarmyNumber As String
Private mvarDataPath As String
Private mvarUserID As String
Private mvarPassword As String
Private mvarDataName As String
Private mvarServerName As String

Public myCon As ADODB.Connection

Public Property Let DataName(ByVal vData As String)
    mvarDataName = vData
End Property

Public Property Get DataName() As String
    DataName = mvarDataName
End Property

Public Property Let ServerName(ByVal vData As String)
    mvarServerName = vData
End Property

Public Property Get ServerName() As String
    ServerName = mvarServerName
End Property

Public Property Let Password(ByVal vData As String)
    mvarPassword = vData
End Property

Public Property Get Password() As String
    Password = mvarPassword
End Property

Public Property Let UserID(ByVal vData As String)
    mvarUserID = vData
End Property

Public Property Get UserID() As String
    UserID = mvarUserID
End Property

Public Property Let DataPath(ByVal vData As String)
    mvarDataPath = vData
End Property

Public Property Get DataPath() As String
    DataPath = mvarDataPath
End Property

Public Property Let myNumber(ByVal vData As String)
    mvarmyNumber = vData
End Property

Public Property Get myNumber() As String
    myNumber = mvarmyNumber
End Property

Public Function ConnectSQLServer() As Boolean
 On Error GoTo ErrHandle
 Dim strCon As String
 
 strCon = "Provider=SQLOLEDB.1" & _
          ";Persist Security Info=False" & _
          ";User ID=" & UserID & _
          ";Password=" & Password & _
          ";Initial Catalog=" & DataName & _
          ";Data Source=" & ServerName
 
 Set myCon = New ADODB.Connection
 myCon.Open strCon
 
 ConnectSQLServer = True
 
 Exit Function
 
ErrHandle:
 ConnectSQLServer = False
 Exit Function
End Function

Public Function ConnectAccess() As Boolean
 On Error GoTo ErrHandle
 Dim strCon As String
 
 strCon = "Provider=Microsoft.Jet.OLEDB.4.0" & _
          ";Data Source=" & DataPath & _
          ";Persist Security Info=False"
 
 Set myCon = New ADODB.Connection
 myCon.Open strCon, UserID, Password
 
 ConnectAccess = True
 
 Exit Function
 
ErrHandle:
 ConnectAccess = False
 Exit Function
End Function

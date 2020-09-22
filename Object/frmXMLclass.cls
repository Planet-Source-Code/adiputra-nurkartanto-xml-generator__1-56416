VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "frmXMLclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarfrmNumber As String
Private mvarQueryString As String
Private mvarDataPath As String
Private mvarUserID As String
Private mvarPassword As String
Private mvarDataName As String
Private mvarServerName As String
Private mvarIsOpened As Boolean
Private mvarmyTableName As String
Private mvarPathTempFile As String

Public myCon As ADODB.Connection
Public ListfrmXML As frmXMLList

Public Property Let PathTempFile(ByVal vData As String)
    mvarPathTempFile = vData
End Property

Public Property Get PathTempFile() As String
    PathTempFile = mvarPathTempFile
End Property

Public Property Let myTableName(ByVal vData As String)
    mvarmyTableName = vData
End Property

Public Property Get myTableName() As String
    myTableName = mvarmyTableName
End Property

Public Property Let IsOpened(ByVal vData As Boolean)
    mvarIsOpened = vData
End Property

Public Property Get IsOpened() As Boolean
    IsOpened = mvarIsOpened
End Property

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

Public Property Let QueryString(ByVal vData As String)
    mvarQueryString = vData
End Property

Public Property Get QueryString() As String
    QueryString = mvarQueryString
End Property

Public Property Let frmNumber(ByVal vData As String)
    mvarfrmNumber = vData
End Property

Public Property Get frmNumber() As String
    frmNumber = mvarfrmNumber
End Property

Public Function GetData_ByNo(ByVal No As String) As Boolean
 Dim myItem As frmXMLclass
 For Each myItem In ListfrmXML
  If LCase(No) = LCase(myItem.frmNumber) Then
   frmNumber = myItem.frmNumber
   QueryString = myItem.QueryString
   DataPath = myItem.DataPath
   UserID = myItem.UserID
   Password = myItem.Password
   ServerName = myItem.ServerName
   DataName = myItem.DataName
   IsOpened = myItem.IsOpened
   PathTempFile = myItem.PathTempFile
   myTableName = myItem.myTableName
   Set myCon = myItem.myCon
   
   GetData_ByNo = True
  End If
 Next
End Function

Public Function SetData_ByNo(ByVal No As String) As Boolean
 Dim myItem As frmXMLclass
 For Each myItem In ListfrmXML
  If LCase(myItem.frmNumber) = LCase(No) Then
   myItem.frmNumber = frmNumber
   myItem.QueryString = QueryString
   myItem.DataPath = DataPath
   myItem.UserID = UserID
   myItem.Password = Password
   myItem.ServerName = ServerName
   myItem.DataName = DataName
   myItem.IsOpened = IsOpened
   myItem.myTableName = myTableName
   myItem.PathTempFile = PathTempFile
   Set myItem.myCon = myCon
   
   SetData_ByNo = True
  End If
 Next
End Function

Public Sub AddList()
 ListfrmXML.Add PathTempFile, myTableName, IsOpened, frmNumber, QueryString, DataPath, UserID, Password, myCon, frmNumber
End Sub

Public Sub SetPropertyNull()
 frmNumber = ""
 QueryString = ""
 DataPath = ""
 UserID = ""
 Password = ""
 DataName = ""
 ServerName = ""
 PathTempFile = ""
 myTableName = ""
 IsOpened = False
 Set myCon = Nothing
End Sub

Private Sub Class_Initialize()
 Set ListfrmXML = New frmXMLList
End Sub

Private Sub Class_Terminate()
 Set ListfrmXML = Nothing
End Sub

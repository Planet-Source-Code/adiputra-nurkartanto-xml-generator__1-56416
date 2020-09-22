VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "frmXMLList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"frmXMLclass"
Attribute VB_Ext_KEY = "Member0" ,"frmXMLclass"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Friend Sub Clear()
 On Error Resume Next
 Dim a As Integer
 For a = mCol.Count To 1 Step -1
  mCol.Remove a
 Next
End Sub


Public Function Add(PathTempFile As String, myTableName As String, IsOpened As Boolean, frmNumber As String, QueryString As String, DataPath As String, UserID As String, Password As String, myCon As ADODB.Connection, Optional sKey As String) As frmXMLclass
    'create a new object
    Dim objNewMember As frmXMLclass
    Set objNewMember = New frmXMLclass


    'set the properties passed into the method
    objNewMember.PathTempFile = PathTempFile
    objNewMember.myTableName = myTableName
    objNewMember.IsOpened = IsOpened
    objNewMember.frmNumber = frmNumber
    objNewMember.QueryString = QueryString
    objNewMember.DataPath = DataPath
    objNewMember.UserID = UserID
    objNewMember.Password = Password
    Set objNewMember.myCon = myCon
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As frmXMLclass
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

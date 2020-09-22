VERSION 5.00
Begin VB.Form frmXMLQuery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Query"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmXMLQuery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7095
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame freStrip 
      Height          =   135
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   7095
   End
   Begin VB.TextBox txtTable 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
   End
   Begin VB.ComboBox cboTable 
      Height          =   315
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox txtSQL 
      Appearance      =   0  'Flat
      Height          =   2055
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2040
      Width           =   7095
   End
   Begin VB.Frame freStrip 
      Height          =   135
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   4200
      Width           =   7095
   End
   Begin VB.Label LABEL 
      AutoSize        =   -1  'True
      Caption         =   "(Optional)"
      Height          =   195
      Index           =   2
      Left            =   5280
      TabIndex        =   10
      Top             =   1005
      Width           =   855
   End
   Begin VB.Label LABEL 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Table Name [Query]"
      Height          =   195
      Index           =   1
      Left            =   720
      TabIndex        =   8
      Top             =   1440
      Width           =   1755
   End
   Begin VB.Label LABEL 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Choose From Single Table"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1005
      Width           =   2250
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmXMLQuery.frx":0CCA
      ForeColor       =   &H80000017&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6135
   End
   Begin VB.Image imgInfo 
      Height          =   480
      Left            =   6360
      Picture         =   "frmXMLQuery.frx":0D89
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape shpBack 
      BackColor       =   &H80000018&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "frmXMLQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strMyOwnNumber As String

'[Object]
Dim objRecord As xGenDLL.Records
Dim objTable As xGenDLL.Tables

Private Sub cboTable_Click()
 If cboTable.ListIndex > 0 Then
  txtSQL.Text = "SELECT * FROM " & cboTable.Text
  txtTable.Text = cboTable.Text
 End If
End Sub

Private Sub cmdCancel_Click()
 Unload Me
End Sub

'[FORM VALIDATION] ===============================================================
'
'<verify form>
Private Sub verify_form()
 If Len(Trim(txtSQL.Text)) = 0 Then
  MsgWhole
  txtSQL.SetFocus
  Exit Sub
 End If
 
 If Len(Trim(txtTable.Text)) = 0 Then
  MsgWhole
  txtTable.SetFocus
  Exit Sub
 End If
 
 If CekQuery Then   '[Function pos : bellow]
  MsgQueryOK
  Unload Me
 Else
  MsgQueryErr
  txtSQL.SetFocus
  SendKeys "{Home}+{End}"
 End If
End Sub
'</verify form>

'<Query Checker>
Private Function CekQuery() As Boolean
 If objfrmXML.GetData_ByNo(strMyOwnNumber) Then
  Set objRecord = New xGenDLL.Records
  Set objRecord.myCon = objfrmXML.myCon
 
  objRecord.QueryString = txtSQL.Text
  objRecord.myOpen
  
  If objRecord.StatOpen = OpenIsData Or objRecord.StatOpen = OpenNoData Then
   objfrmXML.QueryString = txtSQL.Text
   objfrmXML.myTableName = UCase(txtTable.Text)
   objRecord.myClose
   
   CekQuery = objfrmXML.SetData_ByNo(strMyOwnNumber)
  End If
  Set objRecord = Nothing
 End If
End Function
'</Query Checker>
'
'[/FORM VALIDATION] ==============================================================

Private Sub cmdOK_Click()
 verify_form    '[SubRutin pos : above]
End Sub

'[FORM INITIALIZE] ===============================================================
'
'<Show Table>
Private Sub GetTable_SetTable()
 Dim a As Integer
 
 cboTable.Clear
 cboTable.AddItem "[-Select Single table-]"
 
 If objfrmXML.GetData_ByNo(strMyOwnNumber) Then
  Set objTable = New xGenDLL.Tables
  Set objTable.myCon = objfrmXML.myCon
  
  If objTable.GetTableName Then
   For a = 1 To objTable.ListTables.Count
    cboTable.AddItem objTable.ListTables(a).TableName
   Next
  End If
  
  Set objTable = Nothing
 End If
 
 cboTable.ListIndex = 0
End Sub
'</Show Table>
'
'[FORM INITIALIZE] ===============================================================

Private Sub Form_Load()
 GetTable_SetTable  '[SubRutin pos : above]
End Sub

Private Sub txtSQL_GotFocus()
 SendKeys "{Home}+{End}"
End Sub

Private Sub txtTable_GotFocus()
 SendKeys "{Home}+{End}"
End Sub

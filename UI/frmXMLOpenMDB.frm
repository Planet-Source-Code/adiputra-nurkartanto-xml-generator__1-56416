VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmXMLOpenMDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open Microsoft Access Database"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmXMLOpenMDB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtIn 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtIn 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox txtIn 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   3
      Top             =   1800
      Width           =   2895
   End
   Begin VB.CheckBox chkPassword 
      Caption         =   "Blank Password"
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   2160
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.Frame freStrip 
      Height          =   135
      Left            =   0
      TabIndex        =   7
      Top             =   2520
      Width           =   4695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   285
      Left            =   3600
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin MSComDlg.CommonDialog cdcDatabase 
      Left            =   120
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open File MS Access [*.MDB]"
      FontName        =   "verdana"
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "If your Database use User ID and Password, uncheck Blank Password and give information with correct UserID && Password"
      ForeColor       =   &H80000017&
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   3855
   End
   Begin VB.Image imgInfo 
      Height          =   480
      Left            =   4080
      Picture         =   "frmXMLOpenMDB.frx":0442
      Top             =   120
      Width           =   480
   End
   Begin VB.Label LABEL 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DB Name"
      Height          =   195
      Index           =   1
      Left            =   570
      TabIndex        =   10
      Top             =   960
      Width           =   810
   End
   Begin VB.Label LABEL 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      Height          =   195
      Index           =   2
      Left            =   720
      TabIndex        =   9
      Top             =   1440
      Width           =   660
   End
   Begin VB.Label LABEL 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Index           =   3
      Left            =   570
      TabIndex        =   8
      Top             =   1800
      Width           =   810
   End
   Begin VB.Shape shpBack 
      BackColor       =   &H80000018&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmXMLOpenMDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strMyOwnNumber As String

Dim strDBpath As String
Dim strDBname As String

'[object]
Dim objConn As xGenDLL.Connection

'[SET OPEN CONNECTION TO DATABASE ACCESS (*.MDB)] ================================
'
'<Set Information DB>
Private Function SetOpenDB() As Boolean
 If objfrmXML.GetData_ByNo(strMyOwnNumber) Then
  If Len(Trim(strDBpath)) <> 0 Then
   objfrmXML.DataName = strDBname
   objfrmXML.DataPath = strDBpath
   SetOpenDB = True
  Else
   objfrmXML.isOpened = False
  End If
 End If
End Function
'</Set Information DB>

'<Set Connection DB>
Private Function SetConnDB() As Boolean
 Set objConn = New xGenDLL.Connection
 objConn.DataName = strDBname
 objConn.DataPath = strDBpath
 objConn.UserID = txtIn(1).Text
 objConn.Password = txtIn(2).Text
   
 If objConn.ConnectAccess Then
  Set objfrmXML.myCon = objConn.myCon
  objfrmXML.QueryString = ""
  objfrmXML.isOpened = True
  SetConnDB = True
 End If
 Set objConn = Nothing
End Function
'</Set Connection DB>

'<Verify Open Access (*.MDB)>
Private Sub verify_open()
 On Error GoTo ErrHandle
 
 With cdcDatabase
  .Flags = cdlOFNFileMustExist & cdlOFNHideReadOnly
  .CancelError = True
  .Filter = "MS Access DB Files|*.MDB|"
  .ShowOpen
  strDBpath = .FileName
  strDBname = .FileTitle
 End With
 
 txtIn(0).Text = strDBpath
 
 Exit Sub
 
ErrHandle:
 If Err.Number = 32755 Then
  Resume Next
 End If
End Sub
'</Verify Open Access (*.MDB)>
'
'[/SET OPEN CONNECTION TO DATABASE ACCESS (*.MDB)] ===============================

Private Sub chkPassword_Click()
 Dim a As Byte
 
 If chkPassword.Value = 1 Then
  For a = 1 To 2
   txtIn(a).Text = ""
   txtIn(a).BackColor = &HE0E0E0
   txtIn(a).Enabled = False
  Next
 Else
  For a = 1 To 2
   txtIn(a).Text = ""
   txtIn(a).BackColor = vbWhite
   txtIn(a).Enabled = True
  Next
  txtIn(1).SetFocus
 End If
End Sub

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdOK_Click()
 If SetOpenDB Then      '[Function pos : above]
  If SetConnDB Then     '[Function pos : above]
   If objfrmXML.SetData_ByNo(strMyOwnNumber) Then
    MsgConScd
    Unload Me
   Else
    MsgConAccesErr
   End If
  Else
   MsgConAccesErr
  End If
 Else
  MsgConAccesErr
 End If
End Sub

Private Sub cmdOpen_Click()
 verify_open            '[SubRutin pos : above]
End Sub

Private Sub txtIn_GotFocus(Index As Integer)
 SendKeys "{Home}+{End}"
End Sub

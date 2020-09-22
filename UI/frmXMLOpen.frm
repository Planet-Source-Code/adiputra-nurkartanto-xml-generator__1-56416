VERSION 5.00
Begin VB.Form frmXMLOpen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open SQL Server Database"
   ClientHeight    =   3720
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
   Icon            =   "frmXMLOpen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame freStrip 
      Height          =   135
      Left            =   0
      TabIndex        =   12
      Top             =   3000
      Width           =   4695
   End
   Begin VB.CheckBox chkPassword 
      Caption         =   "Blank Password"
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   2640
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.TextBox txtIn 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txtIn 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      Text            =   "sa"
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox txtIn 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox txtIn 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label LABEL 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Index           =   3
      Left            =   570
      TabIndex        =   10
      Top             =   2280
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
      Top             =   1920
      Width           =   660
   End
   Begin VB.Label LABEL 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DB Name"
      Height          =   195
      Index           =   1
      Left            =   570
      TabIndex        =   8
      Top             =   1320
      Width           =   810
   End
   Begin VB.Label LABEL 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server Name"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1140
   End
   Begin VB.Image imgInfo 
      Height          =   480
      Left            =   4080
      Picture         =   "frmXMLOpen.frx":0442
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "You can put Sarver Name or IP Server in the Server Name Box. Check the Blank Password to set default User ID."
      ForeColor       =   &H80000017&
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   3855
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
Attribute VB_Name = "frmXMLOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strMyOwnNumber As String

'[object]
Dim objConn As xGenDLL.Connection

Private Sub chkPassword_Click()
 Dim a As Byte
 
 If chkPassword.Value = 1 Then
  For a = 2 To 3
   txtIn(a).Enabled = False
   txtIn(a).Text = ""
   txtIn(a).BackColor = &HE0E0E0
  Next
  txtIn(2).Text = "sa"
 Else
  For a = 2 To 3
   txtIn(a).Enabled = True
   txtIn(a).Text = ""
   txtIn(a).BackColor = vbWhite
  Next
  txtIn(2).Text = "sa"
 End If
End Sub

Private Sub cmdCancel_Click()
 Unload Me
End Sub

'[SET OPEN CONNECTION TO SQL SERVER DATABASE] ====================================
'
'<Set Information SQL Server>
Private Function SetOpenDB() As Boolean
 If objfrmXML.GetData_ByNo(strMyOwnNumber) Then
  objfrmXML.ServerName = txtIn(0).Text
  objfrmXML.DataName = txtIn(1).Text
  objfrmXML.UserID = txtIn(2).Text
  objfrmXML.Password = txtIn(3).Text
  
  SetOpenDB = True
 End If
End Function
'</Set Information SQL Server>

'<Set Connect SQL Server>
Private Function SetConnDB() As Boolean
 Set objConn = New xGenDLL.Connection
 
 objConn.ServerName = txtIn(0).Text
 objConn.DataName = txtIn(1).Text
 objConn.UserID = txtIn(2).Text
 objConn.Password = txtIn(3).Text
 
 objfrmXML.isOpened = False
 If objConn.ConnectSQLServer Then
  Set objfrmXML.myCon = objConn.myCon
  objfrmXML.QueryString = ""
  objfrmXML.isOpened = True
  SetConnDB = True
 End If
 
 Set objConn = Nothing
End Function
'</Set Connect SQL Server>
'
'[/SET OPEN CONNECTION TO SQL SERVER DATABASE] ===================================

'[FORM VALIDATION] ===============================================================
'
'<Verify Form>
Private Sub verify_form()
 Dim a As Integer
 
 For a = 0 To 1
  If Len(Trim(txtIn(a).Text)) = 0 Then
   MsgWhole
   txtIn(a).SetFocus
   Exit Sub
  End If
 Next
 
 If SetOpenDB Then      '[Function pos : above]
  If SetConnDB Then     '[Function pos : above]
   If objfrmXML.SetData_ByNo(strMyOwnNumber) Then
    MsgConScd
    Unload Me
   Else
    MsgConServerErr
   End If
  Else
   MsgConServerErr
  End If
 Else
  MsgConServerErr
 End If
End Sub
'</Verify Form>
'
'[/FORM VALIDATION] ===============================================================

Private Sub cmdOK_Click()
 verify_form    '[SubRutin pos : above]
End Sub

Private Sub txtIn_GotFocus(Index As Integer)
 SendKeys "{Home}+{End}"
End Sub

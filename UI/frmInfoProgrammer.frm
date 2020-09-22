VERSION 5.00
Begin VB.Form frmInfoProgrammer 
   BackColor       =   &H00CCBAB7&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Programmer"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4755
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInfoProgrammer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   4755
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Frame freProgrammer 
      BackColor       =   &H00CCBAB7&
      Height          =   2655
      Left            =   20
      TabIndex        =   0
      Top             =   2640
      Width           =   4695
      Begin VB.Label LABEL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "eVOLVE .:: garage team ::."
         Height          =   195
         Left            =   2160
         TabIndex        =   10
         Top             =   2400
         Width           =   2370
      End
      Begin VB.Label lblProgrammer 
         BackStyle       =   0  'Transparent
         Caption         =   "Birth"
         Height          =   195
         Index           =   7
         Left            =   480
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblProgrammer 
         BackStyle       =   0  'Transparent
         Caption         =   "08 January 1984"
         Height          =   195
         Index           =   6
         Left            =   1440
         TabIndex        =   8
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label lblProgrammer 
         BackStyle       =   0  'Transparent
         Caption         =   "adiputra_n@yahoo.com"
         Height          =   195
         Index           =   5
         Left            =   1440
         TabIndex        =   7
         Top             =   1800
         Width           =   3135
      End
      Begin VB.Label lblProgrammer 
         BackStyle       =   0  'Transparent
         Caption         =   "[See Code]"
         Height          =   735
         Index           =   4
         Left            =   1440
         TabIndex        =   6
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label lblProgrammer 
         BackStyle       =   0  'Transparent
         Caption         =   "Adiputra Nurkartanto"
         Height          =   195
         Index           =   3
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblProgrammer 
         BackStyle       =   0  'Transparent
         Caption         =   "email"
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   4
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblProgrammer 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblProgrammer 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Image imgMe 
      BorderStyle     =   1  'Fixed Single
      Height          =   2610
      Left            =   0
      Picture         =   "frmInfoProgrammer.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4770
   End
End
Attribute VB_Name = "frmInfoProgrammer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strAddress As String

Private Function Info_Address() As String
 strAddress = "Jl. Malabar Ujung No.120" & vbNewLine
 strAddress = strAddress & "Code Pos 16144" & vbNewLine
 strAddress = strAddress & "Bogor - West Java - Indonesia"
  
 Info_Address = strAddress
End Function

Private Sub cmdOK_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 lblProgrammer(4).Caption = Info_Address
End Sub


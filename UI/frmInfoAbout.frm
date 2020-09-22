VERSION 5.00
Begin VB.Form frmInfoAbout 
   BackColor       =   &H00CCBAB7&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About X-GEN"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInfoAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9510
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   8160
      TabIndex        =   1
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Frame freAbout 
      BackColor       =   &H00CCBAB7&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   9255
      Begin VB.Label lblAbout 
         BackStyle       =   0  'Transparent
         Caption         =   "[About Will Appear Here]"
         Height          =   1335
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   9015
      End
   End
   Begin VB.Image imgXMLgen 
      BorderStyle     =   1  'Fixed Single
      Height          =   4335
      Left            =   1080
      Picture         =   "frmInfoAbout.frx":0312
      Top             =   120
      Width           =   7440
   End
End
Attribute VB_Name = "frmInfoAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strAbout As String

Private Function myAbout() As String
 strAbout = "[SORRY, I JUST WRITE ABOUT THIS SOFTWARE WITH INDONESIAN LANGUAGE ONLY]" & vbNewLine
 strAbout = strAbout & "[IF YOU CAN TRANSLATE, PLEASE SEND IT TO MY EMAIL (n_n) THANX BE 4]" & vbNewLine
 strAbout = strAbout & "X-GEN yang merupakan singkatan dari XML GENERATOR merupakan sebuah SOFTWARE GRATIS "
 strAbout = strAbout & "yang dibangun untuk menyalin DATA dari File Database Microsoft Access atau "
 strAbout = strAbout & "Microsoft SQL Server ke dalam bentuk XML." & vbNewLine
 strAbout = strAbout & "Anda bisa mendapatkan SOURCE CODE dari Software ini secara GRATIS."
 
 myAbout = strAbout
End Function

Private Sub cmdOK_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 lblAbout.Caption = myAbout
End Sub

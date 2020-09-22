VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4275
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   4275
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrLoading 
      Interval        =   3000
      Left            =   6840
      Top             =   120
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub tmrLoading_Timer()
 If tmrLoading.Interval = 3000 Then
  frmMother.Show
  Unload Me
 End If
End Sub

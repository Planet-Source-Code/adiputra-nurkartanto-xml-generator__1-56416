VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMother 
   Caption         =   "X-GEN <XML - GENerator>"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMother.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbMother 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6315
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4022
            Text            =   "Created By : Adiputra Nurkartanto"
            TextSave        =   "Created By : Adiputra Nurkartanto"
            Object.ToolTipText     =   "The Programmer"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4022
            Text            =   "email : adiputra_n@yahoo.com"
            TextSave        =   "email : adiputra_n@yahoo.com"
            Object.ToolTipText     =   "email Programmer"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4022
            Text            =   "e-VOLVE .: garage team :."
            TextSave        =   "e-VOLVE .: garage team :."
            Object.ToolTipText     =   "Free Software"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4022
            Text            =   "Free Source Code"
            TextSave        =   "Free Source Code"
            Object.ToolTipText     =   "Download at www.planetsourcecode.com"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgListMother 
      Left            =   8760
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMother.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMother.frx":15A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMother.frx":227E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMother 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   1429
      ButtonWidth     =   1323
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgListMother"
      DisabledImageList=   "imgListMother"
      HotImageList    =   "imgListMother"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Key             =   "a"
            Object.ToolTipText     =   "Close Application"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New XML"
            Key             =   "b"
            Object.ToolTipText     =   "New XML Generator"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Key             =   "c"
            Object.ToolTipText     =   "Information About X-GEN"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.Image imgXML 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   5415
      Left            =   0
      Picture         =   "frmMother.frx":25A0
      Stretch         =   -1  'True
      Top             =   840
      Width           =   9495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFNew 
         Caption         =   "&New XML"
      End
      Begin VB.Menu mnuFstrip 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "&Information"
      Begin VB.Menu mnuIAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuILicense 
         Caption         =   "&License"
      End
      Begin VB.Menu mnuIProgrammer 
         Caption         =   "&Programmer"
      End
   End
End
Attribute VB_Name = "frmMother"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SetObjectPos()
 imgXML.Height = (((Me.Height - tbrMother.Height) - stbMother.Height) - 550)
 imgXML.Width = (Me.Width - 125)
End Sub

Private Sub Form_Resize()
 SetObjectPos   '[SubRutin pos : above]
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If MsgQtnYN("Close Application Will Close All X-GENERATOR that still shown !" & vbNewLine & _
              "Resume This Process ?") Then
  Terminate_App '[SubRutin pos : Module <modMain>]
  End
 Else
  Cancel = 1
 End If
End Sub

Private Sub mnuFExit_Click()
 Unload Me
End Sub

Private Sub mnuFNew_Click()
 NewXMLgen
End Sub

Private Sub mnuIAbout_Click()
 frmInfoAbout.Show 1, Me
End Sub

Private Sub mnuILicense_Click()
 frmInfoLicense.Show 1, Me
End Sub

Private Sub mnuIProgrammer_Click()
 frmInfoProgrammer.Show 1, Me
End Sub

Private Sub tbrMother_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Key
  Case "a" 'Close
   Unload Me
  Case "b" 'New XML-GEN
   NewXMLgen    '[SubRutin pos : Module <modProccess>]
  Case "c"
   frmInfoAbout.Show 1, Me
 End Select
End Sub

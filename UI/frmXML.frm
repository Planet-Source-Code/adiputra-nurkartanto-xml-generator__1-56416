VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmXML 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XML GENERATOR - [##]"
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
   Icon            =   "frmXML.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdcSave 
      Left            =   8760
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save File XML"
      FontName        =   "verdana"
   End
   Begin TabDlg.SSTab tabXML 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   8070
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "List Data View"
      TabPicture(0)   =   "frmXML.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lvwListData"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "XML View"
      TabPicture(1)   =   "frmXML.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "rtbXML"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin RichTextLib.RichTextBox rtbXML 
         Height          =   4095
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7223
         _Version        =   393217
         BackColor       =   16777215
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmXML.frx":0902
      End
      Begin MSComctlLib.ListView lvwListData 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   0
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7223
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   4
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
         Picture         =   "frmXML.frx":0990
      End
   End
   Begin MSComctlLib.ImageList imgListXML 
      Left            =   8160
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmXML.frx":84D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmXML.frx":91AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmXML.frx":9E85
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmXML.frx":AB5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmXML.frx":B439
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmXML.frx":C113
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmXML.frx":CDED
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmXML.frx":D6C7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrXML 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   1429
      ButtonWidth     =   1561
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgListXML"
      DisabledImageList=   "imgListXML"
      HotImageList    =   "imgListXML"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Key             =   "a"
            Object.ToolTipText     =   "Close Menu"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open Data"
            Key             =   "b"
            Object.ToolTipText     =   "Opne Database"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "b1"
                  Text            =   "Open Access Database [*.mdb]"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "b2"
                  Text            =   "Open SQL Server Database"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save XML"
            Key             =   "c"
            Object.ToolTipText     =   "Save To XML FIle"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Query"
            Key             =   "d"
            Object.ToolTipText     =   "Create SQL"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Convert"
            Key             =   "e"
            Object.ToolTipText     =   "Convert To XML"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refersh"
            Key             =   "f"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame freXML 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   9255
      Begin VB.Label lblQuery 
         AutoSize        =   -1  'True
         Caption         =   "No Query [Please Create SQL First]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   3420
      End
      Begin VB.Label lblData 
         Caption         =   "Database [Please Open The Database First]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   8580
      End
   End
End
Attribute VB_Name = "frmXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strMyOwnNumber As String

Dim strFileSavePath As String
Dim strFileSaveName As String

'[Flag]
Dim isOpened As Boolean
Dim isQueryOn As Boolean
Dim isWriteOn As Boolean

'[Object]
Dim objRecord As xGenDLL.Records
Dim objTable As xGenDLL.Tables
Dim objConvert As xGenDLL.Converter

'[OBJECT (CLASS)] ================================================================
'
'<Set New Object>
Private Sub Initialize_Object()
 Set objRecord = New xGenDLL.Records
 Set objTable = New xGenDLL.Tables
End Sub
'</Set New Object>

'<Dispose Object>
Private Sub Terminate_Object()
 Set objTable = Nothing
 Set objRecord = Nothing
End Sub
'</Dispose Object>
'
'[/OBJECT (CLASS)] ===============================================================

Private Sub Form_Load()
 Initialize_Object  '[SubRutin pos : above]
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If MsgQtnYN("XML GENERATOR <-" & Mid(strMyOwnNumber, 4, Len(strMyOwnNumber) - 3) & "-> Will Close !" & vbNewLine & _
             "Close this XML GENERATOR ?") Then
  Terminate_Object  '[SubRutin pos : above]
  Unload Me
 Else
 Cancel = 1
 End If
End Sub

'[QUERY] =========================================================================
'
'<Set Query>
Private Sub SetQueryDB()
 If objfrmXML.GetData_ByNo(strMyOwnNumber) Then
  frmXMLQuery.strMyOwnNumber = objfrmXML.frmNumber
  frmXMLQuery.txtSQL.Text = objfrmXML.QueryString
  frmXMLQuery.txtTable.Text = objfrmXML.myTableName
  
  frmXMLQuery.Show 1, Me
 End If
End Sub
'</Set Query>

'<Get Query>
Private Sub GetQueryDB()
 If objfrmXML.GetData_ByNo(strMyOwnNumber) Then
  Set objRecord = New xGenDLL.Records
  Set objRecord.myCon = objfrmXML.myCon
  objRecord.QueryString = objfrmXML.QueryString
  objRecord.myOpen
    
  isQueryOn = False
  lvwListData.ColumnHeaders.Clear
  lvwListData.ListItems.Clear
  If objRecord.StatOpen = OpenIsData Or objRecord.StatOpen = OpenNoData Then
   GetField                                             '[SubRutin pos : bellow]
   If objRecord.StatOpen = OpenIsData Then GetRecord    '[SubRutin pos : bellow]
   
   isQueryOn = True
   isWriteOn = False
   lblQuery.Caption = "Query Is On [Syntax OK]"
   rtbXML.Text = "[XML Will Shown Here]"
   tabXML.Tab = 0
   
   objRecord.myClose
  End If
  Set objRecord = Nothing
 End If
End Sub
'</Get Query>

'<Get All Field>
Private Sub GetField()
 Dim a As Integer
   
 lvwListData.ColumnHeaders.Add , , "No.", 500
 For a = 1 To objRecord.ListFields.Count
  lvwListData.ColumnHeaders.Add , , objRecord.ListFields(a).FieldName
 Next
End Sub
'</Get All Field>

'<Get All Record>
Private Sub GetRecord()
 Dim a As Long
 Dim ls As ListItem
 
 With objRecord
  Do While Not .myRS.EOF
   Set ls = lvwListData.ListItems.Add(, , lvwListData.ListItems.Count + 1)
   For a = 0 To .myRS.Fields.Count - 1
    ls.ListSubItems.Add , , .myRS(a).Value
   Next
   .myRS.MoveNext
  Loop
 End With
End Sub
'</Get All Record>
'
'[/QUERY] ========================================================================

Private Sub SetSaveXML()
 On Error GoTo ErrHandle

 With cdcSave
  .CancelError = True
  .Filter = "eXtensible Markup Language|*.xml|"
  .ShowSave
  strFileSavePath = .FileName
  strFileSaveName = .FileTitle
 End With
 
 If Len(Trim(strFileSavePath)) <> 0 Then
  LetSaveXML
 End If
 
ErrHandle:
 If Err.Number = 32755 Then
  Resume Next
 End If
End Sub

Private Sub LetSaveXML()
 If objfrmXML.GetData_ByNo(strMyOwnNumber) Then
  Set objConvert = New xGenDLL.Converter
  Set objConvert.myCon = objfrmXML.myCon
  
  objConvert.DataName = objfrmXML.DataName
  objConvert.myTableName = objfrmXML.myTableName
  objConvert.QueryString = objfrmXML.QueryString
  
  If objConvert.WriteXML(strFileSavePath) Then
   MsgSaveScd strFileSaveName
  End If
  
  Set objConvert = Nothing
 End If
End Sub

Private Sub tbrXML_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Key
  Case "a" 'Close
   Unload Me
  Case "b" 'Open
   '[No Action] <event in ButtonMenu>
  Case "c" 'Save
   If isWriteOn Then
    SetSaveXML        '[SubRutin pos : above]
   Else
    MsgCantResume "Convert To XML"
   End If
  Case "d" 'Query
   If isOpened Then
    SetQueryDB      '[SubRutin pos : above]
    GetQueryDB      '[SubRutin pos : above]
   Else
    MsgCantResume "Database"
   End If
  Case "e" 'Write
   If isQueryOn Then
    ConvertXML      '[SubRutin pos : bellow]
   Else
    MsgCantResume "Query SQL"
   End If
  Case "f" 'Refresh
   If isOpened Then
    GetQueryDB      '[SubRutin pos : above]
   End If
   If isQueryOn Then
    ConvertXML      '[SubRutin pos : bellow]
   End If
 End Select
End Sub

'[CONVERT TO XML] ================================================================
'
'<Convert To XML (put to the Temp File)>
Private Sub ConvertXML()
 If objfrmXML.GetData_ByNo(strMyOwnNumber) Then
  Set objConvert = New xGenDLL.Converter
  Set objConvert.myCon = objfrmXML.myCon
  
  objConvert.DataName = objfrmXML.DataName
  objConvert.myTableName = objfrmXML.myTableName
  objConvert.QueryString = objfrmXML.QueryString
  
  isWriteOn = False
  If objConvert.WriteXML(objfrmXML.PathTempFile) Then
   LoadTemp_ViewXML '[SubRutin pos : bellow]
   tabXML.Tab = 1
   isWriteOn = True
  End If
  
  Set objConvert = Nothing
 End If
End Sub
'</Convert To XML (put to the Temp File)>

'<Load XML From Temp File & Show it to the screen>
Private Sub LoadTemp_ViewXML()
 Dim dblLastPos As Double
 Dim dblLength As Double
 Dim dblNewStart As Double
  
 rtbXML.Text = ""
 rtbXML.SelColor = vbBlack
 rtbXML.SelBold = True
 
 rtbXML.LoadFile (objfrmXML.PathTempFile), 1
 
 rtbXML.Span ("<?xml version=" & Chr(34) & "1.0" & Chr(34) & "?>")
 rtbXML.SelColor = vbBlue
 rtbXML.SelBold = False
 dblLastPos = rtbXML.SelLength
 
 Do While dblLastPos > -1
  dblLastPos = rtbXML.Find("<!--", dblLastPos + 1)
  rtbXML.SelColor = vbBlue
  rtbXML.SelBold = False
 Loop
 
 dblLastPos = 0
 
 Do While dblLastPos > -1
  dblLastPos = rtbXML.Find("<!--", dblLastPos + 1)
  If dblLastPos = -1 Then Exit Do
  
  dblLength = (rtbXML.Find("-->", (dblLastPos + 4)) - (dblLastPos + 4))
  rtbXML.SelStart = dblLastPos + 4
  rtbXML.SelLength = dblLength
  rtbXML.SelColor = &H808080
  rtbXML.SelBold = False
 Loop
   
 dblLastPos = 0
 
 Do While dblLastPos > -1
  dblNewStart = dblLastPos
  dblLastPos = rtbXML.Find("-->", dblLastPos + 1)
  rtbXML.SelColor = vbBlue
  rtbXML.SelBold = False
 Loop
 
 dblLastPos = dblNewStart
 
 Do While dblLastPos > -1
  dblLastPos = rtbXML.Find("<", dblLastPos + 1)
  If dblLastPos = -1 Then Exit Do
  rtbXML.SelColor = vbBlue
  rtbXML.SelBold = False
  
  dblLength = (rtbXML.Find(">", dblLastPos) - dblLastPos)
  rtbXML.SelStart = dblLastPos + 1
  rtbXML.SelLength = dblLength
  rtbXML.SelColor = &H80&
  rtbXML.SelBold = False
 Loop
 
 dblLastPos = 0
 
 Do While dblLastPos > -1
  dblLastPos = rtbXML.Find(">", dblLastPos + 1)
  rtbXML.SelColor = vbBlue
  rtbXML.SelBold = False
 Loop
 
 dblLastPos = 0
 
 Do While dblLastPos > -1
  dblLastPos = rtbXML.Find("</", dblLastPos + 1)
  rtbXML.SelColor = vbBlue
  rtbXML.SelBold = False
 Loop
 
 dblLastPos = dblNewStart
 
 Do While dblLastPos > -1
  dblLastPos = rtbXML.Find(">", dblLastPos + 1)
  dblLength = (rtbXML.Find("<", dblLastPos) - dblLastPos - 1)
  If dblLength < 0 Then Exit Do
  rtbXML.SelStart = dblLastPos + 1
  rtbXML.SelLength = dblLength
  rtbXML.SelColor = vbBlack
  rtbXML.SelBold = True
 Loop
 
 'rtbXML.SelColor = vbBlack
 rtbXML.SelStart = 0
 rtbXML.SetFocus
 
 Kill (objfrmXML.PathTempFile)
End Sub
'</Load XML From Temp File & Show it to the screen>
'
'[/CONVERT TO XML] ===============================================================

'[OPEN FILE ACCESS (*.MDB)] ======================================================
'
'<Set Access File>
Private Sub SetOpenAcces()
 If objfrmXML.GetData_ByNo(strMyOwnNumber) Then
  frmXMLOpenMDB.strMyOwnNumber = objfrmXML.frmNumber
  
  frmXMLOpenMDB.Show 1, Me
 End If
End Sub
'</Set Access File>

'<Get Access File>
Private Sub GetOpenAcces()
 If objfrmXML.GetData_ByNo(strMyOwnNumber) Then
  If objfrmXML.isOpened Then
   lblData.Caption = "Data [+path] : " & objfrmXML.DataPath
   lblQuery.Caption = "No Query [Please Create SQL First]"
   isQueryOn = False
   isWriteOn = False
  End If
  isOpened = objfrmXML.isOpened
 End If
End Sub
'</Get Access File>
'
'[/OPEN FILE ACCESS (*.MDB)] =====================================================

Private Sub tbrXML_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
 Select Case ButtonMenu.Key
  Case "b1" 'Open Access Database
   SetOpenAcces     '[SubRutin pos : above]
   GetOpenAcces     '[SubRutin pos : above]
  Case "b2"
   SetOpenSQLServer '[SubRutin pos : bellow]
   GetOpenSQLServer '[SubRutin pos : bellow]
 End Select
End Sub

'[OPEN SQL SERVER DATABASE] ======================================================
'
'<Set SQL Database>
Private Sub SetOpenSQLServer()
 If objfrmXML.GetData_ByNo(strMyOwnNumber) Then
  frmXMLOpen.strMyOwnNumber = objfrmXML.frmNumber
    
  frmXMLOpen.Show 1, Me
 End If
End Sub
'</Set SQL Database>

'<Get SQL Database>
Private Sub GetOpenSQLServer()
 If objfrmXML.GetData_ByNo(strMyOwnNumber) Then
  If objfrmXML.isOpened Then
   lblData.Caption = "Server Name : " & objfrmXML.ServerName & "        " & _
                     "Database Name : " & objfrmXML.DataName
   lblQuery.Caption = "No Query [Please Create SQL First]"
   isQueryOn = False
   isWriteOn = False
  End If
  isOpened = objfrmXML.isOpened
 End If
End Sub
'<Get SQL Database>
'
'[/OPEN SQL SERVER DATABASE] =====================================================

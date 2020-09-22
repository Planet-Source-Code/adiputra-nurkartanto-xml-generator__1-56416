VERSION 5.00
Begin VB.Form frmInfoLicense 
   BackColor       =   &H00CCBAB7&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "License"
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
   Icon            =   "frmInfoLicense.frx":0000
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
      TabIndex        =   0
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label lblLicense 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "[License Will Shown Here]"
      ForeColor       =   &H80000008&
      Height          =   5800
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   9255
   End
End
Attribute VB_Name = "frmInfoLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strLicense As String

Private Function myLicense() As String
 strLicense = "[SORRY, I JUST WRITE LICENSE OF THIS SOFTWARE WITH INDONESIAN LANGUAGE ONLY]" & vbNewLine
 strLicense = strLicense & "[IF YOU CAN TRANSLATE, PLEASE SEND IT TO MY EMAIL (n_n) THANX BE 4]" & vbNewLine & vbNewLine
 strLicense = strLicense & "Software Aplikasi ini dibuat oleh Adiputra Nurkartanto. "
 strLicense = strLicense & "MOHON UNTUK TIDAK MERUBAH NAMA PEMBUAT DAN INFORMASI LAINNYA YANG TERKAIT DENGAN APLIKASI INI." & vbNewLine
 strLicense = strLicense & "Anda dapat menggunakan baik Aplikasi ini ataupun Source Code yang Anda dapatkan untuk keperluan Anda. " & vbNewLine
 strLicense = strLicense & "Untuk keperluan/kepentingan yang mendatangkan hasil (komersil), diharap untuk menghubungi pembuat "
 strLicense = strLicense & "terlebih dahulu. Hal tersebut dilakukan untuk melihat perkembangan Aplikasi ini di mata para pengguna." & vbNewLine & vbNewLine
 strLicense = strLicense & "Point-point penting yang harus diperhatikan dalam mengunakan Software X-GEN dan Source Code-nya !" & vbNewLine
 strLicense = strLicense & " -- Pembuat Aplikasi tidak bertanggung jawab atas ketidak stabilan/ketidak konsistenan data setelah/sebelum Aplikasi ini ter-Install." & vbNewLine
 strLicense = strLicense & " -- Pembuat Aplikasi tidak bertanggung jawab atas ketidak stabilan/ketidak konsistenan data setelah/sebelum Aplikasi ini dijalankan." & vbNewLine
 strLicense = strLicense & " -- Pembuat Aplikasi tidak bertanggung jawab atas semua kerusakan yang terjadi akibat menggunakan Aplikasi ini." & vbNewLine
 strLicense = strLicense & " -- Anda menggunakan/meng-install Aplikasi ini, berarti Anda setuju dengan semua peraturan yang ditetapkan." & vbNewLine & vbNewLine
 strLicense = strLicense & "Atas kerjasama dan pengertiannya, pembuat mengucapkan terima kasih. Selamat Menggunakan !" & vbNewLine & vbNewLine & vbNewLine
 strLicense = strLicense & "[build form scratch]" & vbNewLine
 strLicense = strLicense & "eVOLVE .:: garage team ::." & vbNewLine & vbNewLine
 strLicense = strLicense & ".:: ADIPUTRA NURKARTANTO ::."
 
 myLicense = strLicense
End Function

Private Sub cmdOK_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 lblLicense.Caption = myLicense
End Sub

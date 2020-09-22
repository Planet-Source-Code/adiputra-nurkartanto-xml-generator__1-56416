Attribute VB_Name = "modMain"
Option Explicit

Public objfrmXML As xGenDLL.frmXMLclass

'StartUp App
Sub Main()
 Initialize_App     '[SubRutin pos : bellow]
 
 frmSplash.Show
End Sub

'[OBJECT HANDLE frmXML] ==========================================================
'
'<Call This Method When Application Is Started>
Public Sub Initialize_App() 'Let's Make Contrusctor (n_n)
 Set objfrmXML = New xGenDLL.frmXMLclass
 intCount = 0                              '[use for frmXML <- $increment ->]
End Sub
'</Call This Method When Application Is Started>

'<Call This Method When Application Is Stoped/Closed/Destroyed>
Public Sub Terminate_App() 'Let's Make Destructor (n_n)
 Set objfrmXML = Nothing
End Sub
'</Call This Method When Application Is Stoped/Closed/Destroyed>

'[/OBJECT HANDLE frmXML] =========================================================

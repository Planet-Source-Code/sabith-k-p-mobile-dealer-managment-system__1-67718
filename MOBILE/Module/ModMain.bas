Attribute VB_Name = "ModMain"
'=======================
'+Author:   Sabith Kp  +
'=======================

Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public CurrentUser As tUser
Public Active_Form  As String
Public RptMobileNO As String


Public Sub Main()

CrackKey = GetSetting("SKG", "SKGKey", "SCheck")
   CrackSKey = GetSetting(App.Title, "RegKey", "Key")
   If CrackKey = "" Then
   '''''''''''''''If CrackKey = "" Then
 frmRegistration.Show 'vbModalc
 ElseIf CrackSKey = "" Then
 frmRegistration.Show
 Else
' End If
    'use system appearance style
    InitCommonControls
    
    'show author message
    'frmMSG.ShowForm
    

    'check license
    'If modLicense.CheckLicense = False Then
    '    End
    'End If


    'init global variables
'    Call modGv.InitGV
    
    'set Database Path
    If InitDB = False Then
        Exit Sub
    End If
    
    
    'Show Splash
    frmSplash.ShowSplash
       
End If
End Sub
Public Sub Main_AfterSD()
    

    'Open Database File
    If OpenDB = False Then
        Exit Sub
    End If
     
    
    'TestUnit
  MDIFrm.ShowForm
End Sub

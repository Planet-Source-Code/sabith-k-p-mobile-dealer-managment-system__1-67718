VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRegistration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Registration Wizard"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   6585
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pic2 
      BackColor       =   &H00E0E0E0&
      Height          =   1815
      Left            =   -120
      ScaleHeight     =   1755
      ScaleWidth      =   5115
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdRegNext 
         Caption         =   "Next>>"
         Height          =   375
         Left            =   4080
         TabIndex        =   12
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "<<Back"
         Height          =   375
         Left            =   3120
         TabIndex        =   11
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtSerialKey 
         Height          =   375
         Left            =   960
         TabIndex        =   9
         Top             =   480
         Width           =   3855
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   5040
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter the Serial Key"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.PictureBox Pic3 
      BackColor       =   &H00E0E0E0&
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1755
      ScaleWidth      =   4995
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CommandButton cmdFinish 
         Caption         =   "Finish"
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This product is Licenced to:Tokyo Mahotsavam"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   3960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serial key:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   885
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   5040
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registratin Wizard Completed Succesfuly!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   3525
      End
   End
   Begin VB.PictureBox Pic1 
      BackColor       =   &H00E0E0E0&
      Height          =   2055
      Left            =   -120
      ScaleHeight     =   1995
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5760
         TabIndex        =   16
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "Next"
         Height          =   375
         Left            =   4800
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate"
         Height          =   375
         Left            =   5760
         TabIndex        =   14
         Top             =   600
         Width           =   855
      End
      Begin MSComctlLib.ProgressBar Pbar 
         Height          =   255
         Left            =   3480
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.TextBox txtGeneratelKey 
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   5415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome Tokyo Registration Wizard"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   2955
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   480
         X2              =   7320
         Y1              =   1200
         Y2              =   1200
      End
   End
End
Attribute VB_Name = "frmRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public SerialKey As String
Public RegVal As String
Public EncryptValueofHard  As String
Public CrackKey As String
Public SDbKey As String

Private Sub cmdGenrate_Click()

End Sub

Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdBack_Click()
Pic1.Visible = True
pic2.Visible = False
Pic3.Visible = False
Me.Height = 2475
Me.Width = 6720

End Sub

Private Sub cmdCance_Click()
End
End Sub

Private Sub cmdFinish_Click()
Unload Me
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
       

End Sub

Private Sub cmdGenerate_Click()
On Error GoTo Derror
            Pbar.Visible = True
            Pbar.Value = 55
                HardInfo = ""
                    GenerateHardDiskInfo
            Pbar.Value = 70
                    EncryptValueofHard = Encrypt(HardInfo)
                    SerialKey = Mid(EncryptValueofHard, 5, 6) & Mid(EncryptValueofHard, 7, 8) & Mid(EncryptValueofHard, 1, 5)
              Pbar.Value = 80
                    RegVal = "WINXP" & UCase(Mid(SerialKey, 2, 13))
                        SaveSetting "SKG", "SKGKey", "SCheck", RegVal
                    txtGeneratelKey.Text = EncryptValueofHard
            Pbar.Value = 100
            Pbar.Visible = False
                txtGeneratelKey.Locked = True
                cmdGenerate.Enabled = False
            Randomize
           Exit Sub
Derror:
        MsgBox "Please Check the Date and Time are Correct", vbInformation
       MsgBox "Unhandled error has Occured" & vbCrLf & "Please Contact Developer for More information", vbCritical
       
End Sub

Private Sub cmdNext_Click()
''''' key Genertion


''''''' check empty
If txtGeneratelKey = "" Then
    MsgBox "Please Generate the key " & vbCrLf & "   and Try again!", vbExclamation
    txtGeneratelKey.SetFocus
Exit Sub
End If
'form alignment
Pic1.Visible = False
pic2.Visible = True
Me.Height = 2115
Me.Width = 5130



End Sub

Private Sub cmdRegBack_Click()

End Sub

Private Sub dcButton1_Click()

End Sub

Private Sub dcButton2_Click()

End Sub

Private Sub cmdRegNext_Click()
Dim RegCrack As String
Dim RegCrackKey As String
''''check for key
On Error GoTo RegErr
            RegCrackKey = GetSetting("SKG", "SKGKey", "SCheck")
                CrackKey = Mid(RegCrackKey, 8, 7) & Mid(RegCrackKey, 10, 9) '& Mid(RegCrackKey, 9, 10) & Mid(RegCrackKey, 7, 8)
                If txtSerialKey.Text = CrackKey Then
                    SaveSetting App.Title, "RegKey", "Key", txtSerialKey.Text
                     Pic3.Visible = True
                     Label4.Caption = "Serial key: " & txtSerialKey.Text
                        Pic1.Visible = False
                        pic2.Visible = False
                       
                        
                Else
                MsgBox "Invalid Registration Key,Try Again!", vbExclamation, "Invalid Key"
           End If
           Exit Sub
RegErr:
    'if you erase key from Database
    MsgBox "Unhandled error has Occured" & vbCrLf & "Please Contact the developer for more information!", vbCritical, "Unhandled error"
    End
'''goto finish
    

End Sub

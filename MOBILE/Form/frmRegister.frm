VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration Wizard"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   1695
      Left            =   -120
      ScaleHeight     =   1635
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton dcButton1 
         Caption         =   "Finish"
         Height          =   375
         Left            =   4200
         TabIndex        =   5
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton cmdRegNext 
         Caption         =   "Next>>"
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "<<Back"
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtSerialKey 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   3855
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   5160
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter the Serial Key"
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
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public SerialKey As String
Public RegVal As String
Public EncryptValueofHard  As String
Public CrackKey As String
Public SDbKey As String

Private Sub cmdBack_Click()
frmRegistration.Show
End Sub

Private Sub cmdNext_Click()
On Error GoTo RegErr
       ' If rs.State = adStateOpen Then
         '   rs.Close
       ' End If
               
            CrackKey = GetSetting("SKG", "SKGKey", "SCheck")
           ' rs.Open "select * from Skey", cn, adOpenKeyset, adLockPessimistic
           ' SDbKey = rs.Fields("Key")
        If CrackKey = "win" & LCase(Mid(SDbKey, 2, 9)) Then
                'PbBar.Value = PbBar.Value + 100
            MsgBox "Registration Completed" & vbCrLf & "      Succesfuly", vbInformation, IDentity
                'txtSerialkey.Locked = True
                'frmDemo.Show
        Else
            MsgBox " Invalid Serial Key " & vbCrLf & vbCrLf & "  Try Again!", vbExclamation, IDentity
            txtSerialKey.SetFocus
        End If
            rs.Close
        Exit Sub
RegErr:
    'if you erase key from Database
    MsgBox "Unhandled error has Occured" & vbCrLf & "The Current Record has been Deleted" & vbCrLf & "Generate New Key and Try Again!", vbCritical, "Unhandled error"
    End
End Sub

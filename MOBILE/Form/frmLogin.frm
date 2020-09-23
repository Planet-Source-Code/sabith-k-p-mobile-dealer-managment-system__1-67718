VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4830
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EDEBE9&
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3315
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   -120
      Width           =   4935
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Login"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "="
         TabIndex        =   4
         Top             =   2040
         Width           =   3135
      End
      Begin VB.TextBox txtUserID 
         Height          =   325
         Left            =   1560
         TabIndex        =   3
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Image Image2 
         Height          =   1335
         Left            =   0
         Picture         =   "frmLogin.frx":0000
         Top             =   120
         Width           =   4830
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   5040
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Left            =   360
         TabIndex        =   2
         Top             =   2160
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UserName"
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
         Left            =   360
         TabIndex        =   1
         Top             =   1680
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Dim mShowForm As Boolean
Dim dFailedCount As Integer

Public Function ShowForm() As Boolean

    'show form
   Me.Show vbModal
    
    ShowForm = mShowForm
End Function

Private Sub cmdCancel_Click()
    mShowForm = False
    Unload Me
End Sub

Private Sub cmdLogin_Click()
    
    'check form field
    If IsEmpty(txtUserID.Text) Then
        MsgBox "Please enter User ID", vbExclamation
        HLTxt txtUserID
        Exit Sub
    End If
    
    If IsEmpty(txtPassword.Text) Then
        MsgBox "Please enter Password", vbExclamation
        HLTxt txtPassword
        Exit Sub
    End If
    
    'check user
    If GetUserByID(txtUserID.Text, CurrentUser) = False Then
        MsgBox "User does not exist.", vbExclamation
        HLTxt txtUserID
        Exit Sub
    End If
    
    'check password
    If txtPassword.Text <> CurrentUser.Password Then
    
        If dFailedCount >= 5 Then
            MsgBox "Log Error " & " Err: 0000x000FF", vbCritical
            
            Unload Me
            Exit Sub
        End If
    
        MsgBox "Invalid Password.", vbExclamation
        HLTxt txtPassword
        
        'increment counter
        dFailedCount = dFailedCount + 1
        
        Exit Sub
    End If
    
    
    'set current user
    If GetUserByID(Trim(txtUserID.Text), CurrentUser) = False Then
       MsgBox Me.Name & "cmdLogin_Click" & GetUserByID(Trim(txtUserID.Text), CurrentUser) = False
        Unload Me
    End If
    
    
    'success
    'write to log
    'temp
    
    'set flag
    mShowForm = True
    'close this form
    Unload Me
    
End Sub


Private Sub Form_Load()
    'default
   
    dFailedCount = 0
    
    txtUserID.Text = GetSetting(App.EXEName, "TextBox", txtUserID.Name, "")
    'PaintGrad Me, &HEDEBE9, &HF5F5F5, 135
End Sub


Private Sub Form_Unload(Cancel As Integer)
 
    SaveSetting App.EXEName, "TextBox", txtUserID.Name, txtUserID.Text
End Sub


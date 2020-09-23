VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3330
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   3330
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClear 
      DisabledPicture =   "frmSearch.frx":0000
      Height          =   345
      Left            =   2920
      Picture         =   "frmSearch.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   345
   End
   Begin VB.CommandButton cmdSearch 
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2550
      Picture         =   "frmSearch.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
txtSearch.Text = ""
End Sub
Public Function ShowForm() As Boolean

    'show form
    Me.Show vbModal
    
    ShowForm = mShowForm
End Function
Private Sub cmdSearch_Click()
If txtSearch.Text = "" Then
MsgBox "Please Enter the ID", vbExclamation
txtSearch.SetFocus
Exit Sub
End If
Select Case Active_Form
Case "Coupon"

     frm_aed_Coupon.ShowEdit txtSearch.Text

Case "Sim"
 
    frm_aed_Sim.ShowEdit txtSearch.Text
 
Case "ECharge"

    'frm_aed_Echarge.ShowEdit txtSearch.Text

Case "Activation"

   frm_aed_Activation.ShowEdit txtSearch.Text
End Select
End Sub

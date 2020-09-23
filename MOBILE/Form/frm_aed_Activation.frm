VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_aed_Activation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SIM  Activation Wizard"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   6330
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5953
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Main"
      TabPicture(0)   =   "frm_aed_Activation.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DTDate"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DTPActivation"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtID"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdNext"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtmobileNo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Timer1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "&Records"
      TabPicture(1)   =   "frm_aed_Activation.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line2"
      Tab(1).Control(1)=   "optDN"
      Tab(1).Control(2)=   "Frame5"
      Tab(1).Control(3)=   "Frame4"
      Tab(1).Control(4)=   "Frame3"
      Tab(1).Control(5)=   "Frame2"
      Tab(1).Control(6)=   "Frame1"
      Tab(1).Control(7)=   "cmdRback"
      Tab(1).Control(8)=   "cmdRNext"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "&Details"
      TabPicture(2)   =   "frm_aed_Activation.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label11"
      Tab(2).Control(1)=   "Line3"
      Tab(2).Control(2)=   "Label6"
      Tab(2).Control(3)=   "Label5"
      Tab(2).Control(4)=   "Label4"
      Tab(2).Control(5)=   "Label3"
      Tab(2).Control(6)=   "DTPMeff"
      Tab(2).Control(7)=   "DTDelivery"
      Tab(2).Control(8)=   "txtAddress"
      Tab(2).Control(9)=   "cmdDback"
      Tab(2).Control(10)=   "cmdSave"
      Tab(2).Control(11)=   "txtCoustomerName"
      Tab(2).Control(12)=   "cmbOutlet"
      Tab(2).Control(13)=   "chkComplete"
      Tab(2).ControlCount=   14
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3960
         Top             =   480
      End
      Begin VB.CheckBox chkComplete 
         Caption         =   "Complete Wizard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70680
         TabIndex        =   41
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtmobileNo 
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next>>"
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
         Left            =   5280
         TabIndex        =   2
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
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
         Left            =   4200
         TabIndex        =   3
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton cmdRNext 
         Caption         =   "Next>>"
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
         Left            =   -69720
         TabIndex        =   10
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton cmdRback 
         Caption         =   "<< Back"
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
         Left            =   -70800
         TabIndex        =   11
         Top             =   2880
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Caption         =   "Attest in Photograph"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74640
         TabIndex        =   31
         Top             =   840
         Width           =   2415
         Begin VB.OptionButton optApY 
            Caption         =   "&Yes"
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
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optApN 
            Caption         =   "&No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   15
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Photograph"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74640
         TabIndex        =   30
         Top             =   1440
         Width           =   2415
         Begin VB.OptionButton optPY 
            Caption         =   "&Yes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optPN 
            Caption         =   "&No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   0
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Photo ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74640
         TabIndex        =   28
         Top             =   2040
         Width           =   2415
         Begin VB.OptionButton optPIY 
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optPIN 
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   29
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Coustomer Signature in APEF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -72120
         TabIndex        =   26
         Top             =   840
         Width           =   2535
         Begin VB.OptionButton optAEFY 
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optAEFN 
            Caption         =   "No"
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
            Left            =   1200
            TabIndex        =   27
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Retail Seal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -72120
         TabIndex        =   24
         Top             =   1440
         Width           =   2535
         Begin VB.OptionButton optRY 
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optRN 
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   25
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame optDN 
         Caption         =   "Disitributer Seal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -72120
         TabIndex        =   22
         Top             =   2040
         Width           =   2535
         Begin VB.OptionButton optDSY 
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optDSN 
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   23
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox txtID 
         Height          =   325
         Left            =   1800
         TabIndex        =   21
         Top             =   1080
         Width           =   2055
      End
      Begin VB.ComboBox cmbOutlet 
         Height          =   315
         Left            =   -73080
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtCoustomerName 
         Height          =   375
         Left            =   -73080
         TabIndex        =   13
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Finish"
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
         Left            =   -69720
         TabIndex        =   19
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton cmdDback 
         Caption         =   "<<Back"
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
         Left            =   -70800
         TabIndex        =   20
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtAddress 
         Height          =   375
         Left            =   -73080
         TabIndex        =   14
         Top             =   1440
         Width           =   4335
      End
      Begin MSComCtl2.DTPicker DTDelivery 
         Height          =   330
         Left            =   -73080
         TabIndex        =   18
         Top             =   2280
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         Format          =   50855937
         CurrentDate     =   39091
      End
      Begin MSComCtl2.DTPicker DTPMeff 
         Height          =   330
         Left            =   -73080
         TabIndex        =   17
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         Format          =   50855937
         CurrentDate     =   39091
      End
      Begin MSComCtl2.DTPicker DTPActivation 
         Height          =   330
         Left            =   1800
         TabIndex        =   32
         Top             =   1560
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         Format          =   50855937
         CurrentDate     =   39091
      End
      Begin MSComCtl2.DTPicker DTDate 
         Height          =   330
         Left            =   4440
         TabIndex        =   42
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         Format          =   50855937
         CurrentDate     =   39091
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MobileNo"
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
         Left            =   960
         TabIndex        =   40
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   960
         TabIndex        =   39
         Top             =   1560
         Width           =   405
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6360
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line2 
         X1              =   -74760
         X2              =   -68520
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
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
         Left            =   960
         TabIndex        =   38
         Top             =   1080
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Outlet Name"
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
         Left            =   -74640
         TabIndex        =   37
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coustomer name"
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
         Left            =   -74640
         TabIndex        =   36
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Meff Date"
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
         Left            =   -74640
         TabIndex        =   35
         Top             =   1920
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date"
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
         Left            =   -74640
         TabIndex        =   34
         Top             =   2280
         Width           =   1155
      End
      Begin VB.Line Line3 
         X1              =   -74520
         X2              =   -68520
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   -74640
         TabIndex        =   33
         Top             =   1560
         Width           =   690
      End
   End
End
Attribute VB_Name = "frm_aed_Activation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================
'+Author:   Sabith Kp  +
'=======================
Dim SActivation As aActivation
Dim mShowAdd As Boolean
Dim mShowEdit As Boolean
Dim mFormState As String

Dim AIP As String
Dim PGP As String
Dim PID As String
Dim CSAPEF As String
Dim ARS As String
Dim ADS As String

'
Dim EAIP As String
Dim EPGP As String
Dim EPID As String
Dim ECSAPEF As String
Dim ERS As String
Dim EDS As String

Public Function ShowForm() As Boolean

    'show form
    Me.Show vbModal
    
    ShowForm = mShowForm
End Function
    
Public Function ShowEdit(MobileNo As String) As Boolean
    
    'set form state
    mFormState = "edit"
    
    'set parameter
    SActivation.MobileNo = MobileNo
    
    'show form
    Me.Show vbModal
    
    'return
    ShowEdit = mShowEdit
    
End Function
Public Function ShowAdd() As Boolean
    
      'set form state
    mFormState = "add"
    
    'show form
    Me.Show vbModal
    
    'return
    ShowAdd = mShowAdd
    
End Function


Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
    Select Case mFormState
        Case "add"
            SaveAdd
        
        Case "edit"
            SaveEdit
    End Select
    End Sub

Private Sub cmdDBack_Click()
SSTab1.Tab = 1
End Sub

Private Sub cmdNext_Click()
    Dim oldActivation As aActivation
If txtmobileNo.Text = "" Then
MsgBox "Please Enter the Mobile no", vbExclamation
txtmobileNo.SetFocus
Exit Sub
End If
If DTPActivation.Value <> Date Then
If MsgBox(" Select Date is Incorrect. Do you want to Continue..? ", vbYesNo + vbExclamation, "Warning") = vbYes Then
Else
Exit Sub
End If
End If
If mFormState = "add" Then
If GetActivationNo(txtmobileNo.Text, oldActivation) = True Then
        MsgBox "The Activation MobileNo that you have entered is already exist." & vbNewLine & vbNewLine & _
            "Please enter different value.", vbExclamation
        HLTxt txtmobileNo
        Exit Sub
    End If
End If
SSTab1.Tab = 1
End Sub

Private Sub cmdRBack_Click()
SSTab1.Tab = 0
End Sub

Private Sub cmdRNext_Click()
If optApY.Value = False And optApN.Value = False Then
MsgBox "Mark the detils about Attest in photograph ", vbExclamation
Exit Sub
End If

If optPY.Value = False And optPN.Value = False Then
MsgBox "Mark the detils about photograph ", vbExclamation
Exit Sub
End If

If optPIY.Value = False And optPIN.Value = False Then
MsgBox "Mark the detils about Photo ID ", vbExclamation
Exit Sub
End If

If optAEFY.Value = False And optAEFN.Value = False Then
MsgBox "Mark the detils about Coustmer Signature in APEF ", vbExclamation
Exit Sub
End If

If optRY.Value = False And optRN.Value = False Then
MsgBox "Mark the detils about Retail Seal", vbExclamation
Exit Sub
End If



If optDSY.Value = False And optDSN.Value = False Then
MsgBox "Mark the detils about Distributer Seal", vbExclamation
Exit Sub
End If

    
SSTab1.Tab = 2

End Sub


Private Sub Command2_Click()
Unload Me
End Sub



Private Sub Form_Activate()
Select Case mFormState
        Case "add"
        
            'set caption
            'Me.Caption = "Add Activation"
            'Me.cmdSave.Caption = "&Save"
                txtID = modFunction.ComNumZ(GetNewID, 2)

        Case "edit"
        txtmobileNo.Locked = True
            'get info
            If GetActivationNo(SActivation.MobileNo, SActivation) = False Then
                'show failed
                MsgBox "User entry with Phone : '" & SActivation.ID & "' does not exist.", vbExclamation
                'close this form
                Unload Me
                Exit Sub
            End If
            
            
            'set form ui info
            txtID.Text = SActivation.ID
            DTDate.Value = SActivation.CurDate
            txtmobileNo.Text = SActivation.MobileNo
            DTPActivation.Value = SActivation.aDate
             EAIP = SActivation.Attest
            EPGP = SActivation.Photograph
            EPID = SActivation.PhotoId
            ECSAPEF = SActivation.APEF
            ERS = SActivation.Retailseal
            EDS = SActivation.Distributer
            Call Read_record
            cmbOutlet.Text = SActivation.OutletName
            txtCoustomerName.Text = SActivation.CoustomerName
            txtAddress.Text = SActivation.Address
            DTPMeff.Value = SActivation.MeffDate
            DTDelivery.Value = SActivation.DeliveryDate
             a# = SActivation.CompeletWizard
             
            If a# = 1 Then
            chkComplete.Value = 1
            Else
            chkComplete.Value = 0
            End If
           
            'set caption
            Me.Caption = "Edit Activation"
            'Me.cmdSave.Caption = "&Update"
            
            txtID.Locked = True
            
    
    End Select
    txtmobileNo.SetFocus

End Sub


    


Private Function SaveAdd()
    Dim NewActivation As aActivation
    Dim oldActivation As aActivation
    
    
    'check form field
    If IsEmpty(txtID.Text) Then
        MsgBox "Please Enter ActivationId", vbExclamation
        HLTxt txtID
        Exit Function
    End If
   
    If mFormState = "add" Then
       
    'check duplication
    If GetActivationNo(txtmobileNo.Text, oldActivation) = True Then
        MsgBox "The Activation MobileNo that you have entered is already exist." & vbNewLine & vbNewLine & _
            "Please enter different value.", vbExclamation
        HLTxt txtmobileNo
        Exit Function
    End If
    NewActivation.ID = txtID.Text
    NewActivation.CurDate = DTDate.Value
    NewActivation.MobileNo = txtmobileNo.Text
    NewActivation.aDate = DTPActivation.Value
    Call RecordS_Status
    NewActivation.Attest = AIP
    NewActivation.Photograph = PGP
    NewActivation.PhotoId = PID
    NewActivation.APEF = CSAPEF
    NewActivation.Retailseal = ARS
    NewActivation.Distributer = ADS
    NewActivation.OutletName = cmbOutlet.Text
    NewActivation.CoustomerName = txtCoustomerName.Text
    NewActivation.Address = txtAddress.Text
    NewActivation.MeffDate = DTPMeff.Value
    NewActivation.DeliveryDate = DTDelivery.Value
    a# = chkComplete.Value
    If a# = 1 Then
    NewActivation.CompeletWizard = 1
    Else
    NewActivation.CompeletWizard = 0
    End If
    'try
    
    If ModRsActivation.AddActivation(NewActivation) = True Then
        MsgBox "New Activation entry was successfully created.", vbInformation
        'set flag
        mShowAdd = True
        'close form and return
        Unload Me
        frm_aed_Activation.ShowAdd
        
        
    Else
    
        MsgBox "Unable to add new Activation entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    End If
End Function

Private Function SaveEdit()

    Dim NewActivation As aActivation
    Dim oldActivation As aActivation
    
    'check form field
    
    
      If IsEmpty(txtID.Text) Then
        MsgBox "Please enter Activation Id", vbExclamation
        HLTxt txtID
        Exit Function
    End If
    
    
    'set new
  NewActivation.ID = txtID.Text
  NewActivation.CurDate = DTDate.Value
    NewActivation.MobileNo = txtmobileNo.Text
    NewActivation.aDate = DTPActivation.Value
    Call RecordS_Status
    NewActivation.Attest = AIP
    NewActivation.Photograph = PGP
    NewActivation.PhotoId = PID
    NewActivation.APEF = CSAPEF
    NewActivation.Retailseal = ARS
    NewActivation.Distributer = ADS
    NewActivation.OutletName = cmbOutlet.Text
    NewActivation.CoustomerName = txtCoustomerName.Text
    NewActivation.Address = txtAddress.Text
    NewActivation.MeffDate = DTPMeff.Value
    NewActivation.DeliveryDate = DTDelivery.Value
    a# = chkComplete.Value
    If a# = 1 Then
    NewActivation.CompeletWizard = 1
    Else
    NewActivation.CompeletWizard = 0
    End If
       
    
       
    'try
    'add new Activation
    If ModRsActivation.EditActivation(NewActivation) = True Then
        MsgBox "Activation entry was updated.", vbInformation
        'set flag
        mShowEdit = True
        'close form and return
        Unload Me
       frmSearch.txtSearch.Text = ""
    Else
    
        MsgBox "Unable to update Activation entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    
End Function


Private Sub Form_Load()
ComboBuilder "tbloutlet", "Name", "ID", cmbOutlet
DTPActivation.Value = Date
End Sub
Public Function GetNewID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetNewID = -1
    
    sSQL = "SELECT Max(tblActivation.ID)+1 AS ID" & _
            " From tblActivation"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        MsgBox "GetNewID" & "," & "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetNewID = 1
        GoTo RAE
    End If
    
    On Error Resume Next
    GetNewID = ReadField(vRS.Fields("ID"))
    
    If GetNewID < 1 Then
        GetNewID = 1
        txtID.Text = ID
    End If
    
RAE:
    Set vRS = Nothing
    Err.Clear
End Function

Public Sub RecordS_Status()


'check  Attest in photograph
    If optApY.Value = True Then
        AIP = "Yes"
    ElseIf optApN.Value = True Then
        AIP = "No"
    End If
    
'check Photograph

    If optPY.Value = True Then
        PGP = "Yes"
    ElseIf optPN.Value = True Then
        PGP = "No"
    End If
    
'check Photo ID

    If optPIY.Value = True Then
        PID = "Yes"
    ElseIf optPIN.Value = True Then
        PID = "No"
    End If
    
'check Couster signature in APEF

    If optAEFY.Value = True Then
        CSAPEF = "Yes"
    ElseIf optAEFN.Value = True Then
        CSAPEF = "No"
    End If
    
'Retail seal

    If optRY.Value = True Then
        ARS = "Yes"
    ElseIf optRN.Value = True Then
        ARS = "No"
    End If
    
'Distributer Seal

    If optDSY.Value = True Then
        ADS = "Yes"
    ElseIf optDSN.Value = True Then
        ADS = "No"
    End If
End Sub

Public Sub Read_record()

'Attest in Photograph

If EAIP = "Yes" Then
    optApY.Value = True
ElseIf EAIP = "No" Then
    optApN.Value = True
End If

'Photograph

If EPGP = "Yes" Then
    optPY.Value = True
ElseIf EPGP = "No" Then
    optPN.Value = True
End If

'photoID

If EPID = "Yes" Then
    optPIY.Value = True
ElseIf EPID = "No" Then
    optPIN.Value = True
End If

'Coustomer Signature in APEF

If ECSAPEF = "Yes" Then
    optAEFY.Value = True
ElseIf ECSAPEF = "No" Then
    optAEFN.Value = True
End If

'RetailSeal

If ERS = "Yes" Then
    optRY.Value = True
ElseIf ERS = "No" Then
    optRN.Value = True
End If

'Distributer Seal

If EDS = "Yes" Then
    optDSY.Value = True
ElseIf EDS = "No" Then
    optDSN.Value = True
End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 1 Then
    Dim oldActivation As aActivation
If txtmobileNo.Text = "" Then
SSTab1.Tab = 0
MsgBox "Please Enter the Mobile no", vbExclamation
txtmobileNo.SetFocus
Exit Sub
End If
If DTPActivation.Value <> Date Then
If MsgBox(" Select Date is Incorrect. Do you want to Continue..? ", vbYesNo + vbExclamation, "Warning") = vbYes Then
Else
SSTab1.Tab = 0
Exit Sub
End If
End If
If mFormState = "add" Then
If GetActivationNo(txtmobileNo.Text, oldActivation) = True Then
        MsgBox "The Activation MobileNo that you have entered is already exist." & vbNewLine & vbNewLine & _
            "Please enter different value.", vbExclamation
            SSTab1.Tab = 0
        HLTxt txtmobileNo
        Exit Sub
    End If
End If
SSTab1.Tab = 1

ElseIf SSTab1.Tab = 2 Then

If optApY.Value = False And optApN.Value = False Then
SSTab1.Tab = 1
MsgBox "Mark the detils about Attest in photograph ", vbExclamation
Exit Sub
End If

If optPY.Value = False And optPN.Value = False Then
SSTab1.Tab = 1
MsgBox "Mark the detils about photograph ", vbExclamation
Exit Sub
End If

If optPIY.Value = False And optPIN.Value = False Then
SSTab1.Tab = 1
MsgBox "Mark the detils about Photo ID ", vbExclamation
Exit Sub
End If

If optAEFY.Value = False And optAEFN.Value = False Then
SSTab1.Tab = 1
MsgBox "Mark the detils about Coustmer Signature in APEF ", vbExclamation
Exit Sub
End If

If optRY.Value = False And optRN.Value = False Then
SSTab1.Tab = 1
MsgBox "Mark the detils about Retail Seal", vbExclamation
Exit Sub
End If



If optDSY.Value = False And optDSN.Value = False Then
SSTab1.Tab = 1
MsgBox "Mark the detils about Distributer Seal", vbExclamation
Exit Sub
End If

    
SSTab1.Tab = 2

End If
End Sub

Private Sub Timer1_Timer()
DTDate.Value = Date
End Sub

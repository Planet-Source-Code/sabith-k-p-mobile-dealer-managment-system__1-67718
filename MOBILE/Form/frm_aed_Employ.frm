VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_aed_Employ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employ"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   2775
      Left            =   -120
      ScaleHeight     =   2715
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   360
         Top             =   1680
      End
      Begin VB.CommandButton cmdCancel 
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
         Left            =   4800
         TabIndex        =   4
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
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
         Left            =   3720
         TabIndex        =   3
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   960
         TabIndex        =   2
         Top             =   960
         Width           =   4815
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   960
         TabIndex        =   1
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   3960
         TabIndex        =   6
         Top             =   120
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTEmploy 
         Height          =   330
         Left            =   4200
         TabIndex        =   9
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Format          =   49348609
         CurrentDate     =   39079
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000004&
         X1              =   120
         X2              =   5880
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label3 
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
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         TabIndex        =   7
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label1 
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
         Left            =   3480
         TabIndex        =   5
         Top             =   120
         Width           =   195
      End
   End
End
Attribute VB_Name = "frm_aed_Employ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim SEmploy As aEmploy
Dim mShowAdd As Boolean
Dim mShowEdit As Boolean
Dim mFormState As String

Public Function ShowForm() As Boolean

    'show form
    Me.Show vbModal
    
    ShowForm = mShowForm
End Function
    
Public Function ShowEdit(ID As String) As Boolean
    
    'set form state
    mFormState = "edit"
    
    'set parameter
    SEmploy.ID = ID
    
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

Private Sub Form_Activate()
Select Case mFormState
        Case "add"
        
            'set caption
            Me.Caption = "Add Employ"
            Me.cmdSave.Caption = "&Save"
                txtID.Text = modFunction.ComNumZ(GetNewID, 10)
        Case "edit"
            'get info
            If GetEmployNo(SEmploy.ID, SEmploy) = False Then
                'show failed
                MsgBox "User entry with User ID : '" & SEmploy.ID & "' does not exist.", vbExclamation
                'close this form
                Unload Me
                Exit Sub
            End If
            
            'set form ui info
            txtID.Text = SEmploy.ID
            txtName.Text = SEmploy.Name
            txtAddress.Text = SEmploy.Address
            DTEmploy.Value = SEmploy.eDate
            'set caption
            Me.Caption = "Edit Employ"
            Me.cmdSave.Caption = "&Update"
            
            txtID.Locked = True
            
    End Select
txtName.SetFocus
    End Sub


Private Function SaveAdd()
    Dim NewEmploy As aEmploy
    Dim oldEmploy As aEmploy
    
    
    'check form field
    If IsEmpty(txtID.Text) Then
        MsgBox "Please Enter EmployId", vbExclamation
        HLTxt txtID
        Exit Function
    End If
    If IsEmpty(txtName.Text) Then
        MsgBox "Please Enter the EmployName", vbExclamation
        HLTxt txtName
        Exit Function
    End If
    If IsEmpty(txtAddress.Text) Then
        MsgBox "Please Enter the Address", vbExclamation
        HLTxt txtAddress
        Exit Function
    End If
    
    If mFormState = "add" Then
       
    'check duplication
    If GetEmployNo(txtID.Text, oldEmploy) = True Then
        MsgBox "The EmployID that you have entered is already exist." & vbNewLine & vbNewLine & _
            "Please enter different value.", vbExclamation
        HLTxt txtID
        Exit Function
    End If
    NewEmploy.ID = txtID.Text
    NewEmploy.Name = txtName.Text
    NewEmploy.Address = txtAddress.Text
       NewEmploy.eDate = DTEmploy
    
      
    'try
    
    If ModRsEmploy.AddEmploy(NewEmploy) = True Then
        MsgBox "New Employ entry was successfully created.", vbInformation
        'set flag
        mShowAdd = True
        'close form and return
        Unload Me
        
        
    Else
    
        MsgBox "Unable to add new Employ entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    End If
End Function

Private Function SaveEdit()

    Dim NewEmploy As aEmploy
    Dim oldEmploy As aEmploy
    
    'check form field
    
    
      If IsEmpty(txtID.Text) Then
        MsgBox "Please enter Employ Id", vbExclamation
        HLTxt txtID
        Exit Function
    End If
    If IsEmpty(txtName.Text) Then
        MsgBox "Please Enter the EmployName", vbExclamation
        HLTxt txtName
        Exit Function
    End If
    
    'set new
   NewEmploy.ID = txtID.Text
    NewEmploy.Name = txtName.Text
    NewEmploy.Address = txtAddress.Text
        NewEmploy.eDate = DTEmploy
    'try
    'add new Employ
    If ModRsEmploy.EditEmploy(NewEmploy) = True Then
        MsgBox "Employ entry was updated.", vbInformation
        'set flag
        mShowEdit = True
        'close form and return
        Unload Me
        
    Else
    
        MsgBox "Unable to update Employ entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    
End Function





Public Function GetNewID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetNewID = -1
    
    sSQL = "SELECT Max(tblemploy.ID)+1 AS ID" & _
            " From tblemploy"

    
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


Private Sub Timer1_Timer()
DTEmploy.Value = Now
End Sub



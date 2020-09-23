VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_aed_Sim 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sim"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   5610
      Left            =   0
      ScaleHeight     =   5580
      ScaleWidth      =   5370
      TabIndex        =   0
      Top             =   0
      Width           =   5400
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   240
         Top             =   360
      End
      Begin VB.TextBox txtTotalSim 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1560
         TabIndex        =   4
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   3360
         TabIndex        =   8
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   " &Save"
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
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3480
         Width           =   855
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
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1560
         TabIndex        =   5
         Top             =   2880
         Width           =   1935
      End
      Begin VB.ComboBox cmbEmployName 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   960
         Width           =   2055
      End
      Begin VB.ComboBox cmbOutletName 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1440
         Width           =   2055
      End
      Begin VB.ComboBox cmbProductName 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1920
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTSim 
         Height          =   330
         Left            =   3480
         TabIndex        =   9
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Format          =   20316161
         CurrentDate     =   39079
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000004&
         X1              =   0
         X2              =   5280
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total No of Sim"
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
         Left            =   135
         TabIndex        =   15
         Top             =   2400
         Width           =   1260
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
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
         TabIndex        =   14
         Top             =   1920
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Left            =   3060
         TabIndex        =   13
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
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
         TabIndex        =   12
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1050
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         TabIndex        =   10
         Top             =   2880
         Width           =   675
      End
   End
End
Attribute VB_Name = "frm_aed_Sim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim spSim As pSim
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
    spSim.ID = ID
    
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
            Me.Caption = "Add Sim"
            Me.cmdSave.Caption = "&Save"
            
        Case "edit"
            'get info
            If GetpSimNo(spSim.ID, spSim) = False Then
                'show failed
                MsgBox "User entry with User ID : '" & spSim.ID & "' does not exist.", vbExclamation
                'close this form
                Unload Me
                Exit Sub
            End If
            
            'set form ui info
            txtID.Text = spSim.ID
           cmbEmployName.Text = spSim.EmployName
            cmbOutletName.Text = spSim.OutletName
            cmbProductName.Text = spSim.ProductName
            txtAmount.Text = spSim.Amount
            txtTotalSim.Text = spSim.NoofpSim
            DTSim.Value = spSim.eDate
            'set caption
            Me.Caption = "Edit Sim"
            Me.cmdSave.Caption = "&Update"
            
            txtID.Locked = True
            
    
    End Select
    
End Sub
Private Sub Form_Load()
        ComboBuilder "tblproduct", "Name", "Id", cmbProductName
        ComboBuilder "tblEmploy", "Name", "Id", cmbEmployName
        ComboBuilder "tbloutlet", "Name", "id", cmbOutletName
End Sub

    


Private Function SaveAdd()
    Dim NewpSim As pSim
    Dim oldpSim As pSim
    
    
    'check form field
    If IsEmpty(txtID.Text) Then
        MsgBox "Please Enter pSimId", vbExclamation
        HLTxt txtID
        Exit Function
    End If
    If IsEmpty(cmbEmployName.Text) Then
        MsgBox "Please Enter the EmployName", vbExclamation
        HLTxt cmbEmployName
        Exit Function
    End If
    If IsEmpty(cmbOutletName.Text) Then
        MsgBox "Please Enter the Outlet", vbExclamation
        HLTxt cmbOutletName
        Exit Function
    End If
    
    If mFormState = "add" Then
       
    'check duplication
    If GetpSimNo(txtID.Text, oldpSim) = True Then
        MsgBox "The pSimID that you have entered is already exist." & vbNewLine & vbNewLine & _
            "Please enter different value.", vbExclamation
        HLTxt txtID
        Exit Function
    End If
    NewpSim.ID = txtID.Text
    NewpSim.EmployName = cmbEmployName.Text
    NewpSim.ProductName = cmbProductName.Text
    NewpSim.OutletName = cmbOutletName.Text
    NewpSim.NoofpSim = txtTotalSim.Text
    NewpSim.Amount = txtAmount.Text
    
       NewpSim.eDate = DTSim.Value
    
      
    'try
    
    If ModRsSim.AddpSim(NewpSim) = True Then
        MsgBox "New Sim entry was successfully created.", vbInformation
        'set flag
        mShowAdd = True
        'close form and return
        Unload Me
        
        
    Else
    
        MsgBox "Unable to add new pSim entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    End If
End Function

Private Function SaveEdit()

    Dim NewpSim As pSim
    Dim oldpSim As pSim
    
    'check form field
    
    
      If IsEmpty(txtID.Text) Then
        MsgBox "Please enter pSim Id", vbExclamation
        HLTxt txtID
        Exit Function
    End If
    If IsEmpty(cmbEmployName.Text) Then
        MsgBox "Please Enter the EmployName", vbExclamation
        HLTxt cmbEmployName
        Exit Function
    End If
    
    'set new
  NewpSim.ID = txtID.Text
    NewpSim.EmployName = cmbEmployName.Text
    NewpSim.ProductName = cmbProductName.Text
    NewpSim.OutletName = cmbOutletName.Text
    NewpSim.NoofpSim = txtTotalSim.Text
    NewpSim.Amount = txtAmount.Text
    
       NewpSim.eDate = DTSim.Value
    
    'try
    'add new pSim
    If ModRsSim.EdipSim(NewpSim) = True Then
        MsgBox "Sim entry was updated.", vbInformation
        'set flag
        mShowEdit = True
        'close form and return
        Unload Me
        
    Else
    
        MsgBox "Unable to update Sim entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    
End Function






Private Sub Timer1_Timer()
DTSim.Value = Now
End Sub









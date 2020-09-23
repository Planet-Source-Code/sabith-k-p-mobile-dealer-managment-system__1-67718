VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_aed_Coupon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Coupon"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   5610
      Left            =   0
      ScaleHeight     =   5580
      ScaleWidth      =   7530
      TabIndex        =   0
      Top             =   0
      Width           =   7560
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   480
         Top             =   3720
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1920
         TabIndex        =   5
         Top             =   2880
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00E0E0E0&
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
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00E0E0E0&
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   3240
         TabIndex        =   8
         Text            =   " "
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox txtTotalCoupon 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1920
         TabIndex        =   4
         Top             =   2400
         Width           =   1455
      End
      Begin VB.ComboBox cmbEmploy 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cmbOutlet 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1440
         Width           =   2295
      End
      Begin VB.ComboBox cmbProduct 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1920
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTCoupon 
         Height          =   330
         Left            =   3600
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
         X1              =   0
         X2              =   5880
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   " Amount"
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
         TabIndex        =   15
         Top             =   2880
         Width           =   720
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
         TabIndex        =   14
         Top             =   1440
         Width           =   1050
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
         TabIndex        =   13
         Top             =   960
         Width           =   1350
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
         Left            =   2850
         TabIndex        =   12
         Top             =   120
         Width           =   195
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
         TabIndex        =   11
         Top             =   1920
         Width           =   1185
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total no of Coupon"
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
         Top             =   2400
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frm_aed_Coupon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim SCoupon As aCoupon
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
    SCoupon.ID = ID
    
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
            Me.Caption = "Add Coupon"
            Me.cmdSave.Caption = "&Save"
            
        Case "edit"
            'get info
            If GetCouponNo(SCoupon.ID, SCoupon) = False Then
                'show failed
                MsgBox "User entry with User ID : '" & SCoupon.ID & "' does not exist.", vbExclamation
                'close this form
                Unload Me
                Exit Sub
            End If
            
            'set form ui info
            txtID.Text = SCoupon.ID
           cmbEmploy.Text = SCoupon.EmployName
            cmbOutlet.Text = SCoupon.OutletName
            cmbProduct.Text = SCoupon.ProductName
            txtAmount.Text = SCoupon.Amount
            txtTotalCoupon = SCoupon.NoofCoupon
            DTCoupon.Value = SCoupon.eDate
            'set caption
            Me.Caption = "Edit Coupon"
            Me.cmdSave.Caption = "&Update"
            
            txtID.Locked = True
            
    
    End Select
    cmbEmploy.SetFocus
End Sub
Private Sub Form_Load()
        ComboBuilder "tblproduct", "Name", "Id", cmbProduct
        ComboBuilder "tblEmploy", "Name", "Id", cmbEmploy
        ComboBuilder "tbloutlet", "Name", "id", cmbOutlet
End Sub

    


Private Function SaveAdd()
    Dim NewCoupon As aCoupon
    Dim oldCoupon As aCoupon
    
    
    'check form field
    If IsEmpty(txtID.Text) Then
        MsgBox "Please Enter CouponId", vbExclamation
        HLTxt txtID
        Exit Function
    End If
    If IsEmpty(cmbEmploy.Text) Then
        MsgBox "Please Enter the EmployName", vbExclamation
        HLTxt cmbEmploy
        Exit Function
    End If
    If IsEmpty(cmbOutlet.Text) Then
        MsgBox "Please Enter the Outlet", vbExclamation
        HLTxt cmbOutlet
        Exit Function
    End If
    
    If mFormState = "add" Then
       
    'check duplication
    If GetCouponNo(txtID.Text, oldCoupon) = True Then
        MsgBox "The CouponID that you have entered is already exist." & vbNewLine & vbNewLine & _
            "Please enter different value.", vbExclamation
        HLTxt txtID
        Exit Function
    End If
    NewCoupon.ID = txtID.Text
    NewCoupon.EmployName = cmbEmploy.Text
    NewCoupon.ProductName = cmbProduct.Text
    NewCoupon.OutletName = cmbOutlet.Text
    NewCoupon.NoofCoupon = txtTotalCoupon.Text
    NewCoupon.Amount = txtAmount.Text
    
       NewCoupon.eDate = DTCoupon.Value
    
      
    'try
    
    If modRsCoupon.AddCoupon(NewCoupon) = True Then
        MsgBox "New Coupon entry was successfully created.", vbInformation
        'set flag
        mShowAdd = True
        'close form and return
        Unload Me
        
        
    Else
    
        MsgBox "Unable to add new Coupon entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    End If
End Function

Private Function SaveEdit()

    Dim NewCoupon As aCoupon
    Dim oldCoupon As aCoupon
    
    'check form field
    
    
      If IsEmpty(txtID.Text) Then
        MsgBox "Please enter Coupon Id", vbExclamation
        HLTxt txtID
        Exit Function
    End If
    If IsEmpty(cmbEmploy.Text) Then
        MsgBox "Please Enter the EmployName", vbExclamation
        HLTxt cmbEmploy
        Exit Function
    End If
    
    'set new
  NewCoupon.ID = txtID.Text
    NewCoupon.EmployName = cmbEmploy.Text
    NewCoupon.ProductName = cmbProduct.Text
    NewCoupon.OutletName = cmbOutlet.Text
    NewCoupon.NoofCoupon = txtTotalCoupon.Text
    NewCoupon.Amount = txtAmount.Text
    
       NewCoupon.eDate = DTCoupon.Value
    
    'try
    'add new Coupon
    If modRsCoupon.EditCoupon(NewCoupon) = True Then
        MsgBox "Coupon entry was updated.", vbInformation
        'set flag
        mShowEdit = True
        'close form and return
        Unload Me
        
    Else
    
        MsgBox "Unable to update Coupon entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    
End Function








Private Sub Timer1_Timer()
DTCoupon.Value = Date
End Sub





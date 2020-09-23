VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_aed_Product 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Product"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   4170
      Left            =   0
      ScaleHeight     =   4140
      ScaleWidth      =   7530
      TabIndex        =   0
      Top             =   0
      Width           =   7560
      Begin VB.TextBox txtPurchase 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1440
         TabIndex        =   3
         Top             =   1920
         Width           =   1815
      End
      Begin VB.ComboBox cmbCategory 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   480
         Top             =   240
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1440
         TabIndex        =   4
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   3360
         TabIndex        =   7
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox txtProductName 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1440
         TabIndex        =   1
         Top             =   960
         Width           =   2415
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
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3120
         Width           =   975
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
         TabIndex        =   6
         Top             =   3120
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTProduct 
         Height          =   330
         Left            =   3600
         TabIndex        =   8
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Format          =   49348609
         CurrentDate     =   39079
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Rate"
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
         Top             =   2040
         Width           =   1230
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Catogory"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   780
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000004&
         X1              =   -240
         X2              =   5400
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Rate"
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
         Top             =   2640
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Product ID"
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
         Left            =   2280
         TabIndex        =   10
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label2 
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
         TabIndex        =   9
         Top             =   960
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frm_aed_Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim SProduct As aProduct
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
    SProduct.ID = ID
    
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
            Me.Caption = "Add Product"
            Me.cmdSave.Caption = "&Save"
                txtID.Text = modFunction.ComNumZ(GetNewID, 10)

        Case "edit"
            'get info
            If GetProductNo(SProduct.ID, SProduct) = False Then
                'show failed
                MsgBox "User entry with User ID : '" & SProduct.ID & "' does not exist.", vbExclamation
                'close this form
                Unload Me
                Exit Sub
            End If
            
            'set form ui info
            txtID.Text = SProduct.ID
            txtProductName.Text = SProduct.Name
           cmbCategory.Text = SProduct.Catogory
            txtAmount = SProduct.Amount
            txtPurchase = SProduct.PurchaseRate
            DTProduct.Value = SProduct.oDate
            'set caption
            Me.Caption = "Edit Product"
            Me.cmdSave.Caption = "&Update"
            
            txtID.Locked = True
            
    End Select
    txtProductName.SetFocus
    End Sub


Private Function SaveAdd()
    Dim NewProduct As aProduct
    Dim oldProduct As aProduct
    
    
    'check form field
    If IsEmpty(txtID.Text) Then
        MsgBox "Please Enter ProductId", vbExclamation
        HLTxt txtID
        Exit Function
    End If
    If IsEmpty(txtProductName.Text) Then
        MsgBox "Please Enter the ProductName", vbExclamation
        HLTxt txtProductName
        Exit Function
    End If
   
    
    If mFormState = "add" Then
       
    'check duplication
    If GetProductName(txtProductName.Text, oldProduct) = True Then
        MsgBox "The ProductName that you have entered is already exist." & vbNewLine & vbNewLine & _
            "Please enter different value.", vbExclamation
        HLTxt txtProductName
        Exit Function
    End If
    NewProduct.ID = txtID.Text
    NewProduct.Name = txtProductName.Text
    NewProduct.Catogory = cmbCategory.Text
    NewProduct.PurchaseRate = txtPurchase.Text
        NewProduct.Amount = txtAmount.Text
    NewProduct.oDate = DTProduct
    
      
    'try
    
    If modRsProduct.AddProduct(NewProduct) = True Then
        MsgBox "New Product entry was successfully created.", vbInformation
        'set flag
        mShowAdd = True
        'close form and return
        Unload Me
        
        
    Else
    
        MsgBox "Unable to add new Product entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    End If
End Function

Private Function SaveEdit()

    Dim NewProduct As aProduct
    Dim oldProduct As aProduct
    
    'check form field
    
    
      If IsEmpty(txtID.Text) Then
        MsgBox "Please enter Product Id", vbExclamation
        HLTxt txtID
        Exit Function
    End If
    If IsEmpty(txtProductName.Text) Then
        MsgBox "Please Enter the ProductName", vbExclamation
        HLTxt txtProductName
        Exit Function
    End If
    
    'set new
   NewProduct.ID = txtID.Text
    NewProduct.Name = txtProductName.Text
       NewProduct.Amount = txtAmount.Text
       NewProduct.Catogory = cmbCategory.Text
       NewProduct.PurchaseRate = txtPurchase.Text
    NewProduct.oDate = DTProduct
    'try
    'add new Product
    If modRsProduct.EditProduct(NewProduct) = True Then
        MsgBox "Product entry was updated.", vbInformation
        'set flag
        mShowEdit = True
        'close form and return
        Unload Me
        
    Else
    
        MsgBox "Unable to update Product entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    
End Function








Private Sub Form_Load()
'LoadCategory Me, cmbCategory

ComboBuilder "tblCategory", "Name", "ID", cmbCategory
End Sub

Private Sub Timer1_Timer()
DTProduct.Value = Now
End Sub
Public Function GetNewID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    'default
    GetNewID = -1
    
    sSQL = "SELECT Max(tblProduct.ID)+1 AS ID" & _
            " From tblProduct"

    
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


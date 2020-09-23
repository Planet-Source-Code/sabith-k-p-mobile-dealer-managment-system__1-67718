VERSION 5.00
Begin VB.Form frm_aed_Category 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Category"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   0
      ScaleHeight     =   2685
      ScaleWidth      =   7050
      TabIndex        =   0
      Top             =   0
      Width           =   7080
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtCategoryName 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1560
         TabIndex        =   2
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00C0C0C0&
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
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
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000004&
         X1              =   0
         X2              =   5040
         Y1              =   1200
         Y2              =   1200
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
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Category Name"
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
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frm_aed_Category"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim SCategory As aCategory
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
    SCategory.ID = ID
    
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
            Me.Caption = "Add Category"
            Me.cmdSave.Caption = "&Save"
           txtID.Text = modFunction.ComNumZ(GetNewID, 5)
        Case "edit"
            'get info
            If GetCategoryNo(SCategory.ID, SCategory) = False Then
                'show failed
                MsgBox "User entry with User ID : '" & SCategory.ID & "' does not exist.", vbExclamation
                'close this form
                Unload Me
                Exit Sub
            End If
            
            'set form ui info
            txtID.Text = SCategory.ID
            txtCategoryName.Text = SCategory.Name
            
            'set caption
            Me.Caption = "Edit Category"
            Me.cmdSave.Caption = "&Update"
            
            txtID.Locked = True
            
    End Select
    txtCategoryName.SetFocus
    End Sub


Private Function SaveAdd()
    Dim NewCategory As aCategory
    Dim oldCategory As aCategory
    
    
    'check form field
    If IsEmpty(txtID.Text) Then
        MsgBox "Please Enter CategoryId", vbExclamation
        HLTxt txtID
        Exit Function
    End If
    If IsEmpty(txtCategoryName.Text) Then
        MsgBox "Please Enter the CategoryName", vbExclamation
        HLTxt txtCategoryName
        Exit Function
    End If
    
    
    If mFormState = "add" Then
       
    'check duplication
    If GetCategoryNo(txtID.Text, oldCategory) = True Then
        MsgBox "The CategoryID that you have entered is already exist." & vbNewLine & vbNewLine & _
            "Please enter different value.", vbExclamation
        HLTxt txtID
        Exit Function
    End If
    NewCategory.ID = txtID.Text
    NewCategory.Name = txtCategoryName.Text
    
    
      
    'try
    
    If modRsCategory.AddCategory(NewCategory) = True Then
        MsgBox "New Category entry was successfully created.", vbInformation
        'set flag
        mShowAdd = True
        'close form and return
        Unload Me
        
        
    Else
    
        MsgBox "Unable to add new Category entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    End If
End Function

Private Function SaveEdit()

    Dim NewCategory As aCategory
    Dim oldCategory As aCategory
    
    'check form field
    
    
      If IsEmpty(txtID.Text) Then
        MsgBox "Please enter Category Id", vbExclamation
        HLTxt txtID
        Exit Function
    End If
    If IsEmpty(txtCategoryName.Text) Then
        MsgBox "Please Enter the CategoryName", vbExclamation
        HLTxt txtCategoryName
        Exit Function
    End If
    
    'set new
   NewCategory.ID = txtID.Text
    NewCategory.Name = txtCategoryName.Text
    
    'try
    'add new Category
    If modRsCategory.EditCategory(NewCategory) = True Then
        MsgBox "Category entry was updated.", vbInformation
        'set flag
        mShowEdit = True
        'close form and return
        Unload Me
        
    Else
    
        MsgBox "Unable to update Category entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    
End Function
Public Function GetNewID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetNewID = -1
    
    sSQL = "SELECT Max(tblCategory.ID)+1 AS ID" & _
            " From tblCategory"

    
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

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEsale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Easy Sale"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   0
      ScaleHeight     =   3675
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.ComboBox cmbOutlet 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtFinalAmount 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtID 
         Height          =   325
         Left            =   4080
         TabIndex        =   7
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox txtAmount 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   1815
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
         Left            =   4200
         TabIndex        =   4
         Top             =   1920
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
         Left            =   5160
         TabIndex        =   5
         Top             =   1920
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPESale 
         Height          =   330
         Left            =   4080
         TabIndex        =   6
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Format          =   50397185
         CurrentDate     =   39091
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         TabIndex        =   12
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Value"
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
         Top             =   1200
         Width           =   465
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
         Left            =   3600
         TabIndex        =   10
         Top             =   120
         Width           =   195
      End
      Begin VB.Label Label2 
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
         Left            =   3600
         TabIndex        =   9
         Top             =   720
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         TabIndex        =   8
         Top             =   840
         Width           =   675
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   6120
         Y1              =   1800
         Y2              =   1800
      End
   End
End
Attribute VB_Name = "frmEsale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SESale As aESale
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
    SESale.ID = ID
    
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
            Me.Caption = "Add ESale"
            Me.cmdSave.Caption = "&Save"
            txtID = modFunction.ComNumZ(GetNewID, 8)
            
        Case "edit"
            'get info
            'If GetESaleNo(SESale.ID, SESale) = False Then
                'show failed
               ' MsgBox "User entry with User ID : '" & SESale.ID & "' does not exist.", vbExclamation
                'close this form
               ' Unload Me
               ' Exit Sub
            'End If
            
            'set form ui info
            txtID.Text = SESale.ID
            DTPESale.Value = SESale.eDate
            cmbOutlet.Text = DTPESale.Value
            txtAmount.Text = SESale.eDate
            txtFinalAmount.Text = SESale.FinalAmount
                        'set caption
            Me.Caption = "Edit ESale"
            Me.cmdSave.Caption = "&Update"
            
            txtID.Locked = True
            
    End Select
    End Sub


Private Function SaveAdd()
    Dim NewESale As aESale
    Dim oldESale As aESale
    
    
    'check form field
    
    If IsEmpty(cmbOutlet.Text) Then
        MsgBox "Please Select the Outlet", vbExclamation
        HLTxt cmbOutlet
        Exit Function
    End If
    If IsEmpty(txtAmount.Text) Then
        MsgBox "Please Enter the Value", vbExclamation
        HLTxt txtAmount
        Exit Function
    End If
    
    If mFormState = "add" Then
       
    'check duplication
    
    NewESale.ID = txtID.Text
    NewESale.eDate = DTPESale.Value
    NewESale.OutletName = cmbOutlet.Text
    NewESale.Amount = txtAmount.Text
    NewESale.FinalAmount = txtFinalAmount.Text
    'try
    If modRsEsale.AddESale(NewESale) = True Then
        MsgBox "New EasySale entry was successfully created.", vbInformation
        'set flag
        mShowAdd = True
        'close form and return
        Unload Me
        frmEsale.ShowAdd
        
    Else
    
        MsgBox "Unable to add new ESale entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    End If
End Function

Private Function SaveEdit()

    Dim NewESale As aESale
    Dim oldESale As aESale
    
    'check form field
    
    
      
    If IsEmpty(txtAmount.Text) Then
        MsgBox "Please Enter the Amount", vbExclamation
        HLTxt txtAmount
        Exit Function
    End If
    
    'set new
  NewESale.ID = txtID.Text
    NewESale.eDate = DTPESale.Value
    NewESale.OutletName = cmbOutlet.Text
    NewESale.Amount = txtFinalAmount.Text
    NewESale.FinalAmount = txtFinalAmount.Text
    
    'try
    'add new ESale
    If modRsEsale.EditESale(NewESale) = True Then
        MsgBox "ESale entry was updated.", vbInformation
        'set flag
        mShowEdit = True
        'close form and return
        Unload Me
        
    Else
    
        MsgBox "Unable to update ESale entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    
End Function

Public Function GetNewID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetNewID = -1
    
    sSQL = "SELECT Max(tblESale.ID)+1 AS ID" & _
            " From tblESale"

    
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
End Sub


Private Sub Form_Load()
ComboBuilder "tblOutlet", "Name", "ID", cmbOutlet
DTPESale.Value = Date
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtFinalAmount.Text = Val(txtAmount) + (Val(txtAmount.Text) * 4 / 100)
txtFinalAmount.SetFocus
End If
End Sub

Private Sub txtFinalAmount_Change()

txtAmount.Text = Val(txtFinalAmount) / 1.04
End Sub

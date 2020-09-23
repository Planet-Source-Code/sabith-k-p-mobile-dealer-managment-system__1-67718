VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_aed_Outlet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Outlet"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6075
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
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   5400
         Top             =   1320
      End
      Begin VB.ComboBox cmbEmploy 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   960
         Width           =   2415
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
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2520
         Width           =   975
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
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txtOutletName 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   4080
         TabIndex        =   7
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox txtPlace 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1560
         TabIndex        =   3
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtPhoneNo 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1560
         TabIndex        =   4
         Top             =   1920
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTOutlet 
         Height          =   330
         Left            =   4320
         TabIndex        =   8
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Format          =   49348609
         CurrentDate     =   39079
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000004&
         X1              =   120
         X2              =   6240
         Y1              =   2400
         Y2              =   2400
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
         TabIndex        =   13
         Top             =   480
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
         TabIndex        =   12
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
         Left            =   3840
         TabIndex        =   11
         Top             =   120
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Place"
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
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No"
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
         Top             =   1920
         Width           =   915
      End
   End
End
Attribute VB_Name = "frm_aed_Outlet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Soutlet As aOutlet
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
    Soutlet.ID = ID
    
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
            Me.Caption = "Add outlet"
            Me.cmdSave.Caption = "&Save"
            txtID = modFunction.ComNumZ(GetNewID, 8)
            
        Case "edit"
            'get info
            If GetoutletNo(Soutlet.ID, Soutlet) = False Then
                'show failed
                MsgBox "User entry with User ID : '" & Soutlet.ID & "' does not exist.", vbExclamation
                'close this form
                Unload Me
                Exit Sub
            End If
            
            'set form ui info
            txtID.Text = Soutlet.ID
            txtOutletName.Text = Soutlet.Name
            cmbEmploy.Text = Soutlet.EmployName
            txtPhoneNo = Soutlet.PhoneNo
            txtPlace = Soutlet.Place
            'set caption
            Me.Caption = "Edit outlet"
            Me.cmdSave.Caption = "&Update"
            
            txtID.Locked = True
            
    End Select
    txtOutletName.SetFocus
    End Sub


Private Function SaveAdd()
    Dim Newoutlet As aOutlet
    Dim oldoutlet As aOutlet
    
    
    'check form field
    If IsEmpty(txtID.Text) Then
        MsgBox "Please Enter outletId", vbExclamation
        HLTxt txtID
        Exit Function
    End If
    If IsEmpty(txtOutletName.Text) Then
        MsgBox "Please Enter the outletName", vbExclamation
        HLTxt txtOutletName
        Exit Function
    End If
    If IsEmpty(cmbEmploy.Text) Then
        MsgBox "Please select the employ", vbExclamation
        HLTxt cmbEmploy
        Exit Function
    End If
    
    If mFormState = "add" Then
       
    'check duplication
    If GetoutletNo(txtID.Text, oldoutlet) = True Then
        MsgBox "The outletID that you have entered is already exist." & vbNewLine & vbNewLine & _
            "Please enter different value.", vbExclamation
        HLTxt txtID
        Exit Function
    End If
    Newoutlet.ID = txtID.Text
    Newoutlet.Name = txtOutletName.Text
    Newoutlet.EmployName = cmbEmploy.Text
    Newoutlet.Place = txtPlace.Text
    Newoutlet.Date = DTOutlet
    Newoutlet.PhoneNo = txtPhoneNo.Text
      
    'try
    
    If ModRsOutlet.Addoutlet(Newoutlet) = True Then
        MsgBox "New outlet entry was successfully created.", vbInformation
        'set flag
        mShowAdd = True
        'close form and return
        Unload Me
        
        
    Else
    
        MsgBox "Unable to add new outlet entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    End If
End Function

Private Function SaveEdit()

    Dim Newoutlet As aOutlet
    Dim oldoutlet As aOutlet
    
    'check form field
    
    
      If IsEmpty(txtID.Text) Then
        MsgBox "Please enter outlet Id", vbExclamation
        HLTxt txtID
        Exit Function
    End If
    If IsEmpty(txtOutletName.Text) Then
        MsgBox "Please Enter the outletName", vbExclamation
        HLTxt txtOutletName
        Exit Function
    End If
    
    'set new
   Newoutlet.ID = txtID.Text
    Newoutlet.Name = txtOutletName.Text
    Newoutlet.EmployName = cmbEmploy.Text
    Newoutlet.Place = txtPlace
    Newoutlet.PhoneNo = txtPhoneNo
    Newoutlet.Date = DTOutlet
    'try
    'add new outlet
    If ModRsOutlet.Editoutlet(Newoutlet) = True Then
        MsgBox "outlet entry was updated.", vbInformation
        'set flag
        mShowEdit = True
        'close form and return
        Unload Me
        
    Else
    
        MsgBox "Unable to update outlet entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    
End Function

Public Function GetNewID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetNewID = -1
    
    sSQL = "SELECT Max(tbloutlet.ID)+1 AS ID" & _
            " From tbloutlet"

    
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




Private Sub Form_Load()
ComboBuilder "tblemploy", "Name", "Id", cmbEmploy
End Sub

Private Sub Timer1_Timer()
DTOutlet.Value = Date
End Sub

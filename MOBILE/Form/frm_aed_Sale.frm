VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_aed_Sale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sale"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   5445
      Left            =   0
      ScaleHeight     =   5415
      ScaleWidth      =   6330
      TabIndex        =   0
      Top             =   0
      Width           =   6360
      Begin VB.CommandButton cmdremove 
         Caption         =   "RemoveRow"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   5640
         TabIndex        =   5
         Top             =   1080
         Width           =   495
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   480
         Top             =   4680
      End
      Begin VB.ComboBox cmbEmployName 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox cmbProductName 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtQty 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   2280
         TabIndex        =   3
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   120
         Width           =   1575
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
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4680
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox txtTotalAmount 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox txtSaleRs 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdNewItem 
         Height          =   315
         Left            =   1920
         Picture         =   "frm_aed_Sale.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "New Item Entry"
         Top             =   1080
         Width           =   285
      End
      Begin MSFlexGridLib.MSFlexGrid Flx1 
         Height          =   2415
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4260
         _Version        =   393216
         Cols            =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTSale 
         Height          =   330
         Left            =   4440
         TabIndex        =   13
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Format          =   49283073
         CurrentDate     =   39079
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         TabIndex        =   21
         Top             =   120
         Width           =   1350
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
         TabIndex        =   20
         Top             =   840
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
         Left            =   4200
         TabIndex        =   19
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total  Amount"
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
         Left            =   2880
         TabIndex        =   18
         Top             =   4080
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
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
         TabIndex        =   16
         Top             =   840
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs #"
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
         Left            =   3360
         TabIndex        =   15
         Top             =   840
         Width           =   390
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         Left            =   4440
         TabIndex        =   14
         Top             =   840
         Width           =   675
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000004&
         X1              =   9600
         X2              =   0
         Y1              =   4560
         Y2              =   4560
      End
   End
End
Attribute VB_Name = "frm_aed_Sale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public TotalStock As Long
  



Private Sub cmbEmployName_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 And cmbEmployName.Text <> "" Then
    cmbProductName.SetFocus
Else
    MsgBox "Enter the Employ Name", vbExclamation
End If

End Sub

Private Sub cmbProductName_Change()
Dim vRS As New ADODB.Recordset
Dim sSQL As String
sSQL = "SELECT * FROM tblProduct WHERE Name='" & cmbProductName.Text & "'"

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "AddAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        'GoTo RAE
    End If
txtSaleRs.Text = vRS.Fields("Amount")
End Sub

Private Sub cmbProductName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And cmbProductName.Text <> "" Then
Dim vRS As New ADODB.Recordset
Dim sSQL As String
sSQL = "SELECT * FROM tblProduct WHERE Name='" & cmbProductName.Text & "'"

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "AddAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        'GoTo RAE
    End If
txtSaleRs.Text = vRS.Fields("Amount")
txtQty.SetFocus
Else
MsgBox "Please select the Product Name", vbExclamation
End If
End Sub

Private Sub cmbProductName_Validate(Cancel As Boolean)
On Error Resume Next
Dim vRS As New ADODB.Recordset
Dim sSQL As String
sSQL = "SELECT * FROM tblProduct WHERE Name='" & cmbProductName.Text & "'"

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "AddAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        'GoTo RAE
    End If
txtSaleRs.Text = vRS.Fields("Amount")

End Sub

Private Sub cmdAdd_Click()
StockManage
If txtQty.Text = "" Or txtSaleRs.Text = "" Or cmbProductName.Text = "" Then
MsgBox "Some Fields are Empty, Please Check It!", vbExclamation
Exit Sub
End If
If Me.Caption = "New Sale" Then
If txtQty.Text > TotalStock Then
    If MsgBox("Stock  Only " & TotalStock & "   Do you want to Continue?", vbYesNo + vbExclamation) = vbNo Then
    HLTxt txtQty
    txtQty.SetFocus
        Exit Sub
    End If
End If
End If
Row = Flx1.Rows - 1
With Flx1

        .Rows = .Rows + 1
        
        .TextMatrix(Row, 1) = cmbProductName
        .TextMatrix(Row, 2) = txtQty
        .TextMatrix(Row, 3) = txtSaleRs
        .TextMatrix(Row, 4) = txtAmount
         
         
        txtAmount = ""
        txtQty = ""
        txtSaleRs.Text = ""
        cmbProductName.SetFocus
  End With
  CalcTotalAmount

  

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdNewItem_Click()
frm_aed_Product.ShowAdd
End Sub

Private Sub cmdremove_Click()
On Error Resume Next
Flx1.RemoveItem Flx1.RowSel
CalcTotalAmount
End Sub

Private Sub cmdSave_Click()
'On Error GoTo DatabaseError:
Dim SEmployName, SCategoryName, SProductName, SRate, SAmount, PQTY, PTotal As String
If txtTotalAmount.Text = "" Or Val(txtTotalAmount.Text) = 0 Then
MsgBox "No Sales data Found", vbCritical
cmbEmployName.SetFocus
Exit Sub
End If
If Me.Caption = "New Sale" Then
'PrimeDB.Execute ("INSERT INTO  tblPurchaseBill (pDate,BillNo,Naration,TotalAmount,Discount,GrandTotal) values (" & "'" & DtP1 & "'" & "," & "'" & txtInvoiceNo & "'" & "," & "'" & txtNarration & "'" & "," & "'" & txtTotalAmount & "'" & "," & "'" & txtDiscount & "'" & "," & "'" & txtGrandTotal & "'" & ")")
With Flx1
For i = 1 To .Rows - 2
.Row = i
   
    .Col = 1                 ''''''''''''''Save Purchase
    SProductName = .Text
    .Col = 2
    SQty = .Text
    .Col = 3
    SRate = .Text
    .Col = 4
    SAmount = .Text
     
    
PrimeDB.Execute ("INSERT INTO  tblsale (ID,sDate,EmployName,productName,Qty,Rate,Amount,TotalAmount) values (" & "'" & txtID & "'" & "," & "'" & DTSale.Value & "'" & "," & "'" & cmbEmployName & "'" & "," & "'" & SProductName & "'" & "," & SQty & "," & "'" & SRate & "'" & "," & "'" & SAmount & "'" & "," & "'" & txtTotalAmount & "'" & ")")

'PrimeDB.Execute ("INSERT INTO tblMainstockmng (Scode,BillNo,itemName,Qty ) values(" & "'" & txtPcode & "'" & "," & "'" & txtInvoiceNo & "'" & "," & "'" & PItemName & "'" & "," & "'" & PQTY & "'" & ")")
    Next i
    MsgBox "Record Save Succsesfuly", vbInformation
        End With
        ClearFlex Flx1
       ClearTExt
       txtID.Text = modFunction.ComNumZ(GetNewID, 12)
     Exit Sub
     
Else
With Flx1
For i = 1 To .Rows - 2
.Row = i
   
    .Col = 1                 ''''''''''''''Save Purchase
    SProductName = .Text
    .Col = 2
    SQty = .Text
    .Col = 3
    SRate = .Text
    .Col = 4
    SAmount = .Text
     
    
PrimeDB.Execute ("INSERT INTO  tblsale (ID,sDate,EmployName,productName,Qty,Rate,Amount,TotalAmount) values (" & "'" & txtID & "'" & "," & "'" & DTSale & "'" & "," & "'" & cmbEmployName & "'" & "," & "'" & SProductName & "'" & "," & "'" & SQty & "'" & "," & "'" & SRate & "'" & "," & "'" & SAmount & "'" & "," & "'" & txtTotalAmount & "'" & ")")

'PrimeDB.Execute ("INSERT INTO tblMainstockmng (Scode,BillNo,itemName,Qty ) values(" & "'" & txtPcode & "'" & "," & "'" & txtInvoiceNo & "'" & "," & "'" & PItemName & "'" & "," & "'" & PQTY & "'" & ")")
    Next i
    MsgBox "Record Save Succsesfuly", vbInformation
        End With
         ClearFlex Flx1
       ClearTExt
       txtID.Text = modFunction.ComNumZ(GetNewIID, 8)
     Exit Sub
     End If
DatabaseError:
        MsgBox "Unhandeled error is occured,Please Check the Bill No and Try Agin" & vbCrLf & Err.Description, vbCritical
       
        
End Sub










Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()

If Me.Caption = "New Sale" = True Then
    txtID.Text = modFunction.ComNumZ(GetNewID, 12)
Else
    txtID.Text = modFunction.ComNumZ(GetNewIID, 8)
End If

cmbEmployName.SetFocus

End Sub

Private Sub Form_Load()
ComboBuilder "tblEmploy", "Name", "ID", cmbEmployName
ComboBuilder "tblProduct", "Name", "ID", cmbProductName
Call Heading

End Sub

Private Sub Timer1_Timer()
DTSale.Value = Date
End Sub

Private Sub txtQty_Change()
CalcRate
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtQty.Text = "" Then
        MsgBox "Please enter the Qty", vbExclamation
        txtQty.SetFocus
        End If
        CalcRate
cmdAdd.SetFocus
  End If
  CheckAsci txtQty, KeyAscii
End Sub

Public Function CalcTotalAmount()
 Amt = 0
         For i = 1 To Flx1.Rows - 1
         Amt = Amt + Val(Flx1.TextMatrix(i, 4))
         Next i
         
         txtTotalAmount = Amt
End Function


Public Sub TotAmount()
txtAmount.Text = Val(txtQty) * Val(txtSaleRs)
End Sub

Public Sub Heading()
Flx1.Col = 0
Flx1.Row = 0
Flx1.ColWidth(0) = 600
Flx1.Col = 1
Flx1.ColWidth(1) = 800
Flx1.Text = "Product Name"
Flx1.ColAlignment(1) = 2
Flx1.Col = 2
Flx1.Text = "Qty"
Flx1.ColWidth(2) = 1000
Flx1.ColAlignment(2) = 2
Flx1.Col = 3
Flx1.Text = "Rate#"
Flx1.ColWidth(3) = 1200
Flx1.ColAlignment(3) = 2
Flx1.Col = 4
Flx1.Text = "Amount"
Flx1.ColWidth(4) = 1400
Flx1.ColAlignment(4) = 2
End Sub
Public Function GetNewID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetNewID = -1
    
    sSQL = "SELECT Max(tblsale.ID)+1 AS ID" & _
            " From tblsale"

    
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
Public Function GetNewIID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetNewIID = -1
    
    sSQL = "SELECT Max(tblsalereturn.ID)+1 AS ID" & _
            " From tblsalereturn"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        MsgBox "GetNewIID" & "," & "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetNewIID = 1
        GoTo RAE
    End If
    
    On Error Resume Next
    GetNewIID = ReadField(vRS.Fields("ID"))
    
    If GetNewIID < 1 Then
        GetNewIID = 1
        txtID.Text = ID
    End If
    
RAE:
    Set vRS = Nothing
    Err.Clear
End Function
Public Sub StockManage()
On Error Resume Next
Dim vRS As New ADODB.Recordset
Dim vRS1 As New ADODB.Recordset
Dim vRS2 As New ADODB.Recordset
Dim vRS3 As New ADODB.Recordset

    Dim sSQL As String
    Dim sSQL1 As String
    Dim sSQL2 As String
    Dim sSQL3 As String
    Dim sale As Long
    Dim SaleReturn As Long
    Dim Issue  As Long
    Dim IssueReturn As Long
    Dim totalStck As Long
    
    sSQL = "SELECT * FROM tblIssueQty WHERE ProductName='" & cmbProductName.Text & "'"
    sSQL1 = "SELECT * FROM tblissuereturnQty WHERE ProductName='" & cmbProductName.Text & "'"
    sSQL2 = "SELECT * FROM tblSaleQty WHERE ProductName='" & cmbProductName.Text & "'"
    sSQL3 = "SELECT * FROM tblSaleReturnQty WHERE ProductName='" & cmbProductName.Text & "'"

    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "AddAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        'GoTo RAE
    End If
           Issue = vRS.Fields("SumOfQty")
            
    If ConnectRS(PrimeDB, vRS1, sSQL1) = False Then
        'WriteErrorLog "modRAgent", "AddAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        'GoTo RAE
    End If
    IssueReturn = vRS1.Fields("Sum Of Qty")
    
    If ConnectRS(PrimeDB, vRS2, sSQL2) = False Then
        'WriteErrorLog "modRAgent", "AddAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        'GoTo RAE
    End If
    sale = vRS2.Fields("Sum Of Qty")
    If ConnectRS(PrimeDB, vRS3, sSQL3) = False Then
        'WriteErrorLog "modRAgent", "AddAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        'GoTo RAE
    End If
    SaleReturn = vRS3.Fields("Sum Of Qty")
    
    TotalStock = Issue - IssueReturn - sale + SaleReturn
    'MsgBox "Total" & "=" & Purchase & "-" & purchaseReturn & "-" & "(" & sale & "+" & SaleReturn & ")" & "=" & TotalStock

End Sub

Public Sub CalcRate()
txtAmount = Val(txtQty) * Val(txtSaleRs)
End Sub

Private Sub txtSaleRs_Change()
Call CalcRate
End Sub

Private Sub txtSaleRs_KeyPress(KeyAscii As Integer)
CheckAsci txtSaleRs, KeyAscii
End Sub
Public Sub ClearTExt()
'cmbEmployName.Text = ""
cmbEmployName.ListIndex = -1
txtTotalAmount.Text = ""
End Sub

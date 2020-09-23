VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPurchase 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   5685
      Left            =   -120
      ScaleHeight     =   5655
      ScaleWidth      =   9570
      TabIndex        =   0
      Top             =   0
      Width           =   9600
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3240
         Top             =   240
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Remove Row"
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
         Left            =   360
         TabIndex        =   20
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox txtInvoice 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdNewItem 
         Height          =   315
         Left            =   1920
         Picture         =   "frm_aed_Purchase.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "New Item Entry"
         Top             =   1200
         Width           =   285
      End
      Begin MSFlexGridLib.MSFlexGrid Flx1 
         Height          =   2550
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4498
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
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtPurchaseRs 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtTotalAmount 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   4680
         TabIndex        =   8
         Top             =   4200
         Width           =   1695
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
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4800
         Width           =   855
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
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4800
         Width           =   855
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox txtQty 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   2280
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ComboBox cmbProductName 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPurchase 
         Height          =   330
         Left            =   4680
         TabIndex        =   9
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Format          =   49414145
         CurrentDate     =   39079
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No"
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
         TabIndex        =   19
         Top             =   120
         Width           =   885
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000004&
         X1              =   9600
         X2              =   120
         Y1              =   4680
         Y2              =   4680
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
         Left            =   4680
         TabIndex        =   16
         Top             =   960
         Width           =   675
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
         Left            =   3480
         TabIndex        =   15
         Top             =   960
         Width           =   390
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
         TabIndex        =   14
         Top             =   960
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   45
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
         Left            =   3360
         TabIndex        =   12
         Top             =   4320
         Width           =   1200
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
         Left            =   4080
         TabIndex        =   11
         Top             =   240
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
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbProductName_Change()
Dim vRS As New ADODB.Recordset
Dim sSQL As String
sSQL = "SELECT * FROM tblProduct WHERE Name='" & cmbProductName.Text & "'"

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "AddAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        'GoTo RAE
    End If
txtPurchaseRs.Text = vRS.Fields("prvalue")
End Sub

Private Sub cmbProductName_Click()
Dim vRS As New ADODB.Recordset
Dim sSQL As String
sSQL = "SELECT * FROM tblProduct WHERE Name='" & cmbProductName.Text & "'"

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "AddAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        'GoTo RAE
    End If
txtPurchaseRs.Text = vRS.Fields("prvalue")

End Sub

Private Sub cmbProductName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtQty.SetFocus
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdNewItem_Click()
frm_aed_Product.ShowAdd
End Sub

Private Sub cmdSave_Click()
On Error GoTo DatabaseError:
Dim PCatogoryName, PProductName, PAmount, PQTY, PTotal, PRate As String
If txtTotalAmount.Text = "" Then
MsgBox "Purchase Data Not Found", vbExclamation
Exit Sub
End If

If Me.Caption = "New Purchase" Then
If txtInvoice = "" Then
MsgBox "Please Enter the Invoice No!", vbExclamation
txtInvoice.SetFocus
Exit Sub
End If
PrimeDB.Execute ("INSERT INTO  tblPurchaseBill (pDate,BillNo,TotalAmount) values (" & "'" & DTPurchase.Value & "'" & "," & "'" & txtInvoice & "'" & "," & "'" & txtTotalAmount & "'" & ")")
With Flx1
For i = 1 To .Rows - 2
.Row = i
    .Col = 1
    PProductName = .Text
    .Col = 2
    PQTY = .Text
    .Col = 3
    PRate = .Text
    .Col = 4
        PAmount = .Text
     
    
PrimeDB.Execute ("INSERT INTO  tblPurchase (ID,pDate,InvoiceNo,productName,Qty,Rate,Amount) values (" & "'" & txtID & "'" & "," & "'" & DTPurchase.Value & "'" & "," & "'" & txtInvoice & "'" & "," & "'" & PProductName & "'" & "," & "'" & PQTY & "'" & "," & "'" & PRate & "'" & "," & "'" & PAmount & "'" & ")")

'PrimeDB.Execute ("INSERT INTO tblMainstockmng (Scode,BillNo,itemName,Qty ) values(" & "'" & txtPcode & "'" & "," & "'" & txtInvoiceNo & "'" & "," & "'" & PItemName & "'" & "," & "'" & PQTY & "'" & ")")
    Next i
    MsgBox "Record Save Succsesfuly", vbInformation
        End With
        ClearFlex Flx1
        ClearTExt
        txtID.Text = modFunction.ComNumZ(GetNewID, 10)
     Exit Sub
     
Else
If txtInvoice = "" Then
MsgBox "Please Enter the Invoice No!", vbExclamation
txtInvoice.SetFocus
Exit Sub
End If
'PrimeDB.Execute ("INSERT INTO  tblPurchaseBill (pDate,BillNo,Naration,TotalAmount,Discount,GrandTotal) values (" & "'" & DtP1 & "'" & "," & "'" & txtInvoiceNo & "'" & "," & "'" & txtNarration & "'" & "," & "'" & txtTotalAmount & "'" & "," & "'" & txtDiscount & "'" & "," & "'" & txtGrandTotal & "'" & ")")
With Flx1
For i = 1 To .Rows - 2
.Row = i
    .Col = 1
    PProductName = .Text
    .Col = 2
    PQTY = .Text
    .Col = 3
    PRate = .Text
    .Col = 4
     PAmount = .Text
     
    
PrimeDB.Execute ("INSERT INTO  tblPurchasereturn (ID,pDate,InvoiceNo,productName,Qty,Rate,Amount,TotalAmount) values (" & "'" & txtID & "'" & "," & "'" & DTPurchase.Value & "'" & "," & "'" & txtInvoice & "'" & "," & "'" & PProductName & "'" & "," & "'" & PQTY & "'" & "," & "'" & PRate & "'" & "," & "'" & PAmount & "'" & "," & "'" & txtTotalAmount & "'" & ")")

'PrimeDB.Execute ("INSERT INTO tblMainstockmng (Scode,BillNo,itemName,Qty ) values(" & "'" & txtPcode & "'" & "," & "'" & txtInvoiceNo & "'" & "," & "'" & PItemName & "'" & "," & "'" & PQTY & "'" & ")")
    Next i
    MsgBox "Record Save Succsesfuly", vbInformation
        End With
        ClearFlex Flx1
        ClearTExt
        txtID = modFunction.ComNumZ(GetNewIID, 8)
     Exit Sub
End If
DatabaseError:
        MsgBox "Unhandeled error is occured,Please Check the Bill No and Try Agin" & vbCrLf & Err.Description, vbCritical
        txtInvoice.SetFocus
        HLTxt txtInvoice
End Sub



Private Sub Command2_Click()
On Error Resume Next
Flx1.RemoveItem Flx1.RowSel
CalcTotalAmount

End Sub

Private Sub Form_Activate()
txtInvoice.SetFocus
If Me.Caption = "New Purchase" Then
txtID.Text = modFunction.ComNumZ(GetNewID, 10)
Else
txtID.Text = modFunction.ComNumZ(GetNewIID, 8)
End If

End Sub

Private Sub Form_Load()
ComboBuilder "tblProduct", "Name", "ID", cmbProductName
Call Heading
End Sub

Private Sub Label9_Click()

End Sub

Private Sub Timer1_Timer()
DTPurchase.Value = Date
End Sub

Private Sub txtInvoice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbProductName.SetFocus
End If
End Sub

Private Sub txtPurchaseRs_Change()
Call TotAmount
End Sub

Private Sub txtQty_Change()
Call TotAmount
End Sub


Public Function CalcTotalAmount()
 Amt = 0
         For i = 1 To Flx1.Rows - 1
         Amt = Amt + Val(Flx1.TextMatrix(i, 4))
         Next i
         
         txtTotalAmount = Amt
End Function

Public Sub Heading()
Flx1.Col = 0
Flx1.Row = 0
Flx1.ColWidth(0) = 600
Flx1.Col = 1
Flx1.ColWidth(1) = 1200
Flx1.ColAlignment(1) = 2
Flx1.Text = "Product Name"
Flx1.Col = 2
Flx1.ColWidth(2) = 800
Flx1.Text = "Qty"
Flx1.Col = 3
Flx1.Text = "Rate"
Flx1.ColWidth(3) = 1000
Flx1.ColAlignment(3) = 2
Flx1.Col = 4
Flx1.Text = "Amount"
Flx1.ColWidth(4) = 1400
Flx1.ColAlignment(4) = 2
End Sub
Public Sub TotAmount()
txtAmount.Text = Val(txtQty) * Val(txtPurchaseRs)
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    If txtInvoice.Text = "" Then
    MsgBox "Please Enter the Invoice No", vbExclamation
    txtInvoice.SetFocus
    Exit Sub
    End If
    
     If cmbProductName.Text = "" Then
     MsgBox "Please Select the Product ", vbExclamation
    cmbProductName.SetFocus
    Exit Sub
    End If
     If txtQty.Text = "" Then
     MsgBox "Please Enter Qty", vbExclamation
    txtQty.SetFocus
    Exit Sub
    End If
    
     If txtPurchaseRs.Text = "" Then
     MsgBox "Please Enter Rate#", vbExclamation
    txtPurchaseRs.SetFocus
    Exit Sub
    End If
    
If Val(txtAmount) = 0 Then
    MsgBox "Quantity Cannot be Zero", vbCritical
    Exit Sub
End If
Row = Flx1.Rows - 1
With Flx1

        .Rows = .Rows + 1
                
       
        
        .TextMatrix(Row, 1) = cmbProductName
        .TextMatrix(Row, 2) = txtQty
        .TextMatrix(Row, 3) = txtPurchaseRs
        .TextMatrix(Row, 4) = txtAmount
         
         
        txtAmount = ""
        txtQty = ""
        txtPurchaseRs = ""
       
        cmbProductName.SetFocus
  End With
  CalcTotalAmount
  Else
  End If
End Sub
Public Function GetNewID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetNewID = -1
    
    sSQL = "SELECT Max(tblPurchase.ID)+1 AS ID" & _
            " From tblPurchase"

    
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
    
    sSQL = "SELECT Max(tblPurchasereturn.ID)+1 AS ID" & _
            " From tblPurchasereturn"

    
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

Public Sub ClearTExt()
txtInvoice.Text = ""
cmbProductName.ListIndex = -1
txtTotalAmount.Text = ""
End Sub










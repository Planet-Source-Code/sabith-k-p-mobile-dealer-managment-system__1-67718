VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEReport 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Easy Report"
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
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
         Left            =   4920
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtTotalAmount 
         Height          =   325
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   4560
         Width           =   2055
      End
      Begin MSFlexGridLib.MSFlexGrid flx1 
         Height          =   3495
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6165
         _Version        =   393216
         Cols            =   5
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Show"
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
         Left            =   3840
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTE 
         Height          =   325
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20381697
         CurrentDate     =   39121
      End
      Begin MSComCtl2.DTPicker DTS 
         Height          =   325
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20381697
         CurrentDate     =   39121
      End
      Begin VB.Label Label1 
         Caption         =   "Total Amount"
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
         Left            =   3000
         TabIndex        =   6
         Top             =   4560
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmEReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub View_Date_Purchaserpt()
    On Error GoTo RAE:
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    'Dim PreBill As Long
    
    
  ' PreBill = modRSGroupPreEnt.GetNewBill
    
    'set SQL Expression
    'sSQL = "SELECT * from tblActivation where aDate='" & DT_Start.Value & "'" '.ID, tblActivation.aDate" & _
            " From tblActivation " 'where Date= '" & txtAgentName.DisplayData & "'" & _
            " ORDER BY tblActivation.ActivationID"
            
            
               sSQL = "select * from tblEPurchase WHERE tblEPurchase.eDate Between " & CLng(DTS.Value) & " And " & CLng(DTE.Value) & "" ' " order by DT_Start " & ""
 
            

                    
           
                

             
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog Me.Name, "LoadEntries", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    'If AnyRecordExisted(vRS) = False Then
    '    GoTo RAE
    'End If
    
    'add entries to list
    vRS.MoveFirst
    
    flx1.Rows = vRS.RecordCount + 1
    Row = 0
    While vRS.EOF = False
    
        With flx1
            Row = Row + 1
            '.TextMatrix(Row, 0) = PreBill
            .TextMatrix(Row, 0) = ReadField(vRS.Fields("ID"))
            .TextMatrix(Row, 1) = ReadField(vRS.Fields("eDate"))
            .TextMatrix(Row, 2) = ReadField(vRS.Fields("Amount"))
            .TextMatrix(Row, 3) = ReadField(vRS.Fields("EVal"))
            .TextMatrix(Row, 4) = ReadField(vRS.Fields("FinalAmount"))
            
            End With
           vRS.MoveNext
    Wend
    Exit Sub
RAE:

    Set vRS = Nothing
    MsgBox Err.Description, vbExclamation
    cmdClear_Click
End Sub

Public Sub PurHeading()
flx1.Col = 0
flx1.Row = 0
flx1.ColWidth(0) = 1000
flx1.ColWidth(1) = 1500
flx1.Text = "ID"
flx1.ColAlignment(1) = 2
flx1.Col = 1
flx1.Text = "Date"
flx1.ColAlignment(1) = 1
flx1.Col = 2
flx1.ColWidth(2) = 1500
flx1.Text = "Amount"
flx1.ColAlignment(2) = 2
flx1.Col = 3
flx1.ColWidth(3) = 1500
flx1.Text = "EVal "
flx1.ColAlignment(3) = 2
flx1.Col = 4
flx1.ColWidth(4) = 1500
flx1.Text = "Final Amount "
flx1.ColAlignment(4) = 2
End Sub
Public Sub SaleHeading()
flx1.Col = 0
flx1.Row = 0
flx1.ColWidth(0) = 1000
flx1.ColWidth(1) = 1500
flx1.Text = "ID"
flx1.ColAlignment(1) = 2
flx1.Col = 1
flx1.Text = "Date"
flx1.ColAlignment(1) = 1
flx1.Col = 2
flx1.ColWidth(2) = 2000
flx1.Text = "OutletName"
flx1.ColAlignment(2) = 2
flx1.Col = 3
flx1.ColWidth(3) = 1500
flx1.Text = "Amount "
flx1.ColAlignment(3) = 2
flx1.Col = 4
flx1.ColWidth(4) = 1500
flx1.Text = "Final Amount "
flx1.ColAlignment(4) = 2
End Sub

Private Sub cmdClear_Click()
'ClearFlex flx1
Dim Cap As String
Cap = Me.Caption
Unload Me
Me.Caption = Cap
frmEReport.Show vbModal
txtTotalAmount.Text = ""
End Sub

Private Sub cmdShow_Click()
If Me.Caption = "Easy Purchase Report" Then
PurHeading
View_Date_Purchaserpt
Else
SaleHeading
View_Date_Salerpt
End If
CalcTotalAmount
End Sub
Public Sub View_Date_Salerpt()
    On Error GoTo RAE:
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    'Dim PreBill As Long
    
    
  ' PreBill = modRSGroupPreEnt.GetNewBill
    
    'set SQL Expression
    'sSQL = "SELECT * from tblActivation where aDate='" & DT_Start.Value & "'" '.ID, tblActivation.aDate" & _
            " From tblActivation " 'where Date= '" & txtAgentName.DisplayData & "'" & _
            " ORDER BY tblActivation.ActivationID"
            
            
               sSQL = "select * from tblESale WHERE tblESale.eDate Between " & CLng(DTS.Value) & " And " & CLng(DTE.Value) & "" ' " order by DT_Start " & ""
 
            

                    
           
                

             
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog Me.Name, "LoadEntries", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    'If AnyRecordExisted(vRS) = False Then
    '    GoTo RAE
    'End If
    
    'add entries to list
    vRS.MoveFirst
    
    flx1.Rows = vRS.RecordCount + 1
    Row = 0
    While vRS.EOF = False
    
        With flx1
            Row = Row + 1
            '.TextMatrix(Row, 0) = PreBill
            .TextMatrix(Row, 0) = ReadField(vRS.Fields("ID"))
            .TextMatrix(Row, 1) = ReadField(vRS.Fields("eDate"))
            .TextMatrix(Row, 2) = ReadField(vRS.Fields("OutletName"))
            .TextMatrix(Row, 3) = ReadField(vRS.Fields("Amount"))
            .TextMatrix(Row, 4) = ReadField(vRS.Fields("FinalAmount"))
            
            End With
           vRS.MoveNext
    Wend
    Exit Sub
RAE:

    Set vRS = Nothing
    MsgBox Err.Description, vbExclamation
    cmdClear_Click
End Sub

Private Sub Form_Load()
DTS.Value = Date
DTE.Value = Date
End Sub
Public Function CalcTotalAmount()
 Amt = 0
         For i = 1 To flx1.Rows - 1
         Amt = Amt + Val(flx1.TextMatrix(i, 4))
         Next i
         
         txtTotalAmount = Amt
End Function
Public Function CalcPurTotalAmount()
 Amt = 0
         For i = 1 To flx1.Rows - 1
         Amt = Amt + Val(flx1.TextMatrix(i, 5))
         Next i
         
         txtTotalAmount = Amt
End Function


VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReport 
   Caption         =   "Form1"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   10935
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Select Month"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   5775
      Begin VB.CommandButton cmdExport 
         Caption         =   "Export to Excel"
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
         Left            =   3840
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "Show"
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
         Left            =   2640
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPReport 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50528257
         CurrentDate     =   39098
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flx1 
      Height          =   6255
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   11033
      _Version        =   393216
      Cols            =   40
      FixedCols       =   0
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PrdName As String
'Dim OPStock As String
Dim CloseStock As Long
Dim PurchaseRate As Long
Dim Product As String
Dim TotalMonthSale As Long
Dim TotalMonthPurchase As Long
Dim TotalMonthQT As Currency
Dim StBalance As Long
Dim objExcel As Excel.Application
Dim objworkbook As Excel.Workbook
Public Sub Heading()
Dim a  As Integer
a = 3
flx1.Col = 0
flx1.Row = 0
flx1.ColWidth(0) = 2000
flx1.ColWidth(1) = 1500
flx1.Text = "Item Name"
flx1.ColAlignment(1) = 2
flx1.Col = 1
flx1.Text = "Opening Stock"
flx1.ColAlignment(1) = 1
flx1.Col = 2
flx1.ColWidth(2) = 1500
flx1.Text = "Opening Value"
flx1.ColAlignment(2) = 2
flx1.Col = 3
flx1.ColWidth(3) = 1500
flx1.Text = "Primary Stock"
flx1.ColAlignment(3) = 2
flx1.Col = 4
flx1.ColWidth(4) = 1500
flx1.Text = "Primary Value "
flx1.ColAlignment(4) = 2
flx1.Col = 5
flx1.ColWidth(5) = 1500
flx1.Text = "Closing Stock"
flx1.ColAlignment(5) = 2
flx1.Col = 6
flx1.ColWidth(6) = 1500
flx1.Text = "Closing Value"
flx1.ColAlignment(6) = 2
flx1.Col = 7
flx1.ColWidth(7) = 1500
flx1.Text = "MTD Sec"
flx1.ColAlignment(7) = 2
flx1.Col = 8
flx1.ColWidth(8) = 1500
flx1.Text = "MTD Sec Value"
flx1.ColAlignment(8) = 2
flx1.Col = 9
flx1.ColWidth(9) = 1000
flx1.Text = "1st"
flx1.ColAlignment(9) = 2
flx1.Col = 10
flx1.ColWidth(10) = 1000
flx1.Text = "2nd"
flx1.ColAlignment(10) = 2
flx1.Col = 11
flx1.ColWidth(11) = 1000
flx1.Text = "3rd"
flx1.ColAlignment(11) = 2

For i = 12 To Daysinmonth(DTPReport) + 8
flx1.Col = i
a = a + 1
flx1.ColWidth(i) = 1000
flx1.Text = a & "th"
Next i
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdExport_Click()
Dim i As Long
Dim n As Long
On Error Resume Next
Set objExcel = GetObject(, "Excel.Application")
If Err.Number Then
   Err.Clear
   Set objExcel = CreateObject("Excel.Application")
   If Err.Number Then
      MsgBox "Can't open Excel."
   End If
End If
objExcel.Visible = True
Set objworkbook = objExcel.Workbooks.Add
AppActivate "Activation Detilas"

For i = 0 To 3
    flx1.Row = i
    For n = 0 To 40
        flx1.Col = n
        objworkbook.ActiveSheet.Cells(i + 1, n + 1).Value = flx1.Text
    Next
Next

End Sub

Private Sub cmdShow_Click()
If DTPReport.Day <> 1 Then
MsgBox "Please Enter the First Date of the Month", vbExclamation
Exit Sub
End If
LoadFlxValue
End Sub

Private Sub Form_Load()
Heading
End Sub
Public Function LoadFlxValue()
    On Error GoTo RAE:
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim sSQL1 As String
    Dim openingStock As String
    Dim ProductName As String
    Dim PrimaryValue As Long
    Dim CloseValue As Long
    Dim i As Integer
    
   sSQL = "select * from tblProduct"

    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog Me.Name, "LoadEntries", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
        GoTo RAE
    End If
      
    PrNo = vRS.RecordCount
    vRS.MoveFirst

For i = 1 To PrNo

    flx1.Rows = PrNo + 1
    Row = i
    ProductName = vRS.Fields("Name")   ' Get Product Name
    ProductValue = vRS.Fields("PrValue")
    ProductAmount = vRS.Fields("Amount")
    
    '   Load Its Opening Stock of Product Name
   
      OpeningStcok ProductName, DTPReport.Value, flx1, i
        

    
       '   Load this months sale of Product Name by day
      
      ThisMonthSale ProductName, DTPReport.Value, flx1, i
         
         
    '///   *********************************************************************
    '       Values to the Grid
    '/////// ****************************************************************
         
         
    OpenValue = StBalance * ProductValue
    MonthValue = TotalMonthSale * ProductAmount
    CloseStock = StBalance + TotalMonthPurchase - TotalMonthSale
    CloseValue = CloseStock * ProductValue
    PrimaryValue = TotalMonthPurchase * ProductValue
    
    flx1.TextMatrix(Row, 0) = ProductName
    flx1.TextMatrix(Row, 1) = CStr(StBalance)
    flx1.TextMatrix(Row, 2) = CStr(OpenValue)
    flx1.TextMatrix(Row, 3) = CStr(TotalMonthPurchase)
    flx1.TextMatrix(Row, 4) = CStr(PrimaryValue)
    flx1.TextMatrix(Row, 5) = CStr(CloseStock)
    flx1.TextMatrix(Row, 6) = CStr(CloseValue)
    flx1.TextMatrix(Row, 7) = CStr(TotalMonthSale)
    flx1.TextMatrix(Row, 8) = CStr(MonthValue)
      
    ' ******************************************************************************
    '*******************************************************************************
  
  vRS.MoveNext

Next i

    Exit Function
RAE:

    Set vRS = Nothing
    MsgBox Err.Description, vbExclamation
End Function

Public Sub OpeningStcok(ProName As String, PurDate As Date, Flx As MSFlexGrid, PRow As Integer)
'On Error Resume Next
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim PurQty As Long
    Dim SaleQty As Long
    
    
    sSQL = "Select * from tblPurchase where pdate < " & CLng(PurDate) & " And ProductName = '" & ProName & "'"
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
       'WriteErrorLog Me.Name, "LoadEntries", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
       ' GoTo RAE
    End If

    

If vRS.EOF = False And vRS.BOF = False Then
   
   RecNo = vRS.RecordCount
   vRS.MoveFirst
   For i = 1 To RecNo
   X = vRS.Fields("Qty")
   PurQty = PurQty + X
   vRS.MoveNext
   Next i

End If

    sSQL = "Select * from tblsale where sdate < " & CLng(PurDate) & " And ProductName = '" & ProName & "'"
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
       'WriteErrorLog Me.Name, "LoadEntries", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
       ' GoTo RAE
    End If

If vRS.EOF = False And vRS.BOF = False Then
   
   RecNo = vRS.RecordCount
   vRS.MoveFirst
   For i = 1 To RecNo
   X = vRS.Fields("Qty")
   SaleQty = SaleQty + X
   vRS.MoveNext
   Next i
End If

StBalance = PurQty - SaleQty
'Flx.TextMatrix(PRow, 1) = StBalance

End Sub



Public Sub ThisMonthSale(ProName As String, PurDate As Date, Flx As MSFlexGrid, PRow As Integer)
'On Error Resume Next
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim PurQty As Long
    Dim SaleQty As Long
    Dim QrDay As Date
   
    
    DayinMonth = Daysinmonth(PurDate)  ' Number of days in this month
    
    
    ' *************************************************************************
    '/////**********************   This Months Total Sale ( Day by Day)
    
    TotalMonthSale = 0
    For D = 1 To DayinMonth
    
    DaySaleQty = 0
    
    QrDay = D & "/" & DTPReport.Month & "/" & DTPReport.Year
    
    sSQL = "Select * from tblSale where sdate = " & CLng(QrDay) & " And ProductName = '" & ProName & "'"
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
       'WriteErrorLog Me.Name, "LoadEntries", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
       ' GoTo RAE
    End If
     
 
  
    If vRS.EOF = False And vRS.BOF = False Then
    
        RecNo = vRS.RecordCount
        vRS.MoveFirst
        X = 0
    
        For i = 1 To RecNo
    
        X = vRS.Fields("Qty")
        DaySaleQty = DaySaleQty + X     ' Days Sale
        vRS.MoveNext
    
        Next i
    
    End If
    
    Flx.TextMatrix(PRow, 8 + D) = CStr(DaySaleQty)    ''''' Days sale to Grid
    
    TotalMonthSale = TotalMonthSale + DaySaleQty ' This  Month's Total Sale

    Next D
       
'  **********************   \\\\\\\\\\\\\  This Months Total Sale ( Day by Day)
' *************************************************************************************





'\\\\\\\\\\\\\\\\\********************** This Months Total Purchase ( Day by Day)
'*******************************************************************************

    TotalMonthPurchase = 0
    For D = 1 To DayinMonth
    
    DayPurchaseQty = 0
    
    QrDay = D & "/" & DTPReport.Month & "/" & DTPReport.Year
    
    sSQL = "Select * from tblPurchase where pdate = " & CLng(QrDay) & " And ProductName = '" & ProName & "'"
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
       'WriteErrorLog Me.Name, "LoadEntries", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
       ' GoTo RAE
    End If
      
    If vRS.EOF = False And vRS.BOF = False Then
    
    RecNo = vRS.RecordCount
    vRS.MoveFirst
    X = 0
    
    For i = 1 To RecNo
    
    X = vRS.Fields("Qty")
    DayPurchaseQty = DayPurchaseQty + X
    vRS.MoveNext
    
    Next i
    
    
End If
    
    'Flx.TextMatrix(PRow, 8 + D) = CStr(DaySaleQty)
    
    TotalMonthPurchase = TotalMonthPurchase + DayPurchaseQty

    Next D

' ///  ********************************************************************************
' ******************************************************************************************


End Sub

Public Function Daysinmonth(mydate)

    Dim Nextmonth, EOFmonth
    Nextmonth = DateAdd("m", 1, mydate)
    EOFmonth = Nextmonth - DatePart("d", Nextmonth)
    Daysinmonth = DatePart("d", EOFmonth)
    
End Function

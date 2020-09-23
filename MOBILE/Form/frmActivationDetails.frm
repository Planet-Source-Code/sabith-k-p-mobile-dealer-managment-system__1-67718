VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmActivationDetails 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activation Details"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   10365
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Change the mobile No from the Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   3495
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   120
         Top             =   360
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtmobNo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "9895832009"
         Top             =   360
         Width           =   2175
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flx1 
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7858
      _Version        =   393216
      Cols            =   14
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Select Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4200
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton cmdPrintDate 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Print"
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
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdShow 
         BackColor       =   &H00C0C0C0&
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
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optnotCompleted 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Not Completed"
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
         Left            =   2040
         TabIndex        =   11
         Top             =   1680
         Width           =   1695
      End
      Begin VB.OptionButton optCompleted 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Completed"
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
         Left            =   2040
         TabIndex        =   10
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton optAll 
         BackColor       =   &H00C0C0C0&
         Caption         =   "All"
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
         Left            =   2040
         TabIndex        =   9
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdReport 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Print"
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
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Find"
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
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DT_Start 
         Height          =   330
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   20250625
         CurrentDate     =   38977
      End
      Begin MSComCtl2.DTPicker DT_End 
         Height          =   330
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   20250625
         CurrentDate     =   38977
      End
      Begin MSComCtl2.DTPicker DTDACtivation 
         Height          =   330
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   20250625
         CurrentDate     =   38977
      End
      Begin MSComCtl2.DTPicker DTActivation 
         Height          =   330
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   20250625
         CurrentDate     =   38977
      End
      Begin VB.Line Line1 
         X1              =   600
         X2              =   6000
         Y1              =   1080
         Y2              =   1080
      End
   End
End
Attribute VB_Name = "frmActivationDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objExcel As Excel.Application
Dim objworkbook As Excel.Workbook
Private Sub cmdConvExcel_Click()
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
    For n = 0 To 9
        flx1.Col = n
        objworkbook.ActiveSheet.Cells(i + 1, n + 1).Value = flx1.Text
        
    Next
Next

End Sub

Private Sub cmdFind_Click()

If optAll.Value = False And optCompleted.Value = False And optnotCompleted.Value = False Then
MsgBox "Please Select the option", vbExclamation
Exit Sub
End If

ClearFlex flx1
LoadAllValue
End Sub

Private Sub cmdPrintDate_Click()
ShowReport "tblActivation", DataReport2, " WHERE tblActivation.CurDate Between " & CLng(DTDACtivation.Value) & " And " & CLng(DTActivation.Value) & "" ' " order by dtActivation.value " & """

RptMobileNO = txtmobNo.Text
DataReport2.Sections("Section4").Controls("lblMobil").Caption = RptMobileNO

DataReport2.Show vbModal



End Sub

Private Sub cmdReport_Click()
If optAll.Value = False And optCompleted.Value = False And optnotCompleted.Value = False Then
MsgBox "Please Select the option", vbExclamation
Exit Sub
End If
If optAll.Value = True Then
ShowReport "tblActivation", DataReport1, " WHERE tblActivation.aDate Between " & CLng(DT_Start.Value) & " And " & CLng(DT_End.Value) & "" ' " order by DT_Start " & """
ElseIf optCompleted.Value = True Then
ShowReport "tblActivation", DataReport1, " WHERE tblActivation.aDate Between " & CLng(DT_Start.Value) & " And " & CLng(DT_End.Value) & "and Complete=" & 1 & ""  ' " order by DT_Start " & ""

ElseIf optnotCompleted.Value = True Then
ShowReport "tblActivation", DataReport1, " WHERE tblActivation.aDate Between " & CLng(DT_Start.Value) & " And " & CLng(DT_End.Value) & "and Complete=" & 0 & ""  ' " order by DT_Start " & ""

End If

RptMobileNO = txtmobNo.Text
DataReport1.Sections("Section4").Controls("lblMobil").Caption = RptMobileNO
DataReport1.Show
'Unload Me

End Sub

Private Sub cmdShow_Click()
ClearFlex flx1

View_Date_ActivationDetails
End Sub

Private Sub LoadEntries()
   'On Error GoTo RAE
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim rec_count, X
    Dim l_results As ListItem
    
    lvw_recordsfound.ListItems.Clear

    'set SQL Expression
    'sSQL = "Select * from tblActivation" 'WHERE aDate=" & DT_Start
    sSQL = "select * from tblActivation WHERE (((tblActivation.aDate) Between # " + DT_Start.Value + " # And # " + DT_End.Value + " #))"

      If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog Me.Name, "LoadEntries", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
   With vRS
     .MoveLast
        rec_count = .RecordCount
        
    .MoveFirst
        For X = 1 To rec_count
    
    Set l_results = lvw_recordsfound.ListItems.Add(X, , !ID)
                    l_results.SubItems(1) = !MobileNo
                    l_results.SubItems(2) = !aDate
                    
                   
    .MoveNext
        Next
End With
RAE:
    Set vRS = Nothing
      lvw_recordsfound.Refresh
      
End Sub



Private Sub Command2_Click()

End Sub

Private Sub Command1_Click()


End Sub

Private Sub Form_Load()
DT_Start.Value = Date
DT_End.Value = Date
DTDACtivation.Value = Date
DTActivation.Value = Date
Heading
End Sub
Private Sub LoadAllValue()
    On Error GoTo RAE:
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    'Dim PreBill As Long
    
    
  ' PreBill = modRSGroupPreEnt.GetNewBill
    
    'set SQL Expression
    'sSQL = "SELECT * from tblActivation where aDate='" & DT_Start.Value & "'" '.ID, tblActivation.aDate" & _
            " From tblActivation " 'where Date= '" & txtAgentName.DisplayData & "'" & _
            " ORDER BY tblActivation.ActivationID"
            If optAll.Value = True Then
                    sSQL = "select * from tblActivation WHERE tblActivation.aDate Between " & CLng(DT_Start.Value) & " And " & CLng(DT_End.Value) & "" ' " order by DT_Start " & ""

            ElseIf optCompleted = True Then

                    sSQL = "select * from tblActivation WHERE tblActivation.aDate Between " & CLng(DT_Start.Value) & " And " & CLng(DT_End.Value) & "and Complete=" & 1 & ""  ' " order by DT_Start " & ""

            ElseIf optnotCompleted.Value = True Then

                    sSQL = "select * from tblActivation WHERE tblActivation.aDate Between " & CLng(DT_Start.Value) & " And " & CLng(DT_End.Value) & "and Complete=" & 0 & ""  ' " order by DT_Start " & ""

            End If
                

             
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
            .TextMatrix(Row, 1) = ReadField(vRS.Fields("adate"))
            .TextMatrix(Row, 2) = ReadField(vRS.Fields("MobileNo"))
            .TextMatrix(Row, 3) = "WAJID COMMUNICAION"
            .TextMatrix(Row, 4) = ReadField(vRS.Fields("Attestphoto"))
            .TextMatrix(Row, 5) = ReadField(vRS.Fields("photograph"))
            .TextMatrix(Row, 6) = ReadField(vRS.Fields("photoid"))
            .TextMatrix(Row, 7) = ReadField(vRS.Fields("CSinAPEF"))
            .TextMatrix(Row, 8) = ReadField(vRS.Fields("Retailseal"))
            .TextMatrix(Row, 9) = ReadField(vRS.Fields("DistributerSeal"))
            .TextMatrix(Row, 10) = ReadField(vRS.Fields("Outletname"))
           .TextMatrix(Row, 11) = ReadField(vRS.Fields("CustomerName"))
            .TextMatrix(Row, 12) = ReadField(vRS.Fields("Address"))
            End With
           vRS.MoveNext
    Wend
    Exit Sub
RAE:

    Set vRS = Nothing
    MsgBox Err.Description, vbExclamation
    Unload Me
    frmActivationDetails.Show
End Sub
Public Sub Heading()
flx1.Col = 0
flx1.Row = 0
flx1.ColWidth(0) = 1000
flx1.ColWidth(1) = 1200
flx1.Text = "ID"
flx1.ColAlignment(1) = 2
flx1.Col = 1
flx1.Text = "Date"
flx1.ColAlignment(1) = 1
flx1.Col = 2
flx1.ColWidth(2) = 2000
flx1.Text = "Mobile No"
flx1.ColAlignment(2) = 2
flx1.Col = 3
flx1.ColWidth(3) = 1500
flx1.Text = "Distributor Name "
flx1.ColAlignment(3) = 2
flx1.Col = 4
flx1.ColWidth(4) = 1500
flx1.Text = "AttestPhoto "
flx1.ColAlignment(4) = 2
flx1.Col = 5
flx1.ColWidth(5) = 1500
flx1.Text = "Photograph"
flx1.ColAlignment(5) = 2
flx1.Col = 6
flx1.ColWidth(5) = 1500
flx1.Text = "PhotoID"
flx1.ColAlignment(5) = 2
flx1.Col = 7
flx1.ColWidth(7) = 1600
flx1.Text = "CustomerSigninAPEF"
flx1.ColAlignment(7) = 2
flx1.Col = 8
flx1.ColWidth(8) = 1500
flx1.Text = "Retail Seal"
flx1.ColAlignment(8) = 2
flx1.Col = 9
flx1.ColWidth(9) = 1500
flx1.Text = "Distributer Seal"
flx1.ColAlignment(9) = 2
flx1.Col = 10
flx1.ColWidth(10) = 1500
flx1.Text = "OutletName"
flx1.ColAlignment(10) = 2
flx1.Col = 11
flx1.ColWidth(11) = 2500
flx1.Text = "Coustomer Name"
flx1.ColAlignment(11) = 2
flx1.ColWidth(11) = 2500
flx1.Col = 12
flx1.ColWidth(12) = 2500
flx1.Text = "Address"
flx1.ColAlignment(12) = 2
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Timer1_Timer()
If Check1.Value = 1 Then
txtmobNo.Locked = False
Else
txtmobNo.Locked = True

End If
End Sub
Public Sub View_Date_ActivationDetails()
    On Error GoTo RAE:
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    'Dim PreBill As Long
    
    
  ' PreBill = modRSGroupPreEnt.GetNewBill
    
    'set SQL Expression
    'sSQL = "SELECT * from tblActivation where aDate='" & DT_Start.Value & "'" '.ID, tblActivation.aDate" & _
            " From tblActivation " 'where Date= '" & txtAgentName.DisplayData & "'" & _
            " ORDER BY tblActivation.ActivationID"
            
    sSQL = "select * from tblActivation WHERE tblActivation.CurDate Between " & CLng(DTDACtivation.Value) & " And " & CLng(DTActivation.Value) & "" ' " order by DT_Start " & ""

            

                    
           
                

             
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
            .TextMatrix(Row, 1) = ReadField(vRS.Fields("Curdate"))
            .TextMatrix(Row, 2) = ReadField(vRS.Fields("MobileNo"))
            .TextMatrix(Row, 3) = "WAJID COMMUNICAION"
            .TextMatrix(Row, 4) = ReadField(vRS.Fields("Attestphoto"))
            .TextMatrix(Row, 5) = ReadField(vRS.Fields("photograph"))
            .TextMatrix(Row, 6) = ReadField(vRS.Fields("photoid"))
            .TextMatrix(Row, 7) = ReadField(vRS.Fields("CSinAPEF"))
            .TextMatrix(Row, 8) = ReadField(vRS.Fields("Retailseal"))
            .TextMatrix(Row, 9) = ReadField(vRS.Fields("DistributerSeal"))
            .TextMatrix(Row, 10) = ReadField(vRS.Fields("Outletname"))
           .TextMatrix(Row, 11) = ReadField(vRS.Fields("CustomerName"))
            .TextMatrix(Row, 12) = ReadField(vRS.Fields("Address"))
            End With
           vRS.MoveNext
    Wend
    Exit Sub
RAE:

    Set vRS = Nothing
    MsgBox Err.Description, vbExclamation
    Unload Me
    frmActivationDetails.Show
End Sub



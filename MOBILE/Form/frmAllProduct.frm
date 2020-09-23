VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAllProduct 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "All Product"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   6480
      TabIndex        =   5
      Top             =   0
      Width           =   3495
      Begin VB.TextBox txtSearch 
         Height          =   325
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
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
         Left            =   2400
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Close"
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
      Left            =   9120
      TabIndex        =   6
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      Left            =   8160
      TabIndex        =   4
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "&Modify"
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
      Left            =   7080
      TabIndex        =   3
      Top             =   5400
      Width           =   975
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
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   5400
      Width           =   975
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllProduct.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListProduct 
      Height          =   4455
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "i16x16"
      SmallIcons      =   "i16x16"
      ColHdrIcons     =   "i16x16"
      ForeColor       =   5053698
      BackColor       =   14145495
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Product Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Category"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Purchase Rate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "SaleRate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmAllProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mShowForm As Boolean
Public Function ShowForm() As Boolean
    'show form
    Me.Show vbModal
    
    ShowForm = mShowForm
    
End Function
Private Sub LoadEntries()
    On Error GoTo RAE
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim rec_count, X
    Dim l_results As ListItem
    
    ListProduct.ListItems.Clear

    'set SQL Expression
    sSQL = "SELECT * From tblProduct" & _
            " ORDER BY tblProduct.Id"
             
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog Me.Name, "LoadEntries", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
   With vRS
     .MoveLast
        rec_count = .RecordCount
        
    .MoveFirst
        For X = 1 To rec_count
    
    Set l_results = ListProduct.ListItems.Add(X, , !ID, 1, 1)
                    l_results.SubItems(1) = !Name
                    l_results.SubItems(2) = !Category
                    l_results.SubItems(3) = !prValue
                    l_results.SubItems(4) = !Amount
                    l_results.SubItems(5) = !oDate
                    
                    
                   
    .MoveNext
        Next
End With
RAE:
    Set vRS = Nothing
      ListProduct.Refresh
End Sub

Private Sub cmdAdd_Click()
If frm_aed_Product.ShowAdd = True Then
    LoadEntries
End If
End Sub

Private Sub cmdDelete_Click()
If ListProduct.ListItems.Count > 0 Then
            If MsgBox("Are you sure you want to delete Product '" & ListProduct.SelectedItem.Text & "'?", vbQuestion + vbYesNo) = vbYes Then
                If DeleteProduct(ListProduct.SelectedItem.Text) = True Then
                    LoadEntries
                Else
                    MsgBox Me.Name & "cmdDelete_Click" & "Faild:" & DeleteProduct(ListProduct.SelectedItem.Text) = True
                End If
            End If
    End If
End Sub

Private Sub cmdEdit_Click()
 
    End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdNew_Click()
End Sub



Private Sub cmdModify_Click()
If ListProduct.ListItems.Count > 0 Then
        If frm_aed_Product.ShowEdit(ListProduct.SelectedItem.Text) = True Then
            LoadEntries
        End If
    End If
End Sub

Private Sub cmdSearch_Click()
If ListProduct.ListItems.Count < 1 Then Exit Sub
    Call search_in_listview(ListProduct, txtSearch.Text)
End Sub

Private Sub Form_Activate()
 LoadEntries
End Sub

Private Sub listProduct_DblClick()
    cmdModify_Click
End Sub


Private Sub txtSearch_Change()
cmdSearch_Click
End Sub









VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAllCatogory 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All Catogory"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   2640
      TabIndex        =   4
      Top             =   0
      Width           =   3495
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
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtSearch 
         Height          =   325
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2175
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
      Left            =   5280
      TabIndex        =   5
      Top             =   4320
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
      Left            =   4320
      TabIndex        =   3
      Top             =   4320
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
      Left            =   3240
      TabIndex        =   2
      Top             =   4320
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
      Left            =   2160
      TabIndex        =   1
      Top             =   4320
      Width           =   975
   End
   Begin MSComctlLib.ListView ListCategory 
      Height          =   3375
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5953
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   8819
      EndProperty
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
            Picture         =   "frmAllCatogory.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAllCatogory"
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
    
    ListCategory.ListItems.Clear

    'set SQL Expression
    sSQL = "SELECT * From tblCategory" & _
            " ORDER BY tblCategory.Id"
             
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog Me.Name, "LoadEntries", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
   With vRS
     .MoveLast
        rec_count = .RecordCount
        
    .MoveFirst
        For X = 1 To rec_count
    
    Set l_results = ListCategory.ListItems.Add(X, , !ID, 1, 1)
                    l_results.SubItems(1) = !Name
                    
                    
                   
    .MoveNext
        Next
End With
RAE:
    Set vRS = Nothing
      ListCategory.Refresh
End Sub

Private Sub cmdAdd_Click()
 If frm_aed_Category.ShowAdd = True Then
    LoadEntries
End If

End Sub

Private Sub cmdDelete_Click()
If ListCategory.ListItems.Count > 0 Then
            If MsgBox("Are you sure you want to delete Category '" & ListCategory.SelectedItem.Text & "'?", vbQuestion + vbYesNo) = vbYes Then
                If DeleteCategory(ListCategory.SelectedItem.Text) = True Then
                    LoadEntries
                Else
                    MsgBox Me.Name & "cmdDelete_Click" & "Faild:" & DeleteCategory(ListCategory.SelectedItem.Text) = True
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

Private Sub cmdRefresh_Click()
    
End Sub

Private Sub cmdModify_Click()
If ListCategory.ListItems.Count > 0 Then
        If frm_aed_Category.ShowEdit(ListCategory.SelectedItem.Text) = True Then
            LoadEntries
        End If
    End If
End Sub

Private Sub cmdSearch_Click()
If ListCategory.ListItems.Count < 1 Then Exit Sub
    Call search_in_listview(ListCategory, txtSearch.Text)
End Sub

Private Sub Form_Activate()
 LoadEntries
End Sub

Private Sub listCategory_DblClick()
    cmdModify_Click
End Sub

Private Sub txtSearch_Change()
cmdSearch_Click
End Sub













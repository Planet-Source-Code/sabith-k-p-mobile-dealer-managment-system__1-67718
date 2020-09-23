VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRestore 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Restore"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   548
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   3480
      TabIndex        =   17
      Top             =   5100
      Width           =   1425
   End
   Begin VB.PictureBox bgBF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   180
      ScaleHeight     =   1185
      ScaleWidth      =   7785
      TabIndex        =   7
      Top             =   930
      Visible         =   0   'False
      Width           =   7785
      Begin VB.TextBox txtBF 
         Height          =   345
         Left            =   1050
         TabIndex        =   11
         Top             =   690
         Width           =   6285
      End
      Begin VB.CommandButton cmdGetBF 
         Caption         =   "..."
         Height          =   345
         Left            =   7350
         TabIndex        =   10
         Top             =   690
         Width           =   405
      End
      Begin VB.Label lblBFCaption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Backup File to restore:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1050
         TabIndex        =   8
         Top             =   0
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00EDEBE9&
         Height          =   2040
         Left            =   -60
         TabIndex        =   9
         Top             =   -480
         Width           =   975
      End
   End
   Begin VB.PictureBox bgHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EDEBE9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   548
      TabIndex        =   0
      Top             =   0
      Width           =   8220
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to Database Restore "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   330
         TabIndex        =   1
         Top             =   270
         Width           =   3780
      End
   End
   Begin VB.PictureBox bgStep 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4635
      Index           =   0
      Left            =   60
      ScaleHeight     =   309
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   547
      TabIndex        =   2
      Top             =   870
      Width           =   8205
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   345
         Left            =   5070
         TabIndex        =   6
         Top             =   4230
         Width           =   1425
      End
      Begin VB.CommandButton cmdNext1 
         Caption         =   "&Next"
         Enabled         =   0   'False
         Height          =   345
         Left            =   6600
         TabIndex        =   5
         Top             =   4230
         Width           =   1425
      End
      Begin MSComctlLib.ListView listBF 
         Height          =   3735
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   6350
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Size"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date Modified"
            Object.Width           =   4233
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
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
               Picture         =   "frmRestore.frx":0000
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 1:    Select Database Backup File to Restore"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   4755
      End
   End
   Begin VB.PictureBox bgStep 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3165
      Index           =   1
      Left            =   30
      ScaleHeight     =   211
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   547
      TabIndex        =   12
      Top             =   2310
      Visible         =   0   'False
      Width           =   8205
      Begin MSComctlLib.ProgressBar progCopy 
         Height          =   315
         Left            =   1140
         TabIndex        =   15
         Top             =   1230
         Visible         =   0   'False
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "&Start"
         Height          =   345
         Left            =   6570
         TabIndex        =   13
         Top             =   2790
         Width           =   1425
      End
      Begin VB.Label lblStartMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click 'Start' button to continue restoring file."
         Height          =   195
         Left            =   1260
         TabIndex        =   16
         Top             =   1290
         Width           =   3165
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 2:    Extract Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   60
         TabIndex        =   14
         Top             =   60
         Width           =   2175
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F5F5F5&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   75
      Left            =   -30
      Top             =   780
      Width           =   12000
   End
End
Attribute VB_Name = "frmRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim BFPath As String

Private WithEvents clsHuf As clsHuffman
Attribute clsHuf.VB_VarHelpID = -1

Public Sub ShowForm()
    
    'check current user
    If LCase(CurrentUser.UserID) <> "administrator" Then
        MsgBox "You are not permitted to access Database Restore.", vbExclamation
        Unload Me
        Exit Sub
    End If

    'load BF files
    RefreshBF
        
    Me.Show vbModal
End Sub


Private Sub clsHuf_Progress(Procent As Integer)
    progCopy.Value = Procent
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    
    Dim FSO As New FileSystemObject
    Dim li As ListItem
    
    If listBF.ListItems.Count < 1 Then
        MsgBox "There is no Backup File to delete.", vbExclamation
        Unload Me
        GoTo ReleaseAndExit
    End If
    
    If MsgBox("Deleting file cannot be undo. Do you want continue deleting Backup File/s", vbQuestion + vbOKCancel + vbDefaultButton2) = vbOK Then
        On Error GoTo errh:
        For Each li In listBF.ListItems
            If li.Selected = True Then
                FSO.DeleteFile BFPath & "\" & li.Text, True
            End If
        Next
    End If
    
    RefreshBF
    
    
ReleaseAndExit:
    Set FSO = Nothing
    Set li = Nothing
    Exit Sub
errh:
    MsgBox Me.Name & "cmdDelete_Click" & Err.Description
    Resume ReleaseAndExit
End Sub

Private Sub cmdGetBF_Click()
    
    bgBF.Visible = False
    bgStep(1).Visible = False
    
    bgStep(0).Visible = True
    
    
End Sub

Private Sub cmdNext1_Click()
    
    If listBF.ListItems.Count < 1 Then
        MsgBox "There is no Backup File to restore.", vbExclamation
        Unload Me
        GoTo ReleaseAndExit
    End If
    
    txtBF.Text = BFPath & "\" & listBF.SelectedItem.Text
    lblBFCaption = "Backup File to restore: " & listBF.SelectedItem.Text & vbNewLine & _
                    "Size: " & listBF.SelectedItem.SubItems(1) & vbNewLine & _
                    "Last Modified: " & listBF.SelectedItem.SubItems(2)
                    
    bgStep(0).Visible = False
    bgBF.Visible = True
    bgStep(1).Visible = True
    
ReleaseAndExit:
    
End Sub

Private Sub cmdStart_Click()
    
    Dim stmpFP As String
    Dim FSO As New FileSystemObject
    
    If listBF.ListItems.Count < 1 Then
        MsgBox "There is no Backup File to restore.", vbExclamation
        GoTo ReleaseAndExit
    End If
    
    If Not FSO.FileExists(txtBF.Text) Then
        MsgBox "Please enter valid backup file name", vbExclamation
        GoTo ReleaseAndExit
    End If
    
    cmdCancel.Enabled = False
    lblStartMsg.Visible = False
    progCopy.Visible = True
    cmdStart.Enabled = False
    
    Set clsHuf = New clsHuffman
    
    stmpFP = FSO.GetSpecialFolder(2).Name & "DIMS1tmpCF.tmp"
    
    'On Error GoTo errh
    
    If FSO.FileExists(stmpFP) = True Then
        FSO.DeleteFile stmpFP
    End If
    
    clsHuf.DecodeFile txtBF.Text, stmpFP
    
    progCopy.Visible = False
    lblStartMsg.Caption = "Copying File, please wait..."
    lblStartMsg.Visible = True
    DoEvents
    
    
    If MsgBox("Do you want to overwrite the current Database file wth this newly restored file?", vbQuestion + vbOKCancel) = vbOK Then

        'close databse
        ModConn.PrimeDB.Close
        
        
        'delete current db file
        If FSO.FileExists(ModConn.DBPathFileName) = True Then
            FSO.DeleteFile ModConn.DBPathFileName
        End If
        
        'move database file
        FSO.MoveFile stmpFP, ModConn.DBPathFileName
        
        'reconnect database file
        If ModConn.OpenDB = True Then
            
            'success
            MsgBox "The backup file was successfully restored and it is now ready to use.", vbInformation
        End If
        
        'close this form
        Unload Me
    Else
        lblStartMsg.Caption = "Click 'Start' button to continue restoring file."
        lblStartMsg.Visible = True
        progCopy.Visible = False
    End If
    
ReleaseAndExit:
    Set FSO = Nothing
    progCopy.Visible = False
    progCopy.Visible = False
    cmdCancel.Enabled = True
    cmdStart.Enabled = True
    Exit Sub
errh:
   MsgBox Me.Name & "cmdStart_Click" & Err.Description
    Resume ReleaseAndExit
End Sub



Private Sub RefreshBF()
    
    Dim FSO As New FileSystemObject
    Dim FFolder As Folder
    Dim FFile As File
    
    
    
    
    listBF.ListItems.Clear
    
    'set path
    BFPath = App.Path & "\Backup"
    
    'check path
    If FSO.FolderExists(BFPath) = False Then
        MsgBox "There is no Backup File to restore.", vbExclamation
        Unload Me
        GoTo ReleaseAndExit
    End If
    
    Set FFolder = FSO.GetFolder(BFPath)
    
    For Each FFile In FFolder.Files
        listBF.ListItems.Add , , FFile.Name, 1, 1
        With listBF.ListItems.Item(listBF.ListItems.Count)
            .SubItems(1) = FormatNumber(FFile.Size / 1048576, 1) & " MB"
            .SubItems(2) = Format$(FFile.DateLastModified, "MMM - dd - yyyy")
        End With
    Next
    
    
    If listBF.ListItems.Count > 0 Then
        cmdNext1.Enabled = True
        cmdDelete.Enabled = True
    Else
        MsgBox "There is no Backup File to restore.", vbExclamation
        Unload Me
        GoTo ReleaseAndExit
    End If
    
ReleaseAndExit:
    Set FFile = Nothing
    Set FSO = Nothing
    Set FFolder = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set clsHuf = Nothing
End Sub





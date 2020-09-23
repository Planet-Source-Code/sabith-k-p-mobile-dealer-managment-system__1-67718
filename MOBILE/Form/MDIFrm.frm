VERSION 5.00
Begin VB.MDIForm MDIFrm 
   BackColor       =   &H8000000C&
   Caption         =   "Mobile"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "MDIFrm.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   10200
      Left            =   0
      Picture         =   "MDIFrm.frx":624A
      ScaleHeight     =   10140
      ScaleWidth      =   15180
      TabIndex        =   8
      Top             =   615
      Width           =   15240
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   15180
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   3600
         Picture         =   "MDIFrm.frx":18A42
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdEasyPurchase 
         Height          =   555
         Left            =   2400
         Picture         =   "MDIFrm.frx":18E78
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Easy Purchase"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdEasySale 
         Height          =   555
         Left            =   3000
         Picture         =   "MDIFrm.frx":192A4
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Easy Sale"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdNewIssue 
         Height          =   555
         Left            =   1200
         Picture         =   "MDIFrm.frx":196B2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "NewIssue"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdSale 
         Height          =   555
         Left            =   1800
         Picture         =   "MDIFrm.frx":19AC8
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "New Sale"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdPurchase 
         Height          =   555
         Left            =   600
         Picture         =   "MDIFrm.frx":19F2A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "New Purchase"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdActivation 
         Height          =   555
         Left            =   0
         Picture         =   "MDIFrm.frx":1A3D7
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Main"
      Begin VB.Menu mnuAddUser 
         Caption         =   "Add User"
      End
      Begin VB.Menu SS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogin 
         Caption         =   "User Login"
      End
      Begin VB.Menu sa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuRecords 
      Caption         =   "&Records"
      Begin VB.Menu mnuEmployee 
         Caption         =   "Employee"
      End
      Begin VB.Menu mnuOutlet 
         Caption         =   "Outlet"
      End
      Begin VB.Menu mnuCategory 
         Caption         =   "Category"
      End
      Begin VB.Menu mnuProduct 
         Caption         =   "Product"
      End
   End
   Begin VB.Menu mnuOrder 
      Caption         =   "&Order"
      Visible         =   0   'False
      Begin VB.Menu mnuSim 
         Caption         =   "Sim"
         Begin VB.Menu mnuSAdd 
            Caption         =   "Add"
         End
         Begin VB.Menu mnuSEdit 
            Caption         =   "Edit"
         End
      End
      Begin VB.Menu mnuCoupon 
         Caption         =   "Coupon"
         Begin VB.Menu mnuCAdd 
            Caption         =   "&Add"
         End
         Begin VB.Menu mnuCEdit 
            Caption         =   "E&dit"
         End
      End
      Begin VB.Menu mnuEcharge 
         Caption         =   "E-Charge"
         Begin VB.Menu mnuEAdd 
            Caption         =   "Add"
         End
         Begin VB.Menu mnuEEdit 
            Caption         =   "Edit"
         End
      End
   End
   Begin VB.Menu mnuDealer 
      Caption         =   "&Dealer"
      Begin VB.Menu mnuPurchase1 
         Caption         =   "Purchase"
         Begin VB.Menu mnuPurchase 
            Caption         =   "Purchase"
         End
         Begin VB.Menu mnuPurchaseReturn 
            Caption         =   "Purchase Return"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuEasycharge 
         Caption         =   "Easy Charge"
         Begin VB.Menu mnuEpurchase 
            Caption         =   "Easy Purchase"
         End
         Begin VB.Menu mnuEsale 
            Caption         =   "Easy Sale"
         End
      End
      Begin VB.Menu mnuissue 
         Caption         =   "Issue"
         Begin VB.Menu mnunewIssue 
            Caption         =   "New Issue"
         End
         Begin VB.Menu mnureturn 
            Caption         =   "Issue Return"
         End
      End
      Begin VB.Menu mnuSale1 
         Caption         =   "Sale"
         Begin VB.Menu mnuSale 
            Caption         =   "Sale"
         End
         Begin VB.Menu SaleReturn 
            Caption         =   "SaleReturn"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuActivation 
         Caption         =   "Purchase Activation"
         Begin VB.Menu mnuNew 
            Caption         =   "New*"
         End
         Begin VB.Menu mnuupdate 
            Caption         =   "Update"
         End
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "R&eport"
      Begin VB.Menu mnuActivationrpt 
         Caption         =   "Activation Details"
      End
      Begin VB.Menu mnuSaleReport 
         Caption         =   "Sale && Stock Report"
      End
      Begin VB.Menu mnuPurchaserpt 
         Caption         =   "Easypurchse Report"
      End
      Begin VB.Menu mnuEsalerpt 
         Caption         =   "EasySales Report"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnudatabase 
         Caption         =   "&Database Utilities"
         Begin VB.Menu mnuBackup 
            Caption         =   "&Backup Database"
         End
         Begin VB.Menu mnuRestore 
            Caption         =   "Database Restore"
         End
      End
      Begin VB.Menu mnuChangepassword 
         Caption         =   "Change Password"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu seanet 
         Caption         =   "About Seanettechnologies"
      End
   End
End
Attribute VB_Name = "MDIFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================
'+Author:   Sabith Kp  +
'=======================

Public Function ShowForm()
    
    'default
    bUserLoggedOn = False
    
    'show form
    Me.WindowState = vbMaximized
    Me.Show
    DoEvents
    
    'show weclome
    'frmWelcome.ShowForm
    
    'unload splash
    frmSplash.UnloadSplash
    
BeginLogin:
    If AnyUserExist = False Then
        If frmUserEntry.ShowAddAdmin = False Then
            Unload Me
        Else
            GoTo BeginLogin
        End If
    Else
        If frmLogin.ShowForm = False Then
            Unload Me
            Exit Function
        End If
    End If
    
    'set log flag
    bUserLoggedOn = True
    
      
    'set UI
    'current user info
    'lblCurrentUser.Caption = CurrentUser.UserID
    
    'set date
    'timeUpdateDate_Timer
    
    
    'if the current user is not the administrator,
    'disable user related menus
    If LCase(Trim(CurrentUser.UserID)) <> "administrator" Then
        'mnuAddUser.Enabled = False
        'mnuManageuser.Enabled = False
        'mnuAgent.Enabled = False
        'mnuAddMbr.Enabled = False
        'mnup.Enabled = False
        'mnuGift.Enabled = False
        'mnuVillage.Enabled = False
        'listQL.Enabled = False
    Else
        'mnuAddUser.Enabled = True
        'mnuManageuser.Enabled = True
        'mnuAgent.Enabled = True
        'mnuAddMbr.Enabled = True
        ' mnup.Enabled = True
        'mnuGift.Enabled = True
        'mnuVillage.Enabled = True
        'listQL.Enabled = True
    End If
            'mnulogoff.Caption = "Logoff  " & CurrentUser.UserID


End Function

Private Sub ds_Click()
End Sub

Private Sub cmdNew_Click()

End Sub

Private Sub cmdActivation_Click()
mnuNew_Click
End Sub

Private Sub cmdEasyPurchase_Click()
mnuEpurchase_Click
End Sub

Private Sub cmdEasySale_Click()
mnuEsale_Click
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdNewIssue_Click()
mnunewIssue_Click
End Sub

Private Sub cmdPurchase_Click()
mnuPurchase_Click
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub cmdSale_Click()
mnuSale_Click
End Sub

Private Sub MDIForm_Load()

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

OpenWeb "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67718&lngWId=1", Me.Hwnd
End Sub

Private Sub mnuAbout_Click()
frmSplash.ShowForm
End Sub

Private Sub mnuActivationrpt_Click()
frmActivationDetails.Show
End Sub

Private Sub mnuAddUser_Click()
frmUserEntry.ShowAdd
End Sub

Private Sub mnuBackup_Click()
frmDBBackup.ShowForm
End Sub

Private Sub mnuCAdd_Click()
frm_aed_Coupon.ShowAdd
End Sub



Private Sub mnuCategory_Click()
frmAllCatogory.Show
End Sub

Private Sub mnuCEdit_Click()
Active_Form = "Coupon"

frmSearch.Show vbModal
End Sub

Private Sub mnuChangepassword_Click()
If frmUserEntry.ShowEdit("Administrator") = True Then
Exit Sub
End If
End Sub

Private Sub mnuEEdit_Click()
Active_Form = "ECharge"
frmSearch.ShowForm
End Sub

Private Sub mnuEmployee_Click()
frmAllEmploy.ShowForm
End Sub

Private Sub mnuEpurchase_Click()
frmEpurchase.ShowAdd
End Sub

Private Sub mnuEsale_Click()
frmEsale.ShowAdd
End Sub

Private Sub mnuEsalerpt_Click()
frmEReport.Caption = "Easy Sale Report"
frmEReport.Show vbModal

End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub



Private Sub mnuLogin_Click()
frmLogin.ShowForm
End Sub

Private Sub mnuNew_Click()
frm_aed_Activation.ShowAdd
End Sub

Private Sub mnunewIssue_Click()
frm_aed_issue.Caption = "New Issue"
frm_aed_issue.Show vbModal
End Sub

Private Sub mnuOutlet_Click()
frmAllOutlet.ShowForm
End Sub

Private Sub mnuProduct_Click()
frmAllProduct.ShowForm
End Sub

Private Sub mnuPurchase_Click()
frmPurchase.Caption = "New Purchase"

frmPurchase.Show vbModal
End Sub

Private Sub mnuPurchaseReturn_Click()
frmPurchase.Caption = "Purchase Return"

frmPurchase.Show vbModal
End Sub

Private Sub mnuPurchaserpt_Click()
frmEReport.Caption = "Easy Purchase Report"
frmEReport.Show vbModal
End Sub

Private Sub mnuRestore_Click()
frmRestore.ShowForm
End Sub

Private Sub mnureturn_Click()
frm_aed_issue.Caption = "Issue Return"
frm_aed_issue.Show vbModal
End Sub

Private Sub mnuSAdd_Click()
frm_aed_Sim.ShowAdd
End Sub

Private Sub mnuSale_Click()
frm_aed_Sale.Caption = "New Sale"
frm_aed_Sale.Show vbModal
End Sub

Private Sub mnuSaleReport_Click()
8 frmReport.Show vbModal
End Sub

Private Sub mnuSEdit_Click()
Active_Form = "Sim"
frmSearch.ShowForm

End Sub

Private Sub mnuupdate_Click()
Active_Form = "Activation"
frmSearch.ShowForm

End Sub

Private Sub SaleReturn_Click()
frm_aed_Sale.Caption = "Sale Return"
frm_aed_Sale.Show vbModal
End Sub

Private Sub seanet_Click()
OpenWeb "http://www.seanettechnologies.com", Me.Hwnd
End Sub

VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form MDILMS 
   BackColor       =   &H00808000&
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   FillColor       =   &H8000000F&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000080&
   Icon            =   "MDILMS.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport RptControl 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   0   'False
      WindowMaxButton =   0   'False
      WindowMinButton =   0   'False
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Menu KimGl_Mtn 
      Caption         =   "Maintain"
      Begin VB.Menu FCM_Facility 
         Caption         =   "Company Operations"
      End
      Begin VB.Menu FCM_CIBClass 
         Caption         =   "CIB Classifications"
      End
      Begin VB.Menu lms_doc 
         Caption         =   "Document Process"
      End
      Begin VB.Menu FCM_CustRegis 
         Caption         =   "Customer Registration"
      End
      Begin VB.Menu FCM_LMSRecover 
         Caption         =   "Recovery Officers"
      End
      Begin VB.Menu Lm_Cib1 
         Caption         =   "CIB-1"
      End
   End
   Begin VB.Menu FCM_Trans 
      Caption         =   "Transactions"
      Begin VB.Menu Lm_leaseoffer 
         Caption         =   "Lease Offer"
      End
      Begin VB.Menu lms_creditmemo 
         Caption         =   "Credit Memo"
      End
      Begin VB.Menu FCM_LeaseAgree 
         Caption         =   "Lease Agreement"
      End
      Begin VB.Menu FCM_LMSAttrib 
         Caption         =   "Lease Attributes"
      End
      Begin VB.Menu FCM_LMSPmts 
         Caption         =   "Payments"
      End
      Begin VB.Menu lm_insurancePayment 
         Caption         =   "Insurance Payments"
      End
      Begin VB.Menu lm_Cdocument 
         Caption         =   "Lease Documents"
      End
      Begin VB.Menu LM_Comments 
         Caption         =   "Comments"
      End
      Begin VB.Menu FCM_LMSPosting 
         Caption         =   "Ledger Posting"
      End
      Begin VB.Menu FCM_LMSGLPosting 
         Caption         =   "Post Accruals To GL"
      End
   End
   Begin VB.Menu KimGl_Rpt 
      Caption         =   "Reports"
      Begin VB.Menu FCM_LMSClientShdl 
         Caption         =   "Client Payment Schedule"
      End
      Begin VB.Menu FCM_LMSIntrShdl 
         Caption         =   "Internal Payment Schedule"
         Begin VB.Menu FCM_IntrShdl1 
            Caption         =   "Base To Accounts"
         End
         Begin VB.Menu FCM_IntrShdl2 
            Caption         =   "Base To Recovery"
         End
      End
      Begin VB.Menu LMS_Line1 
         Caption         =   "-"
      End
      Begin VB.Menu LM_Confirm 
         Caption         =   "Receivable As On"
      End
      Begin VB.Menu FCM_LMS 
         Caption         =   "Customer Statement"
      End
      Begin VB.Menu FCM_LMSProj 
         Caption         =   "Inflow Projections"
      End
      Begin VB.Menu FCM_LMSReminders 
         Caption         =   "Lease Reminders"
      End
      Begin VB.Menu FCM_RecoveryStatus 
         Caption         =   "Recovery Statements"
         Begin VB.Menu FCM_LMSRecovery 
            Caption         =   "Overall Status"
         End
         Begin VB.Menu lm_OverstatCwise 
            Caption         =   "Overall Status (Credit Officer)"
         End
         Begin VB.Menu FCM_LMSOverdue 
            Caption         =   "Overdue Status"
         End
         Begin VB.Menu Lm_IndPer 
            Caption         =   "Individual Performance"
         End
      End
      Begin VB.Menu LMS_Line2 
         Caption         =   "-"
      End
      Begin VB.Menu FCM_LMSRegister 
         Caption         =   "Lease Asset Register"
      End
      Begin VB.Menu FCM_LMSAPortfilo 
         Caption         =   "Lease Active Portfolia"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu FCM_LMSReceipts 
         Caption         =   "Receipts"
      End
      Begin VB.Menu FCM_LMSAgeing 
         Caption         =   "Lease Rentals Ageing"
      End
      Begin VB.Menu lm_insurance 
         Caption         =   "Insurance Status"
      End
      Begin VB.Menu FCM_LMCustHist 
         Caption         =   "Customer History Sheet"
      End
      Begin VB.Menu FCM_LMSpecial 
         Caption         =   "Special Reports"
         Begin VB.Menu FCM_LMAssetExp 
            Caption         =   "Asset wise Exposure"
         End
         Begin VB.Menu FCM_LMSector 
            Caption         =   "Sector wise Exposure"
         End
         Begin VB.Menu FCM_LMEntity 
            Caption         =   "Entity wise Exposure"
         End
         Begin VB.Menu FCM_LMPortfolio 
            Caption         =   "Portfolio Analysis"
         End
         Begin VB.Menu FCM_LMRecoAnalysis 
            Caption         =   "Recovery Analysis"
            Begin VB.Menu FCM_LMReco1 
               Caption         =   "Based On Total Rentals"
            End
            Begin VB.Menu FCM_LMReco2 
               Caption         =   "Based On Portfolio"
            End
         End
         Begin VB.Menu FCM_LMLitigation 
            Caption         =   "Litigation portfolio Analysis"
         End
      End
      Begin VB.Menu FCM_LMAccr 
         Caption         =   "Lease Accruals"
      End
   End
   Begin VB.Menu KimGl_Ret 
      Caption         =   "Return"
   End
End
Attribute VB_Name = "MDILMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FCM_CIBClass_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
  frmARClasses.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_CustList_Click()
'   With CntrlReports
'        .ReportFileName = App.Path & Gs_CstRepoPath & "\CustomerInfo.RPT"
'        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
'        .Formulas(1) = "ReportName = 'Customer Listing'"
'       '.Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & dtpto & "'"
'        .Action = 1
'   End With
End Sub

Private Sub FCM_CustRegis_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
    frmCustomer.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_Facility_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
   frmFacilityType.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_IntrShdl1_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
    FrmLMSReports.Caption = "Internal Payment Schedule [Accounts]"
    FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_IntrShdl2_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
    FrmLMSReports.Caption = "Internal Payment Schedule [Recovery]"
    FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_LeaseAgree_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
  frmLeaseAgree.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_LMAccr_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
   FrmLMSReports.Caption = "Lease Rental Accruals"
   FrmLMSReports.dtpason.Format = dtpCustom
   FrmLMSReports.dtpason.CustomFormat = "MM/yyyy"
   FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_LMAssetExp_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
FrmLMSReports.Caption = "Asset Wise Exposure"
FrmLMSReports.txtbranchcode.Enabled = False
FrmLMSReports.Command5.Enabled = False
FrmLMSReports.txtCustNO.Enabled = False
FrmLMSReports.txtleaseno.Enabled = False
FrmLMSReports.Chvdate.Enabled = False
FrmLMSReports.Command1.Enabled = False
FrmLMSReports.cmdLookup.Enabled = False
FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_LMCustHist_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
    FrmLMSReports.Caption = "Customer History Sheet"
    FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_LMEntity_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
FrmLMSReports.Caption = "Entity Wise Exposure"
FrmLMSReports.txtbranchcode.Enabled = False
FrmLMSReports.Command5.Enabled = False
FrmLMSReports.txtCustNO.Enabled = False
FrmLMSReports.txtleaseno.Enabled = False
FrmLMSReports.Chvdate.Enabled = False
FrmLMSReports.Command1.Enabled = False
FrmLMSReports.cmdLookup.Enabled = False
FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_LMS_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
    FrmLMSReports.Caption = "Balance Confirmation Statement"
    FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_LMSAgeing_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
   FrmARAgging.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_LMSAPortfilo_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
  FrmLMSReports.Caption = "Portfolia"
  FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""

End Sub

Private Sub FCM_LMSAttrib_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
  frmLeaseAttrib.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_LMSClientShdl_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
    FrmLMSReports.Caption = "Client Payment Schedule"
    FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_LMSector_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
FrmLMSReports.Caption = "Sector Wise Exposure"
FrmLMSReports.txtbranchcode.Enabled = False
FrmLMSReports.Command5.Enabled = False
FrmLMSReports.txtCustNO.Enabled = False
FrmLMSReports.txtleaseno.Enabled = False
FrmLMSReports.Chvdate.Enabled = False
FrmLMSReports.Command1.Enabled = False
FrmLMSReports.cmdLookup.Enabled = False
FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_LMSGLPosting_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
   FrmLMSReports.Caption = "Post Accurals To General Ledger"
   FrmLMSReports.dtpason.Format = dtpCustom
   FrmLMSReports.dtpason.CustomFormat = "MM/yyyy"
   FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_LMSOverdue_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
   FrmLMSReports.Caption = "Overdue Statement"
   FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_LMSPmts_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
    frmRentlPmt.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_LMSPosting_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
  frmAskPosting.Caption = "Posting To Lease Ledger"
  frmAskPosting.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

'Private Sub FCM_LMSProjections_Click()
'    FrmLMSReports.Caption = "Projections"
'    FrmLMSReports.Show
'End Sub

Private Sub FCM_LMSProj_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
   FrmLMSReports.Caption = "Projections"
   FrmLMSReports.Chvdate.Enabled = True
   FrmLMSReports.Chvdate.Caption = "Incl.Terminated Lease."
   FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_LMSReceipts_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
   FrmLMSReports.Caption = "Receipts Statement"
   FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_LMSRecover_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
   frmRecoverer.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_LMSRecovery_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
   FrmLMSReports.Caption = "Recovery Statement"
   FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub



Private Sub FCM_LMSRegister_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
  FrmLMSReports.Caption = "Register"
  FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
'
'MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
'Dim choice As Integer
'    With RptControl
'      choice = SetErr("Include Terminated Leases.  ", vbYesNo)
'     .DiscardSavedData = True
'     .ReportFileName = App.Path & Gs_ARRepoPath & "\LM_LeaseInfo.RPT"
'     .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
'     .Formulas(1) = "ReportName = '" & IIf(choice = vbNo, "Active ", "") & "' & 'Lease Asset Register'"
'      If choice = vbNo Then .SelectionFormula = "{LM_LeaseInfo.ActiveStatus} >=1 And {LM_LeaseInfo.Compcode} = '" & Gs_compcode & "'"
'      .Action = 1
'    End With
'MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub FCM_LMSReminders_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
  FrmLMSReports.Caption = "Bills of Lease Facilities"
  FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub Form_Load()
'  FCM_LMSReceipts.Enabled = chkRights("RPTLMRECPT")
'  FCM_LMSRegister.Enabled = chkRights("RPTLMSREGS")
'  FCM_LMSProj.Enabled = chkRights("RPTLMSPROJ")
'  FCM_LMSPosting.Enabled = chkRights("RPTLMSPOST")
'  FCM_IntrShdl1.Enabled = chkRights("RPTLMINTE1")
'  FCM_IntrShdl2.Enabled = chkRights("RPTLMINTE2")
'  FCM_LMS.Enabled = chkRights("RPTLMBALNC")
'  FCM_LMSClientShdl.Enabled = chkRights("RPTLMCLIET")
'  FCM_LMSReminders.Enabled = chkRights("RPTLMBILL0")
'  FCM_LMCustHist.Enabled = chkRights("RPTLMCSTHS")
'  FCM_LMAccr.Enabled = chkRights("RPTLMACCRL")
'  FCM_LMSOverdue.Enabled = chkRights("RPTLMOVERD")
'
'  FCM_LMSGLPosting.Enabled = chkRights("LMSGLPOSTG")
'  FCM_LMSPosting.Enabled = chkRights("LMLMSPOSTG")
  
End Sub
Private Sub KimGl_Ret_Click()
   MDIForm1.Caption = "FSIB Financials."
   MDIForm1.StatusBar1.Panels(2).Text = ""
   MDIForm1.StatusBar1.Panels(4).Text = ""
   MDIForm1.StatusBar1.Panels(6).Text = ""
   Para_Rs.Filter = adFilterNone
   ParaCntr_Rs.Close
   Pr_BranchName.Close
   Unload Me
End Sub

Private Sub lm_Cdocument_Click()
FrmLmDoc.Show
End Sub

Private Sub Lm_Cib1_Click()
frmlmscib1.Show
End Sub

Private Sub LM_Comments_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
  frmlmcomments.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub LM_Confirm_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
    FrmLMSReports.Caption = "Statement of Receivables"
    FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub Lm_Document_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
    FrmLmDocument.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub Lm_IndPer_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
   FrmLMSReports.Caption = "Individual Performance Report"
   FrmLMSReports.txtrecoverer.Visible = True
   FrmLMSReports.txtrecoverer.Top = FrmLMSReports.txtleaseno.Top
   FrmLMSReports.Label5.Top = FrmLMSReports.txtleaseno.Top
   FrmLMSReports.Command2.Top = FrmLMSReports.Command1.Top
   FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub lm_insurance_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
   FrmLMSReports.Caption = "Insurance Status"
   FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub lm_insurancePayment_Click()
frmLMInsur.Show
End Sub

Private Sub Lm_leaseoffer_Click()
frmLeaseoffer.Show
End Sub

Private Sub lm_OverstatCwise_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
   FrmLMSReports.Caption = "Credit Officer Wise"
   FrmLMSReports.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub lms_creditmemo_Click()
frmCreditMemo.Show
End Sub

Private Sub lms_doc_Click()
FrmLmDocument.Show
End Sub

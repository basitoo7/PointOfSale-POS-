VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form MDIGL 
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
   Icon            =   "MDIgl.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport rptNotes 
      Left            =   2070
      Top             =   630
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport rptReports 
      Left            =   840
      Top             =   1815
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport rptcoa 
      Left            =   675
      Top             =   510
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
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
      Begin VB.Menu KimGl_S0 
         Caption         =   "Control Accounts"
         Shortcut        =   ^C
      End
      Begin VB.Menu KimGl_S1 
         Caption         =   "Detail Accounts"
         Shortcut        =   ^D
      End
      Begin VB.Menu KimGl_S2 
         Caption         =   "Sub Ledger Accounts"
         Shortcut        =   ^U
      End
      Begin VB.Menu KimGl_Detail 
         Caption         =   "Subsidiary Sub Ledger A/C"
         Shortcut        =   ^I
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu Balance_Sheet_Heads_Setup 
         Caption         =   "Balance Sheet Heads Setup"
      End
      Begin VB.Menu Balance_Sheet_Notes_Setup 
         Caption         =   "Balance Sheet Notes Setup"
      End
      Begin VB.Menu Balance_Sheet_Notes_Items_Setup 
         Caption         =   "Balance Sheet Notes Items Setup"
      End
      Begin VB.Menu KimGl_L4 
         Caption         =   "-"
      End
      Begin VB.Menu Profit_and_Loss_Heads_Setup 
         Caption         =   "Profit and Loss Heads Setup"
      End
      Begin VB.Menu Profit_and_Loss_Notes_Setup 
         Caption         =   "Profit and Loss Notes Setup"
      End
      Begin VB.Menu Profit_and_Loss_Notes_Items_Setup 
         Caption         =   "Profit and Loss Notes Items Setup"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu KimGl_Vtype 
         Caption         =   "Voucher Types"
      End
      Begin VB.Menu KimGl_Budg 
         Caption         =   "Budgting and Forcasting"
         Visible         =   0   'False
      End
      Begin VB.Menu KimGl_Rptcust 
         Caption         =   "Report Customization"
         Visible         =   0   'False
      End
      Begin VB.Menu KimGl_CustRt 
         Caption         =   "Customized Report A/c Routing"
         Visible         =   0   'False
      End
      Begin VB.Menu Kimgl_Line22 
         Caption         =   "-"
      End
      Begin VB.Menu Kimgl_factivYear 
         Caption         =   "Set Active Financial Year"
      End
      Begin VB.Menu Kimgl_UpdOpen 
         Caption         =   "Update Opening Balances"
         Visible         =   0   'False
      End
      Begin VB.Menu KimGl_l1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Import_COA 
         Caption         =   "Import COA/BS/PL Heads"
         Visible         =   0   'False
      End
      Begin VB.Menu line21 
         Caption         =   "-"
      End
      Begin VB.Menu Return_Home 
         Caption         =   "Return"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu KimGl_Trns 
      Caption         =   "Transactions"
      Begin VB.Menu KimGl_Vchr 
         Caption         =   "GL Transaction"
         Shortcut        =   ^G
      End
      Begin VB.Menu GL_Transaction_Multiple 
         Caption         =   "GL Transaction Multiple"
         Shortcut        =   ^M
      End
      Begin VB.Menu Scan_Voucher_Documents 
         Caption         =   "Scan Voucher Documents"
         Shortcut        =   ^S
         Visible         =   0   'False
      End
      Begin VB.Menu line11 
         Caption         =   "-"
      End
      Begin VB.Menu Payment_to_Vendors_By_Bank 
         Caption         =   "Payment to Vendors By Bank"
         Shortcut        =   ^N
      End
      Begin VB.Menu Payment_to_Vendors_By_Cash 
         Caption         =   "Payment to Vendors By Cash"
         Shortcut        =   ^A
      End
      Begin VB.Menu Posting_To_GL 
         Caption         =   "GL Posting"
         Visible         =   0   'False
      End
      Begin VB.Menu GL_Unposting 
         Caption         =   "GL Unposting"
         Visible         =   0   'False
      End
      Begin VB.Menu Update_Voucher 
         Caption         =   "Update Voucher"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu KimGl_Rpt 
      Caption         =   "Reports"
      Begin VB.Menu KimGl_EdLst 
         Caption         =   "Edit List"
      End
      Begin VB.Menu Printvoucher 
         Caption         =   "Print Voucher"
         Shortcut        =   ^P
      End
      Begin VB.Menu Unposted_Vouchers 
         Caption         =   "Unposted Vouchers"
         Visible         =   0   'False
      End
      Begin VB.Menu line10 
         Caption         =   "-"
      End
      Begin VB.Menu KimGl_Ledg 
         Caption         =   "General Ledger"
         Shortcut        =   ^L
      End
      Begin VB.Menu General_Ledger_Internal 
         Caption         =   "General Ledger (Internal)"
         Visible         =   0   'False
      End
      Begin VB.Menu DailyActivity 
         Caption         =   "Daily Activity"
         Visible         =   0   'False
      End
      Begin VB.Menu ClosingBalance 
         Caption         =   "Closing Balance"
      End
      Begin VB.Menu Kimgl_BBook 
         Caption         =   "Bank Book"
         Visible         =   0   'False
      End
      Begin VB.Menu Kimgl_CBook 
         Caption         =   "Cash Book"
         Visible         =   0   'False
      End
      Begin VB.Menu Kimgl_JBook 
         Caption         =   "Journal"
         Visible         =   0   'False
      End
      Begin VB.Menu KimGl_L2 
         Caption         =   "-"
      End
      Begin VB.Menu KimGl_Trl 
         Caption         =   "Trial Balance"
         Begin VB.Menu FCM_Sub0 
            Caption         =   "Detail"
            Shortcut        =   ^T
         End
         Begin VB.Menu Transaction_Base_Trial 
            Caption         =   "Transaction Base Trial"
         End
         Begin VB.Menu FCM_Sub1 
            Caption         =   "Periodic"
         End
         Begin VB.Menu FCM_Sub2 
            Caption         =   "Closing"
         End
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu KimGl_Bs 
         Caption         =   "Balance Sheet"
         Begin VB.Menu BalanceSheet 
            Caption         =   "Balance Sheet"
            Shortcut        =   ^B
         End
         Begin VB.Menu Balance_Sheet_Notes 
            Caption         =   "Balance Sheet Notes"
         End
         Begin VB.Menu Balance_Sheet_Periodic 
            Caption         =   "Balance Sheet (Periodic)"
         End
         Begin VB.Menu Balance_Sheet_With_GL_Codes 
            Caption         =   "Balance Sheet (Detail)"
         End
         Begin VB.Menu Balance_Sheet_Detail_Periodic 
            Caption         =   "Balance Sheet (Detail-Periodic)"
         End
      End
      Begin VB.Menu KmGl_L3 
         Caption         =   "-"
      End
      Begin VB.Menu KimGl_PfNote 
         Caption         =   "Profit and Loss "
         Begin VB.Menu Profit_and_Loss 
            Caption         =   "Profit and Loss"
            Shortcut        =   ^F
         End
         Begin VB.Menu Profit_and_Loss_Notes 
            Caption         =   "Profit and Loss Notes"
         End
         Begin VB.Menu Profit_and_Loss_Periodic 
            Caption         =   "Profit and Loss (Periodic)"
         End
         Begin VB.Menu Profit_and_Loss_With_GL_Codes 
            Caption         =   "Profit and Loss (Detail)"
         End
         Begin VB.Menu Profit_and_Loss_Detail_Periodic 
            Caption         =   "Profit and Loss (Detail-Periodic)"
         End
      End
      Begin VB.Menu KimGl_l5 
         Caption         =   "-"
      End
      Begin VB.Menu KimGl_CRpt 
         Caption         =   "Customized Reports"
         Visible         =   0   'False
      End
      Begin VB.Menu KimGl_Listing 
         Caption         =   "Listings"
         Begin VB.Menu KimGl_COA 
            Caption         =   "Chart of Accounts"
         End
         Begin VB.Menu KimGl_LVtype 
            Caption         =   "Voucher Types"
         End
         Begin VB.Menu KimGl_Format 
            Caption         =   "Customized Reports Format"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu lline5 
         Caption         =   "-"
      End
      Begin VB.Menu Income_Expense_Report 
         Caption         =   "Income/Expense Report"
      End
      Begin VB.Menu line111 
         Caption         =   "-"
      End
      Begin VB.Menu Account_Receiable_Payable 
         Caption         =   "Account Receiable/Payable"
      End
      Begin VB.Menu Account_Payable_Aging 
         Caption         =   "Account Payable Aging"
      End
      Begin VB.Menu Account_Receiable_Aging 
         Caption         =   "Account Receiable Aging"
      End
   End
   Begin VB.Menu KimGl_Ret 
      Caption         =   "&Return"
   End
End
Attribute VB_Name = "MDIGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub aa_Click()

Dim ls_sql As String
Dim ls_vno As String
Dim pr_dumy1 As New Recordset
Dim PR_Dumy2 As New Recordset
Dim PR_dumy3 As New Recordset


ls_sql = "Select vchrtype from gl_ref where compcode = '" & Gs_compcode & "' and value_date >= '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' and value_date <= '" & Format(Gs_FnEndPeriod, "YYYY/MM/DD") & "' and vchrtype in ('BPP') group by vchrtype"
pr_dumy1.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy1.EOF Then
Do While Not pr_dumy1.EOF

ls_sql = "Select * from gl_ref where compcode = '" & Gs_compcode & "' and value_date >= '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' and value_date <= '" & Format(Gs_FnEndPeriod, "YYYY/MM/DD") & "' and vchrtype = '" & pr_dumy1("Vchrtype") & "' order by value_date"
If PR_Dumy2.State = 1 Then PR_Dumy2.Close
PR_Dumy2.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not PR_Dumy2.EOF Then
Do While Not PR_Dumy2.EOF
ls_vno = vno1(PR_Dumy2("Value_date"), pr_dumy1("Vchrtype"), Gs_compcode, PR_Dumy2("Branchcode"))

ls_sql = "Update gl_ref set newVoucherno = '" & ls_vno & "' where compcode = '" & Gs_compcode & "' and branchcode = '" & PR_Dumy2("Branchcode") & "' and vchrtype = '" & PR_Dumy2("Vchrtype") & "' and voucher_No = '" & PR_Dumy2("Voucher_No") & "' and Value_date = '" & Format(PR_Dumy2("Value_date"), "YYYY/MM/DD") & "'"
gc_dbcon.Execute ls_sql
ls_sql = "Update gl_Trans set newVoucherno = '" & ls_vno & "' where compcode = '" & Gs_compcode & "' and branchcode = '" & PR_Dumy2("Branchcode") & "' and vchrtype = '" & PR_Dumy2("Vchrtype") & "' and voucher_No = '" & PR_Dumy2("Voucher_No") & "' and Value_date = '" & Format(PR_Dumy2("Value_date"), "YYYY/MM/DD") & "'"
gc_dbcon.Execute ls_sql


PR_Dumy2.MoveNext
Loop
PR_Dumy2.Close
End If

pr_dumy1.MoveNext
Loop
End If
'ls_sql = "update gl_ref set voucher_no = newvoucherno where compcode = '" & Gs_compcode & "' and value_date >= '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' and value_date <= '" & Format(Gs_FnEndPeriod, "YYYY/MM/DD") & "'"
'gc_dbcon.Execute ls_sql
'ls_sql = "update gl_trans set voucher_no = newvoucherno where compcode = '" & Gs_compcode & "' and value_date >= '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' and value_date <= '" & Format(Gs_FnEndPeriod, "YYYY/MM/DD") & "'"
'gc_dbcon.Execute ls_sql

Call MsgBox("Successfully Update")
End Sub



Private Sub AnalysisReport_Click()
frmanalysisReport.Show
End Sub

Private Sub Account_Payable_Aging_Click()
frmPoVendorAging.Show
End Sub

Private Sub Account_Receiable_Aging_Click()
frmSoClientAging.Show
End Sub

Private Sub Account_Receiable_Payable_Click()
FrmaskdateClosing.Show
End Sub

Private Sub Balance_Sheet_Detail_Periodic_Click()
Frmbsheetrpt1.Caption = Balance_Sheet_Detail_Periodic.Caption
Frmbsheetrpt1.Show
End Sub

Private Sub Balance_Sheet_Heads_Setup_Click()
frmBalancesheet0.Show
End Sub

Private Sub Balance_Sheet_Notes_Click()
Frmbsheetrpt.Caption = Balance_Sheet_Notes.Caption
Frmbsheetrpt.Show
End Sub

Private Sub Balance_Sheet_Notes_Items_Setup_Click()
frmbalancesheet3.Show
End Sub

Private Sub Balance_Sheet_Notes_Setup_Click()
frmbalancesheet2.Show
End Sub

Private Sub Balance_Sheet_Periodic_Click()
Frmbsheetrpt1.Caption = Balance_Sheet_Periodic.Caption
Frmbsheetrpt1.Show
End Sub

Private Sub Balance_Sheet_With_GL_Codes_Click()
Frmbsheetrpt.Caption = Balance_Sheet_With_GL_Codes.Caption
Frmbsheetrpt.Show
End Sub

Public Sub BalanceSheet_Click()
Frmbsheetrpt.Caption = BalanceSheet.Caption
Frmbsheetrpt.Show
End Sub

Private Sub ClosingBalance_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
   frmAccLedger.Caption = "Closing Balances"
   frmAccLedger.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub DailyActivity_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
   frmAccLedger.Caption = DailyActivity.Caption
   frmAccLedger.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Public Sub FCM_Sub0_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
Frmaskdate.Caption = "Trial Balance (Detail)"
Frmaskdate.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub
Private Sub FCM_Sub1_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
Frmaskdateperiodic.Caption = "Trial Balance (Periodic)"
Frmaskdateperiodic.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub
Private Sub FCM_Sub2_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
Frmaskdate.Caption = "Trial Balance (Closing)"
Frmaskdate.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub
Private Sub Form_Activate()
  'Kimgl_factivYear.Enabled = Not Gl_Demo
  'Kimgl_UpdOpen.Enabled = Not Gl_Demo
  KimGl_S0.Visible = IIf(gn_Maxlevels >= 0, True, False)
  KimGl_S1.Visible = IIf(gn_Maxlevels >= 1, True, False)
  KimGl_S2.Visible = IIf(gn_Maxlevels >= 2, True, False)
  
End Sub

Private Sub Form_Load()
KimGl_S0.Enabled = chkRights1("LEDGR00001")
KimGl_S1.Enabled = chkRights1("LEDGR00002")
KimGl_S2.Enabled = chkRights1("LEDGR00003")
KimGl_Detail.Enabled = chkRights1("LEDGR00004")

MDIForm1.Toolbar3.Buttons(1).ButtonMenus(1).Enabled = KimGl_S0.Enabled
MDIForm1.Toolbar3.Buttons(1).ButtonMenus(2).Enabled = KimGl_S1.Enabled
MDIForm1.Toolbar3.Buttons(1).ButtonMenus(3).Enabled = KimGl_S2.Enabled
MDIForm1.Toolbar3.Buttons(1).ButtonMenus(4).Enabled = KimGl_Detail.Enabled


Balance_Sheet_Heads_Setup.Enabled = chkRights1("LEDGR00005")
Balance_Sheet_Notes_Setup.Enabled = chkRights1("LEDGR00006")
Balance_Sheet_Notes_Items_Setup.Enabled = chkRights1("LEDGR00007")
Profit_and_Loss_Heads_Setup.Enabled = chkRights1("LEDGR00008")
Profit_and_Loss_Notes_Setup.Enabled = chkRights1("LEDGR00009")
Profit_and_Loss_Notes_Items_Setup.Enabled = chkRights1("LEDGR00010")
KimGl_Vtype.Enabled = chkRights1("LEDGR00011")
KimGl_Vchr.Enabled = chkRights1("LEDGR00012")

GL_Transaction_Multiple.Enabled = chkRights1("LEDGR00013")
MDIForm1.Toolbar3.Buttons(2).Enabled = GL_Transaction_Multiple.Enabled
Payment_to_Vendors_By_Bank.Enabled = chkRights1("LEDGR00014")
Payment_to_Vendors_By_Cash.Enabled = chkRights1("LEDGR00015")
KimGl_EdLst.Enabled = chkRights1("LEDGR00016")
Printvoucher.Enabled = chkRights1("LEDGR00017")
MDIForm1.Toolbar3.Buttons(4).Enabled = Printvoucher.Enabled
KimGl_Ledg.Enabled = chkRights1("LEDGR00018")
MDIForm1.Toolbar3.Buttons(5).Enabled = KimGl_Ledg.Enabled
ClosingBalance.Enabled = chkRights1("LEDGR00019")
KimGl_Trl.Enabled = chkRights1("LEDGR00020")
MDIForm1.Toolbar3.Buttons(6).Enabled = KimGl_Trl.Enabled
KimGl_Bs.Enabled = chkRights1("LEDGR00021")
MDIForm1.Toolbar3.Buttons(7).Enabled = KimGl_Bs.Enabled

KimGl_PfNote.Enabled = chkRights1("LEDGR00022")
MDIForm1.Toolbar3.Buttons(8).Enabled = KimGl_PfNote.Enabled

KimGl_COA.Enabled = chkRights1("LEDGR00023")
KimGl_LVtype.Enabled = chkRights1("LEDGR00024")
Income_Expense_Report.Enabled = chkRights1("LEDGR00025")
Account_Receiable_Payable.Enabled = chkRights1("LEDGR00026")
Account_Payable_Aging.Enabled = chkRights1("LEDGR00027")
Account_Receiable_Aging.Enabled = chkRights1("LEDGR00028")

MDIForm1.Toolbar2.Visible = False


End Sub

Private Sub General_Ledger_Internal_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
   frmAccLedger.Caption = "General Ledger (Internal)"
   frmAccLedger.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub


Public Sub GL_Transaction_Multiple_Click()
FrmglTrans.Show
End Sub

Private Sub Import_COA_Click()
'frmImportGlHeads.Show
End Sub

Private Sub Income_Expense_Report_Click()
frmPOPurchaseReport.ChkDetail.Visible = True
frmPOPurchaseReport.txtType.Visible = True
frmPOPurchaseReport.Caption = "Income/Expense Report"
frmPOPurchaseReport.Show
End Sub

Private Sub Kimgl_BBook_Click()
   Call SetBooks("B")
End Sub



Private Sub KimGl_BsNote_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
frmAccLedger.Caption = "Balance Sheet Notes"
frmAccLedger.Frame7.Enabled = False
'frmAccLedger.dtpfrom.Enabled = False
'frmAccLedger.StatusBar1.Visible = False

'frmAccLedger.Height = 2500
'frmAccLedger.cmdGenerate.Top = 1600
'frmAccLedger.cmdCancel.Top = 1600
frmAccLedger.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub


Private Sub Kimgl_CBook_Click()
  Call SetBooks("C")
End Sub

Private Sub KimGl_COA_Click()
'   Call ChkTempTables("COA", False)
 '  Call Module2.COAccounts
 MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
 
   With rptcoa
        .ReportFileName = App.Path & Gs_GlRepoPath & "\ChartofAccounts.RPT"
        .WindowTitle = "Chart Of Accounts"
        .SelectionFormula = "{Gl_Sub0.Compcode} = '" & Gs_compcode & "'"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
    End With
MDIForm1.StatusBar1.Panels(7).Text = ""
  '  gc_dbcon.Execute ("DROP TABLE COA;")
End Sub





Public Sub KimGl_Detail_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
gs_TableId = gn_Maxlevels
frmgldetail.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub KimGl_EdLst_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
    frmeditlist.Caption = "Edit List"
    frmeditlist.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub Kimgl_factivYear_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
    FrmFnYear.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub


Private Sub Kimgl_JBook_Click()
    Call SetBooks("J")
End Sub

Public Sub KimGl_Ledg_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
   frmAccLedger.Caption = "General Ledger"
   frmAccLedger.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub KimGl_LVtype_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
    With rptReports
        .ReportFileName = App.Path & Gs_GlRepoPath & "\VchrList.RPT"
        .WindowTitle = "Voucher Types Listing"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = 'Voucher Types Listing'"
        .SelectionFormula = "{GlVchrType.CompCode} = '" & Gs_compcode & "'"
        '.Destination = PrintOpt(Option1, Option2, Option3, Option4)
        'If Option2.Value = True Then
        '.CopiesToPrinter = IIf(Trim(Text2) <> "", Val(Text2), 1)
        'End If
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
    End With
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub


Private Sub KimGl_PL_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
  Frmaskdate.Caption = "Profit & Loss A/c"
  Frmaskdate.Label5.Visible = False
  'Frmaskdate.DTPicker2.Visible = False
  'Frmaskdate.Label1.Top = Frmaskdate.Label5.Top
  'Frmaskdate.DTPicker1.Top = Frmaskdate.DTPicker2.Top
  Frmaskdate.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Public Sub KimGl_Ret_Click()
   MDIForm1.Caption = "E-Counts 2.0"
   MDIForm1.StatusBar1.Panels(2).Text = ""
   MDIForm1.StatusBar1.Panels(4).Text = ""
   MDIForm1.StatusBar1.Panels(6).Text = ""
   MDIForm1.Toolbar2.Top = 0
'   MDIForm1.Toolbar2.Left = 0
   MDIForm1.Toolbar2.Visible = True
   MDIForm1.Toolbar3.Visible = False
   If ParaCntr_Rs.State = 1 Then ParaCntr_Rs.Close
   Unload Me
End Sub



Public Sub KimGl_S0_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
   frmGlSub0.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Public Sub KimGl_S1_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."

frmGlSub1.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Public Sub KimGl_S2_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
frmGlSub2.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub



Private Sub Kimgl_UpdOpen_Click()
  Module1.SetCalc ("U")
End Sub

Public Sub KimGl_Vchr_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
FrmglTransOthers.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub KimGl_Vtype_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
frmVchrType.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub SetBooks(Ps_SetId As String)

Dim ls_Caption As String

   ls_Caption = IIf(Ps_SetId = "B", "Bank Book", IIf(Ps_SetId = "C", "Cash Book", "Journal"))
   'frmAccLedger.dtpfrom.Enabled = True
   frmAccLedger.Frame7.Enabled = False
   'frmAccLedger.StatusBar1.Visible = False
   'frmAccLedger.Height = 2500
   'frmAccLedger.cmdGenerate.Top = 1600
   'frmAccLedger.cmdCancel.Top = 1600
   frmAccLedger.Caption = ls_Caption
   frmAccLedger.Show
End Sub




Private Sub Payment_to_Vendors_By_Bank_Click()
frmPOVendorPaymentsBank.Show
End Sub

Private Sub Payment_to_Vendors_By_Cash_Click()
frmPOVendorPaymentsCash.Show
End Sub

Public Sub Printvoucher_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
    frmeditlist.Caption = "Print Voucher"
    frmeditlist.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Public Sub Profit_and_Loss_Click()
Frmbsheetrpt.Caption = Profit_and_Loss.Caption
Frmbsheetrpt.Show
End Sub

Private Sub Profit_and_Loss_Detail_Periodic_Click()
Frmbsheetrpt1.Caption = Profit_and_Loss_Detail_Periodic.Caption
Frmbsheetrpt1.Show
End Sub

Private Sub Profit_and_Loss_Heads_Setup_Click()
frmPLSheet1.Show
End Sub

Private Sub Profit_and_Loss_Notes_Click()
Frmbsheetrpt.Caption = Profit_and_Loss_Notes.Caption
Frmbsheetrpt.Show
End Sub

Private Sub Profit_and_Loss_Notes_Items_Setup_Click()
frmPLSheet3.Show
End Sub

Private Sub Profit_and_Loss_Notes_Setup_Click()
frmPLSheet2.Show
End Sub

Private Sub Profit_and_Loss_Periodic_Click()
Frmbsheetrpt1.Caption = Profit_and_Loss_Periodic.Caption
Frmbsheetrpt1.Show
End Sub

Private Sub Profit_and_Loss_With_GL_Codes_Click()
Frmbsheetrpt.Caption = Profit_and_Loss_With_GL_Codes.Caption
Frmbsheetrpt.Show
End Sub

Private Sub Return_Home_Click()
KimGl_Ret_Click
End Sub

Public Sub Scan_Voucher_Documents_Click()
frmScanVoucherBill.Show
End Sub

Private Sub Transaction_Base_Trial_Click()
FrmTransDateTrial.Show
End Sub

Private Sub Unposted_Vouchers_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Loading Please Wait..."
    frmeditlist.Caption = "Unposted Vouchers"
    frmeditlist.Show
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub Update_Voucher_Click()
Form1.Show
End Sub

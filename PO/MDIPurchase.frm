VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form MDIPurchase 
   BackColor       =   &H00808000&
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   FillColor       =   &H00808000&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport rptListigns 
      Left            =   390
      Top             =   630
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Menu KimAr_Mt 
      Caption         =   "Maintain"
      Begin VB.Menu SuppliersSetup 
         Caption         =   "Vendor Setup"
         Shortcut        =   ^V
      End
      Begin VB.Menu Authority_Person_Setup 
         Caption         =   "Authority Person Setup"
      End
      Begin VB.Menu line10 
         Caption         =   "-"
      End
      Begin VB.Menu Vendor_Account_Setup 
         Caption         =   "Vendor Account Setup"
      End
      Begin VB.Menu Vendor_Balance_Sheet_Setup 
         Caption         =   "Vendor Balance Sheet Setup"
      End
      Begin VB.Menu Other_Account_Setup 
         Caption         =   "Purchase Account Setup"
      End
      Begin VB.Menu line24 
         Caption         =   "-"
      End
      Begin VB.Menu setactiveyear 
         Caption         =   "Set Active Financial Year"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu Sector_Setup 
         Caption         =   "Sector Setup"
      End
      Begin VB.Menu Country_Setup 
         Caption         =   "Country Setup"
      End
      Begin VB.Menu City_Setup 
         Caption         =   "City Setup"
      End
      Begin VB.Menu Tehseel_Setup 
         Caption         =   "Tehseel Setup"
      End
      Begin VB.Menu line20 
         Caption         =   "-"
      End
      Begin VB.Menu Bank_Setup 
         Caption         =   "Bank Setup"
      End
      Begin VB.Menu line21 
         Caption         =   "-"
      End
      Begin VB.Menu Return_main 
         Caption         =   "Return"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu KimAr_Trns 
      Caption         =   "Transaction"
      Begin VB.Menu Demand_Note 
         Caption         =   "Demand Note"
         Visible         =   0   'False
      End
      Begin VB.Menu Purchase_Order 
         Caption         =   "Purchase Order"
      End
      Begin VB.Menu line15 
         Caption         =   "-"
      End
      Begin VB.Menu Gate_Pass_Inward 
         Caption         =   "Gate Pass Inward"
         Visible         =   0   'False
      End
      Begin VB.Menu Inspection 
         Caption         =   "Inspection"
         Visible         =   0   'False
      End
      Begin VB.Menu line16 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Good_Receive_Note 
         Caption         =   "Purchase(Good Receive Note)"
         Shortcut        =   ^P
      End
      Begin VB.Menu Good_Return_Note 
         Caption         =   "Purchase (Good Return Note)"
         Shortcut        =   ^R
      End
      Begin VB.Menu Inventory_Conversion 
         Caption         =   "Inventory Conversion"
      End
      Begin VB.Menu line4 
         Caption         =   "-"
      End
      Begin VB.Menu Credit_Sale 
         Caption         =   "Credit Sale"
      End
      Begin VB.Menu line121 
         Caption         =   "-"
      End
      Begin VB.Menu Payment_To_Vendor 
         Caption         =   "Payment To Vendor Bank"
         Shortcut        =   ^B
      End
      Begin VB.Menu Payment_To_Vendor_Cash 
         Caption         =   "Payment To Vendor Cash"
         Shortcut        =   ^C
      End
      Begin VB.Menu line17 
         Caption         =   "-"
      End
      Begin VB.Menu Post_Scan_Documents 
         Caption         =   "Post Scan Documents"
      End
      Begin VB.Menu line23 
         Caption         =   "-"
      End
      Begin VB.Menu Post_Purchase 
         Caption         =   "Post GRN/GL Voucher"
      End
      Begin VB.Menu Post_GRN_Return_GLVoucher 
         Caption         =   "Post GRN Return/GL Voucher"
      End
      Begin VB.Menu Post_Credit_Sale_GL_Voucher 
         Caption         =   "Post Credit Sale  GL Voucher"
      End
      Begin VB.Menu line12 
         Caption         =   "-"
      End
      Begin VB.Menu PVPGLVoucher 
         Caption         =   "Post Payable GL Voucher"
      End
      Begin VB.Menu Post_Payable_Cash_GL_Voucher 
         Caption         =   "Post Payable Cash GL Voucher"
      End
      Begin VB.Menu line5 
         Caption         =   "-"
      End
      Begin VB.Menu Unpost_Purchase_Invoice 
         Caption         =   "Unpost Purchase Status"
      End
      Begin VB.Menu Unpost_Vendor_Payments 
         Caption         =   "Unpost Vendor Payments"
         Visible         =   0   'False
      End
      Begin VB.Menu line9 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Import_Data 
         Caption         =   "Import Data"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu KimAr_Rpt 
      Caption         =   "Reports"
      Begin VB.Menu Print_Purchase_Order_Invoice 
         Caption         =   "Print Notes"
      End
      Begin VB.Menu line45 
         Caption         =   "-"
      End
      Begin VB.Menu Print_Credit_Sale_Note 
         Caption         =   "Print Credit Sale Note"
      End
      Begin VB.Menu Print_Credit_Sale_Register 
         Caption         =   "Print Credit Sale Register"
      End
      Begin VB.Menu loion 
         Caption         =   "-"
      End
      Begin VB.Menu Print_Notes_Register 
         Caption         =   "Print Notes Register"
      End
      Begin VB.Menu Print_Rate_Comparsion_Register 
         Caption         =   "Print Rate Comparsion Register"
      End
      Begin VB.Menu line6 
         Caption         =   "-"
      End
      Begin VB.Menu Purchase_Report 
         Caption         =   "Purchase Report"
      End
      Begin VB.Menu line62 
         Caption         =   "-"
      End
      Begin VB.Menu Print_BarCode 
         Caption         =   "Print BarCode"
      End
      Begin VB.Menu Print_Bar_Code1 
         Caption         =   "Print BarCode1"
      End
      Begin VB.Menu line7 
         Caption         =   "-"
      End
      Begin VB.Menu Stock_Ledger 
         Caption         =   "Stock Ledger"
         Shortcut        =   ^S
      End
      Begin VB.Menu Stock_Ledger_Periodic 
         Caption         =   "Stock Ledger Periodic"
         Visible         =   0   'False
      End
      Begin VB.Menu line41 
         Caption         =   "-"
      End
      Begin VB.Menu Vendor_Account_History 
         Caption         =   "Vendors Account History"
      End
      Begin VB.Menu Vendors_Payments 
         Caption         =   "Vendors Payments"
      End
      Begin VB.Menu Supplier_WhTax_Payments 
         Caption         =   "Supplier Withholding Tax Payments"
      End
      Begin VB.Menu line11 
         Caption         =   "-"
      End
      Begin VB.Menu Listings 
         Caption         =   "Listings"
         Begin VB.Menu SuppliersReport 
            Caption         =   "Suppliers Code Wise"
         End
         Begin VB.Menu SuppliersNW 
            Caption         =   "Suppliers Name Wise"
         End
         Begin VB.Menu Items 
            Caption         =   "Items"
         End
      End
      Begin VB.Menu line29 
         Caption         =   "-"
      End
      Begin VB.Menu Item_Wise_Purchase_Sale_Report 
         Caption         =   "Item Wise Purchase Report"
      End
      Begin VB.Menu Sale_Purchase_Report_Supplier 
         Caption         =   "Sale Purchase Report Supplier"
      End
   End
   Begin VB.Menu KimAr_Ret 
      Caption         =   "Return"
   End
End
Attribute VB_Name = "MDIPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Authority_Person_Setup_Click()
frmPOAuthorityPerson.Show
End Sub

Private Sub Bank_Setup_Click()
frmHRMBank.Show
End Sub

Public Sub City_Setup_Click()
frmcities.Show
End Sub

Private Sub Country_Setup_Click()
frmCountry.Show
End Sub

Private Sub Credit_Sale_Click()
frmSO_PosformCredit.Show
End Sub

Public Sub Demand_Note_Click()
frmPODemandNote.Show
End Sub

Private Sub Form_Load()
MDIForm1.Toolbar2.Visible = False
SuppliersSetup.Enabled = chkRights1("PUR0000001")
MDIForm1.Toolbar4.Buttons(1).Enabled = SuppliersSetup.Enabled
Authority_Person_Setup.Enabled = chkRights1("PUR0000002")
Vendor_Account_Setup.Enabled = chkRights1("PUR0000003")
Vendor_Balance_Sheet_Setup.Enabled = chkRights1("PUR0000004")
Other_Account_Setup.Enabled = chkRights1("PUR0000005")
Sector_Setup.Enabled = chkRights1("PUR0000006")
Country_Setup.Enabled = chkRights1("PUR0000007")
City_Setup.Enabled = chkRights1("PUR0000008")
Tehseel_Setup.Enabled = chkRights1("PUR0000009")
Bank_Setup.Enabled = chkRights1("PUR0000010")
Purchase_Order.Enabled = chkRights1("PUR0000011")
Good_Receive_Note.Enabled = chkRights1("PUR0000012")
MDIForm1.Toolbar4.Buttons(6).Enabled = Good_Receive_Note.Enabled
Good_Return_Note.Enabled = chkRights1("PUR0000013")

MDIForm1.Toolbar4.Buttons(7).Enabled = Good_Return_Note.Enabled

Inventory_Conversion.Enabled = chkRights1("PUR0000014")
Credit_Sale.Enabled = chkRights1("PUR0000015")

Payment_To_Vendor.Enabled = chkRights1("PUR0000016")
Payment_To_Vendor_Cash.Enabled = chkRights1("PUR0000017")
MDIForm1.Toolbar4.Buttons(8).Enabled = Payment_To_Vendor.Enabled
MDIForm1.Toolbar4.Buttons(9).Enabled = Payment_To_Vendor_Cash.Enabled
Post_Scan_Documents.Enabled = chkRights1("PUR0000018")
Post_Purchase.Enabled = chkRights1("PUR0000019")
Post_GRN_Return_GLVoucher.Enabled = chkRights1("PUR0000020")
Post_Purchase.Enabled = chkRights1("PUR0000021")
PVPGLVoucher.Enabled = chkRights1("PUR0000022")
Unpost_Purchase_Invoice.Enabled = chkRights1("PUR0000023")
Print_Purchase_Order_Invoice.Enabled = chkRights1("PUR0000024")
Print_Credit_Sale_Note.Enabled = chkRights1("PUR0000025")
Print_Notes_Register.Enabled = chkRights1("PUR0000026")
Print_Rate_Comparsion_Register.Enabled = chkRights1("PUR0000027")
Purchase_Report.Enabled = chkRights1("PUR0000028")
Print_BarCode.Enabled = chkRights1("PUR0000029")
Print_Bar_Code1.Enabled = chkRights1("PUR0000030")
Stock_Ledger.Enabled = chkRights1("PUR0000031")
MDIForm1.Toolbar4.Buttons(10).Enabled = Stock_Ledger.Enabled
Vendor_Account_History.Enabled = chkRights1("PUR0000032")
Vendors_Payments.Enabled = chkRights1("PUR0000033")
Supplier_WhTax_Payments.Enabled = chkRights1("PUR0000034")
SuppliersReport.Enabled = chkRights1("PUR0000035")
Items.Enabled = chkRights1("PUR0000036")
Item_Wise_Purchase_Sale_Report.Enabled = chkRights1("PUR0000037")
Sale_Purchase_Report_Supplier.Enabled = chkRights1("PUR0000038")
Post_Payable_Cash_GL_Voucher.Enabled = chkRights1("PUR0000039")


End Sub

Public Sub Gate_Pass_Inward_Click()
frmPOGatePassInward.Show
End Sub

Public Sub Good_Receive_Note_Click()
frmPOGoodsReceiveNote.Show
frmPOGoodsReceiveNote.NewRecord_Click
End Sub

Public Sub Good_Return_Note_Click()
frmPOGoodsReturnNote.Show
End Sub

Private Sub Import_Data_Click()
frmImportDataPurchase.Show
End Sub

Public Sub Inspection_Click()
frmPOInspection.Show
End Sub

Private Sub Inventory_Conversion_Click()
frmInventoryConversion.Show
End Sub

Private Sub Item_Wise_Purchase_Sale_Report_Click()
frmSoSalePurchaseReport.txtpstype.ListIndex = 2
frmSoSalePurchaseReport.Show
End Sub

Private Sub Items_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
 
   With rptListigns
        .ReportFileName = App.Path & Gs_ICRepoPath & "\IC_Items.RPT"
        .WindowTitle = "Company Items"
        .SelectionFormula = "{Ic_item.Compcode} = '" & Gs_compcode & "'"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = 'Company Items'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
    End With
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Public Sub KimAr_Ret_Click()
   MDIForm1.Caption = "E-Counts 2.0"
   MDIForm1.StatusBar1.Panels(2).Text = ""
   MDIForm1.StatusBar1.Panels(4).Text = ""
   MDIForm1.StatusBar1.Panels(6).Text = ""
   Para_Rs.Filter = adFilterNone
   MDIForm1.Toolbar4.Visible = False
   MDIForm1.Toolbar2.Visible = True
   Unload Me
End Sub

Private Sub Other_Account_Setup_Click()
frmPOPurchaseAccountSetup.Show
End Sub

Public Sub Payment_To_Vendor_Cash_Click()
frmPOVendorPaymentsCash.Show
End Sub

Public Sub Payment_To_vendor_Click()
frmPOVendorPaymentsBank.Show
End Sub

Private Sub Pending_Items_History_Click()
frmPOPendingVitems.Caption = Pending_Items_History.Caption
frmPOPendingVitems.Label3.Caption = "Item ID :"
frmPOPendingVitems.Show
End Sub

Private Sub Post_Credit_Sale_GL_Voucher_Click()
frmSOPostSaleFTDate.Caption = "Post Credit Sale GL Voucher"
frmSOPostSaleFTDate.Show
End Sub

Private Sub Post_GRN_Return_GLVoucher_Click()
frmPOPostPurchase.Caption = "Post GRRN"
frmPOPostPurchase.Label1.Caption = "Post GRRN"
frmPOPostPurchase.Label3.Caption = "GRRN # :"
frmPOPostPurchase.Command1.Caption = "Post GRRN"
frmPOPostPurchase.Show
End Sub

Private Sub Post_Payable_Cash_GL_Voucher_Click()
frmPOPostPurchase.Caption = "Post Payable GL Voucher"
frmPOPostPurchase.Label1.Caption = "Post Payable Cash"
frmPOPostPurchase.Label3.Caption = "Receipt # :"
frmPOPostPurchase.Command1.Caption = "Post Payments Cash"
frmPOPostPurchase.Command3.Visible = False
frmPOPostPurchase.Show
End Sub

Private Sub Post_Purchase_Click()
frmPOPostPurchase.Caption = "Post GRN"
frmPOPostPurchase.Label1.Caption = "Post GRN"
frmPOPostPurchase.Label3.Caption = "GRN # :"
frmPOPostPurchase.Command1.Caption = "Post GRN"
frmPOPostPurchase.Show
End Sub


Private Sub Post_Scan_Documents_Click()
Dim pr_dumy As New Recordset
pr_dumy.Open "select * from PO_Path ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    Gs_POScanDoc = Trim(pr_dumy("PathName") & "")
End If
pr_dumy.Close

If Gs_POScanDoc <> "" Then
    frmPOScanDoc.Show
Else
    Call MsgBox("Path not found for scan documents", vbCritical)
    Exit Sub
End If

End Sub

Private Sub Print_Bar_Code1_Click()
frmicreport9backup.Show
End Sub

Private Sub Print_BarCode_Click()
frmicreport9.Show
End Sub

Private Sub Print_Credit_Sale_Note_Click()
frmSOInvoice.Show
frmSOInvoice.txtnoteType.ListIndex = 1
frmSOInvoice.txtnoteType.Enabled = False
End Sub

Private Sub Print_Credit_Sale_Register_Click()
frmSoSaleCustomerBaseReport.Caption = Print_Credit_Sale_Register.Caption
frmSoSaleCustomerBaseReport.Show
End Sub

Private Sub Print_Notes_Register_Click()
frmPONoteRegisterReport.Show
End Sub

Private Sub Print_Purchase_Order_Invoice_Click()
frmPONoteReport.Show
End Sub

Private Sub Print_Rate_Comparsion_Register_Click()
frmPONoteRegisterReport.Caption = "Print Rate Comparsion Register"
frmPONoteRegisterReport.Show
End Sub

Public Sub Purchase_Order_Click()
frmPOPurchaseOrder.Show
End Sub

Private Sub Purchase_Register_Click()
frmicreport.Caption = "Purchase Register Report"
frmicreport.Show
End Sub


Private Sub Purchase_Return_Register_Click()
frmicreport.Caption = Purchase_Return_Register.Caption
frmicreport.Show
End Sub

Private Sub Purchase_Report_Click()
frmPOPurchaseReport.Show
End Sub

Private Sub PVPGLVoucher_Click()
frmPOPostPurchase.Caption = "Post Payable GL Voucher"
frmPOPostPurchase.Label1.Caption = "Post Payable"
frmPOPostPurchase.Label3.Caption = "Receipt # :"
frmPOPostPurchase.Command1.Caption = "Post Payments"
frmPOPostPurchase.Command3.Visible = False
frmPOPostPurchase.Show

End Sub

Private Sub Return_main_Click()
KimAr_Ret_Click
End Sub

Private Sub Sale_Purchase_Report_Supplier_Click()
frmSoSalePurchaseReport.txtpstype.ListIndex = 6
frmSoSalePurchaseReport.Show

End Sub

Public Sub Sector_Setup_Click()
frmSector.Show
End Sub

Private Sub setactiveyear_Click()
FrmFnYear.Show
End Sub


Private Sub Stock_Ledger_Click()
frmicreport4.Show
End Sub

Private Sub Stock_Ledger_Periodic_Click()
frmicreport4.Show
End Sub

Private Sub Supplier_WhTax_Payments_Click()
frmicreport3.Caption = "Supplier Withholding Tax Payments"
frmicreport3.Label2.Caption = "Vendor Code :"
frmicreport3.Frame2.Visible = False
frmicreport3.Show

End Sub

Private Sub SuppliersNW_Click()
MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
 
   With rptListigns
        .ReportFileName = App.Path & Gs_ICRepoPath & "\IC_SuppliersNameWise.RPT"
        .WindowTitle = "Company Suppliers"
        .SelectionFormula = "{IC_Supplier.Compcode} = '" & Gs_compcode & "'"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = 'Company Supplier'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
    End With
MDIForm1.StatusBar1.Panels(7).Text = ""
End Sub

Private Sub SuppliersReport_Click()
 MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
 
   With rptListigns
        .ReportFileName = App.Path & Gs_ICRepoPath & "\IC_Suppliers.RPT"
        .WindowTitle = "Company Suppliers"
        .SelectionFormula = "{IC_Supplier.Compcode} = '" & Gs_compcode & "'"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = 'Company Supplier'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
    End With
MDIForm1.StatusBar1.Panels(7).Text = ""

End Sub

Public Sub SuppliersSetup_Click()
frmSupplier.Show
End Sub
Public Sub Tehseel_Setup_Click()
frmtehseel.Show
End Sub
Private Sub Unpost_Purchase_Invoice_Click()
frmpasswordform.txtopt = 2
frmpasswordform.Show

End Sub

Private Sub Unpost_Vendor_Payments_Click()
frmPaymentsUnpost.Show
End Sub

Private Sub Vendor_Account_History_Click()
frmicreport1.Caption = "Vendor Account History"
frmicreport1.Label2.Caption = "Vendor Code :"
frmicreport1.Show
End Sub

Private Sub Vendor_Account_Setup_Click()
frmVendorAccountSetup.Show
End Sub

Private Sub Vendor_Balance_Sheet_Setup_Click()
frmBSAccountSetupVendor.Show
End Sub

Private Sub Vendors_Payments_Click()
frmicreport3.Caption = "Vendor Payments"
frmicreport3.Label2.Caption = "Vendor Code :"
frmicreport3.Frame2.Visible = True
frmicreport3.Show
End Sub

Private Sub Vendors_Pending_Items_History_Click()
frmPOPendingVitems.Caption = Vendors_Pending_Items_History.Caption
frmPOPendingVitems.Label3.Caption = "Vendors ID :"
frmPOPendingVitems.Show
End Sub

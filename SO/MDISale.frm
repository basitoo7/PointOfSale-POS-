VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form MDISale 
   BackColor       =   &H00808000&
   ClientHeight    =   3195
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport rptListigns 
      Left            =   300
      Top             =   60
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
         Caption         =   "Client Setup"
         Shortcut        =   ^C
      End
      Begin VB.Menu Authority_sale_Person_Setup 
         Caption         =   "Authority Sale Person Setup"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu Bank_Setup 
         Caption         =   "Bank Setup"
      End
      Begin VB.Menu Account_Setup 
         Caption         =   "Sale Type Setup"
         Visible         =   0   'False
      End
      Begin VB.Menu line2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Client_Account_Setup 
         Caption         =   "Client Account Setup"
         Visible         =   0   'False
      End
      Begin VB.Menu Client_Balance_Sheet_Setup 
         Caption         =   "Client Balance Sheet Setup"
         Visible         =   0   'False
      End
      Begin VB.Menu Cash_Account_Setup 
         Caption         =   "Sale Accounts Setup"
         Visible         =   0   'False
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu setactiveyear 
         Caption         =   "Set Active Financial Year"
      End
      Begin VB.Menu line4 
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
      Begin VB.Menu line21 
         Caption         =   "-"
      End
      Begin VB.Menu Return_Home 
         Caption         =   "Return"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu KimAr_Trns 
      Caption         =   "Transaction"
      Begin VB.Menu Pos_Sale_Invoice 
         Caption         =   "POS Sale Invoice"
         Shortcut        =   ^P
      End
      Begin VB.Menu line16 
         Caption         =   "-"
      End
      Begin VB.Menu Pos_Edit 
         Caption         =   "Sale Invoice"
         Visible         =   0   'False
      End
      Begin VB.Menu Sale_Return_Invoice 
         Caption         =   "Sale Return Invoice"
         Shortcut        =   ^R
      End
      Begin VB.Menu line25 
         Caption         =   "-"
      End
      Begin VB.Menu Post_Sale 
         Caption         =   "Post Sale"
         Visible         =   0   'False
      End
      Begin VB.Menu Post_Sale_GL_Voucher 
         Caption         =   "Post Sale GL Voucher"
      End
      Begin VB.Menu Post_Sale_Return_GL_Voucher 
         Caption         =   "Post Sale Return GL Voucher"
      End
      Begin VB.Menu Post_Credit_Sale_GL_Voucher 
         Caption         =   "Post Credit Sale  GL Voucher"
      End
      Begin VB.Menu Post_Cost_of_Sale_Voucher 
         Caption         =   "Post Cost of Sale Voucher"
      End
      Begin VB.Menu line7 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Unpost_Sale_Invoice 
         Caption         =   "Unpost Sale Invoice"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu KimAr_Rpt 
      Caption         =   "Reports"
      Begin VB.Menu Print_Sale_Invoice 
         Caption         =   "Print Sale Invoice"
         Shortcut        =   ^I
      End
      Begin VB.Menu line201 
         Caption         =   "-"
      End
      Begin VB.Menu Sale_Register 
         Caption         =   "Sale Report With Department Filter "
         Shortcut        =   ^D
      End
      Begin VB.Menu Sale_Register_Manufacturer_Wise 
         Caption         =   "Sale Report With Supplier Filter"
         Shortcut        =   ^M
      End
      Begin VB.Menu Sale_Purchase_Manufacturer_Wise 
         Caption         =   "In-Out History Manufacturer Wise"
         Shortcut        =   ^G
         Visible         =   0   'False
      End
      Begin VB.Menu Print_Casher_Sale_Report 
         Caption         =   "Sale Report With Casher Filter"
      End
      Begin VB.Menu line50 
         Caption         =   "-"
      End
      Begin VB.Menu Print_Discount_Report 
         Caption         =   "Discount Report With Date Filter"
      End
      Begin VB.Menu line12 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Sale_Report_Category_Wise_With_Cost 
         Caption         =   "Sale Report Category Wise With Cost"
         Visible         =   0   'False
      End
      Begin VB.Menu Sale_Report_Item_Wise_With_Cost 
         Caption         =   "Sale Report Item Wise With Cost"
         Visible         =   0   'False
      End
      Begin VB.Menu Sale_Report_Item_History_With_Cost 
         Caption         =   "Sale Report Item History With Cost"
         Visible         =   0   'False
      End
      Begin VB.Menu line41 
         Caption         =   "-"
      End
      Begin VB.Menu Sale_Return_Report 
         Caption         =   "Sale Return Report"
      End
      Begin VB.Menu line11 
         Caption         =   "-"
      End
      Begin VB.Menu Item_Wise_Sale_Report 
         Caption         =   "Item Wise Sale Report"
      End
      Begin VB.Menu Customer_Wise_Sale_Report 
         Caption         =   "Customer Wise Sale Report"
      End
      Begin VB.Menu line20 
         Caption         =   "-"
      End
      Begin VB.Menu Stock_Ledger 
         Caption         =   "Stock Ledger"
      End
      Begin VB.Menu line33 
         Caption         =   "-"
      End
      Begin VB.Menu WasRep 
         Caption         =   "Waste Report"
      End
      Begin VB.Menu Lin1 
         Caption         =   "-"
      End
      Begin VB.Menu Listings 
         Caption         =   "Listings"
         Begin VB.Menu Client_List 
            Caption         =   "Clients"
         End
         Begin VB.Menu Item_List 
            Caption         =   "Items"
         End
      End
   End
   Begin VB.Menu KimAr_Ret 
      Caption         =   "Return"
   End
End
Attribute VB_Name = "MDISale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Account_Setup_Click()
frmARSaleType.Show
End Sub

Private Sub Authority_sale_Person_Setup_Click()
frmSOAuthoritySalePerson.Show
End Sub

Private Sub Bank_Setup_Click()
frmHRMBank.Show
End Sub

Private Sub Cash_Account_Setup_Click()
frmSOAccountSetup.Show
End Sub

Private Sub City_Setup_Click()
frmcities.Show
End Sub

Private Sub Client_List_Click()
 MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
 
   With rptListigns
        .ReportFileName = App.Path & Gs_ICRepoPath & "\IC_Parties.RPT"
        .WindowTitle = "Company Clients"
        .SelectionFormula = "{IC_Supplier.Compcode} = '" & Gs_compcode & "'"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = 'Company Clients'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
    End With
MDIForm1.StatusBar1.Panels(7).Text = ""

End Sub
Private Sub Country_Setup_Click()
frmCountry.Show
End Sub

Private Sub Customer_Wise_Sale_Report_Click()
frmSoSaleCustomerBaseReport.Show
End Sub

Private Sub Form_Load()

SuppliersSetup.Enabled = chkRights1("SALE000001")
MDIForm1.Toolbar5.Buttons(1).Enabled = SuppliersSetup.Enabled
Authority_sale_Person_Setup.Enabled = chkRights1("SALE000002")
Bank_Setup.Enabled = chkRights1("SALE000003")
Sector_Setup.Enabled = chkRights1("SALE000004")
Country_Setup.Enabled = chkRights1("SALE000005")
City_Setup.Enabled = chkRights1("SALE000006")
Tehseel_Setup.Enabled = chkRights1("SALE000007")
Pos_Sale_Invoice.Enabled = chkRights1("SALE000008")
MDIForm1.Toolbar5.Buttons(2).Enabled = Pos_Sale_Invoice.Enabled
Sale_Return_Invoice.Enabled = chkRights1("SALE000009")
MDIForm1.Toolbar5.Buttons(3).Enabled = Sale_Return_Invoice.Enabled
Post_Sale_GL_Voucher.Enabled = chkRights1("SALE000010")
Post_Sale_Return_GL_Voucher.Enabled = chkRights1("SALE000011")
Print_Sale_Invoice.Enabled = chkRights1("SALE000012")
Sale_Register.Enabled = chkRights1("SALE000013")
MDIForm1.Toolbar5.Buttons(4).Enabled = Sale_Register.Enabled

Sale_Register_Manufacturer_Wise.Enabled = chkRights1("SALE000014")
Print_Casher_Sale_Report.Enabled = chkRights1("SALE000015")
Print_Discount_Report.Enabled = chkRights1("SALE000016")
Sale_Return_Report.Enabled = chkRights1("SALE000017")
Item_Wise_Sale_Report.Enabled = chkRights1("SALE000018")
Customer_Wise_Sale_Report.Enabled = chkRights1("SALE000019")
Client_List.Enabled = chkRights1("SALE000020")
Item_List.Enabled = chkRights1("SALE000021")
Stock_Ledger.Enabled = chkRights1("SALE000022")

MDIForm1.Toolbar5.Buttons(5).Enabled = Stock_Ledger.Enabled


MDIForm1.Toolbar2.Visible = False

End Sub
Private Sub Item_List_Click()
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

Private Sub Item_Wise_Sale_Report_Click()
frmicreportCasher.Caption = "Sale Report"
frmicreportCasher.ChkSummary.Value = 1
frmicreportCasher.txtreporttye.ListIndex = 4
frmicreportCasher.Show

End Sub

Public Sub Pos_Edit_Click()
frmSO_PosformEdit.Show
End Sub
Public Sub KimAr_Ret_Click()
   MDIForm1.Caption = "E-Counts 2.0"
   MDIForm1.StatusBar1.Panels(2).Text = ""
   MDIForm1.StatusBar1.Panels(4).Text = ""
   MDIForm1.StatusBar1.Panels(6).Text = ""
   Para_Rs.Filter = adFilterNone
   MDIForm1.Toolbar5.Visible = False
   MDIForm1.Toolbar2.Visible = True
   If ln_posaccess = 1 Or ln_posaccess = 2 Or ln_posaccess = 3 Then
   End
   Else
   Unload Me
   End If
End Sub

Public Sub Pos_Sale_Invoice_Click()
 frmSO_Posform.Show
End Sub

Private Sub Post_Cost_of_Sale_Voucher_Click()
frmSOPostSaleFTDate.Caption = "Post Cost Of Sale"
frmSOPostSaleFTDate.Command1.Caption = "Post Cost Sale"
frmSOPostSaleFTDate.Show
End Sub

Private Sub Post_Credit_Sale_GL_Voucher_Click()
frmSOPostSaleFTDate.Caption = "Post Credit Sale GL Voucher"
frmSOPostSaleFTDate.Show
End Sub

Private Sub Post_Sale_Click()
frmSOPostSale.Caption = "Post Sale"
frmSOPostSale.Label1.Caption = "Post Sale"
frmSOPostSale.Command1.Caption = "Post Sale"
frmSOPostSale.Show
End Sub

Private Sub Post_Sale_GL_Voucher_Click()
frmSOPostSaleFTDate.Caption = "Post Sale GL Voucher"
frmSOPostSaleFTDate.Show
End Sub

Private Sub Post_Sale_Return_GL_Voucher_Click()
frmSOPostSaleFTDate.Caption = "Post Sale Return GL Voucher"
frmSOPostSaleFTDate.Show
End Sub

Private Sub Print_Casher_Sale_Report_Click()
frmicreportCasher.Caption = "Sale Report"
frmicreportCasher.ChkSummary.Value = 1
frmicreportCasher.Show
End Sub

Private Sub Print_Discount_Report_Click()
frmDiscountReport.Show
End Sub

Private Sub Print_Sale_Invoice_Click()
frmSOInvoice.Show
End Sub

Private Sub Return_Home_Click()
KimAr_Ret_Click
End Sub

Private Sub Sale_Purchase_Manufacturer_Wise_Click()
frmicreport8.Caption = Sale_Purchase_Manufacturer_Wise.Caption
frmicreport8.Show
End Sub

Public Sub Sale_Register_Click()
frmicreport.Caption = "Sale Report Deparment Wise"
frmicreport.Show
End Sub

Public Sub Sale_Register_Manufacturer_Wise_Click()
frmicreport8.Caption = "Sale Report With Supplier"
frmicreport8.Show
End Sub

Private Sub Sale_Report_Category_Wise_With_Cost_Click()
frmicreportCasher.Caption = Sale_Report_Category_Wise_With_Cost.Caption
frmicreportCasher.Show
End Sub

Private Sub Sale_Report_Item_History_With_Cost_Click()
frmSoSalePurchaseReport.Caption = Sale_Report_Item_History_With_Cost.Caption
frmSoSalePurchaseReport.Show
End Sub

Private Sub Sale_Report_Item_Wise_With_Cost_Click()
frmicreportCasher.Caption = Sale_Report_Item_Wise_With_Cost.Caption
frmicreportCasher.Show
End Sub

Public Sub Sale_Return_Invoice_Click()
frmSO_PosformReturn.Show
End Sub

Private Sub Sale_Return_Report_Click()
frmicreportReturn.Show
End Sub

Private Sub Sector_Setup_Click()
frmSector.Show
End Sub
Private Sub setactiveyear_Click()
FrmFnYear.Show
End Sub

Public Sub Stock_Ledger_Click()
frmicreport4.Show
End Sub

Public Sub SuppliersSetup_Click()
frmClients.Show
End Sub
Private Sub Tehseel_Setup_Click()
frmtehseel.Show
End Sub
Private Sub Unpost_Sale_Invoice_Click()
frmInvoiceUnpost.Show
End Sub

Private Sub WasRep_Click()
frmicreportCasher.Caption = "Waste Report"
frmicreportCasher.ChkSummary.Value = 1
frmicreportCasher.Show
End Sub

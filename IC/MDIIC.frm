VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form MDIIC 
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
   Icon            =   "MDIIC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport rptListigns 
      Left            =   0
      Top             =   -15
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
      Begin VB.Menu Item_Setup 
         Caption         =   "Items Setup"
         Shortcut        =   ^I
      End
      Begin VB.Menu Item_Sites_Setup 
         Caption         =   "Item Sites Setup"
         Visible         =   0   'False
      End
      Begin VB.Menu Item_Bins_Setup 
         Caption         =   "Item Bins Setup"
         Visible         =   0   'False
      End
      Begin VB.Menu line14 
         Caption         =   "-"
      End
      Begin VB.Menu Item_Category_Setup 
         Caption         =   "Departments Setup"
      End
      Begin VB.Menu Manufacturer_Setup 
         Caption         =   "Suppliers Setup"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu Item_Class_Setup 
         Caption         =   "Category Setup "
      End
      Begin VB.Menu Sub_Category_Setup 
         Caption         =   "Sub Category Setup "
      End
      Begin VB.Menu Measurement_Unit_Setup 
         Caption         =   "Measurement Unit Setup"
      End
      Begin VB.Menu line4 
         Caption         =   "-"
      End
      Begin VB.Menu Tax_Schedule_Setup 
         Caption         =   "Tax Schedule Setup"
      End
      Begin VB.Menu Department_Setup 
         Caption         =   "Department Setup"
         Visible         =   0   'False
      End
      Begin VB.Menu line10 
         Caption         =   "-"
      End
      Begin VB.Menu setactiveyear 
         Caption         =   "Set Active Financial Year"
      End
      Begin VB.Menu line24 
         Caption         =   "-"
      End
      Begin VB.Menu Stock_Zero_Form 
         Caption         =   "Stock Zero Form"
      End
      Begin VB.Menu line33 
         Caption         =   "-"
      End
      Begin VB.Menu Set_Sale_Cost 
         Caption         =   "Set Sale Cost"
         Visible         =   0   'False
      End
      Begin VB.Menu Return_Home 
         Caption         =   "Return"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu KimAr_Trns 
      Caption         =   "Transaction"
      Begin VB.Menu Inventory_Issue 
         Caption         =   "Inventory Issue"
         Shortcut        =   ^U
      End
      Begin VB.Menu Inventory_Issue_Return 
         Caption         =   "Inventory Issue Return"
         Shortcut        =   ^R
      End
      Begin VB.Menu Inventory_Transfer 
         Caption         =   "Inventory Transfer"
         Visible         =   0   'False
      End
      Begin VB.Menu Inventory_Waste 
         Caption         =   "Inventory Waste"
      End
      Begin VB.Menu Inventory_Adjustment 
         Caption         =   "Inventory Out Adjustment"
         Shortcut        =   ^A
      End
      Begin VB.Menu Inventory_In_Adjustment 
         Caption         =   "Inventory In Adjustment"
      End
      Begin VB.Menu line25 
         Caption         =   "-"
      End
      Begin VB.Menu Post_Consumption_Voucher 
         Caption         =   "Post Consumption Voucher"
         Visible         =   0   'False
      End
      Begin VB.Menu Post_Issue_Return_Voucher 
         Caption         =   "Post Issue Return Voucher"
         Visible         =   0   'False
      End
      Begin VB.Menu Post_Adjustmnet_Voucher 
         Caption         =   "Post Adjustmnet Voucher"
      End
   End
   Begin VB.Menu KimAr_Rpt 
      Caption         =   "Reports"
      Begin VB.Menu Print_Note 
         Caption         =   "Print Note"
      End
      Begin VB.Menu Print_Note_Register 
         Caption         =   "Print Note Register"
      End
      Begin VB.Menu List_of_Items 
         Caption         =   "List of Items"
      End
      Begin VB.Menu line23 
         Caption         =   "-"
      End
      Begin VB.Menu List_of_Items_COR 
         Caption         =   "Sale List Rate Change Periodic"
      End
      Begin VB.Menu Purchase_List_Rate_Change_Periodic 
         Caption         =   "Purchase List Rate Change Periodic"
      End
      Begin VB.Menu Saleregister 
         Caption         =   "Purchase Receipts Report"
         Visible         =   0   'False
      End
      Begin VB.Menu line13 
         Caption         =   "-"
      End
      Begin VB.Menu Stock_Ledger_Balance 
         Caption         =   "Stock Ledger Balance"
         Visible         =   0   'False
      End
      Begin VB.Menu Stock_Ledger 
         Caption         =   "Stock Ledger Summary"
         Visible         =   0   'False
      End
      Begin VB.Menu Stock_Ledger_Detail 
         Caption         =   "Stock Ledger Detail"
         Shortcut        =   ^S
         Visible         =   0   'False
      End
      Begin VB.Menu Stock_Ledger_Periodic 
         Caption         =   "Stock Ledger"
      End
   End
   Begin VB.Menu KimAr_Ret 
      Caption         =   "Return"
   End
End
Attribute VB_Name = "MDIIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Department_Setup_Click()
frmDepartments.Show
End Sub

Private Sub Form_Load()
MDIForm1.Toolbar2.Visible = False

Item_Setup.Enabled = chkRights1("INVTR00001")
MDIForm1.Toolbar1.Buttons(1).Enabled = Item_Setup.Enabled

Item_Category_Setup.Enabled = chkRights1("INVTR00002")
Manufacturer_Setup.Enabled = chkRights1("INVTR00003")
Item_Class_Setup.Enabled = chkRights1("INVTR00004")
Sub_Category_Setup.Enabled = chkRights1("INVTR00005")
Measurement_Unit_Setup.Enabled = chkRights1("INVTR00006")
Tax_Schedule_Setup.Enabled = chkRights1("INVTR00007")
Inventory_Issue.Enabled = chkRights1("INVTR00008")
MDIForm1.Toolbar1.Buttons(2).Enabled = Inventory_Issue.Enabled
Inventory_Issue_Return.Enabled = chkRights1("INVTR00009")
MDIForm1.Toolbar1.Buttons(3).Enabled = Inventory_Issue_Return.Enabled
Inventory_Transfer.Enabled = chkRights1("INVTR00010")
Inventory_Waste.Enabled = chkRights1("INVTR00011")
Inventory_Adjustment.Enabled = chkRights1("INVTR00012")
MDIForm1.Toolbar1.Buttons(5).Enabled = Inventory_Adjustment.Enabled
Inventory_In_Adjustment.Enabled = chkRights1("INVTR00013")
Post_Adjustmnet_Voucher.Enabled = chkRights1("INVTR00014")
Print_Note.Enabled = chkRights1("INVTR00015")
Print_Note_Register.Enabled = chkRights1("INVTR00016")
List_of_Items.Enabled = chkRights1("INVTR00017")
List_of_Items_COR.Enabled = chkRights1("INVTR00018")
Purchase_List_Rate_Change_Periodic.Enabled = chkRights1("INVTR00019")
Stock_Ledger_Periodic.Enabled = chkRights1("INVTR00020")
MDIForm1.Toolbar1.Buttons(6).Enabled = Stock_Ledger_Periodic.Enabled



End Sub

Private Sub Inventory_In_Adjustment_Click()
FrmInventoryInAdjustment.Show
End Sub

Public Sub Inventory_Issue_Click()
frmInventoryIssue.Show
End Sub

Private Sub Inventory_Issue_Return_Click()
frmInventoryIssueReturn.Show
End Sub

Public Sub Inventory_transfer_Click()
frmInventoryTransfer.Show
End Sub
Public Sub Inventory_adjustment_Click()
frmInventoryAdjustment.Show
End Sub

Private Sub Inventory_Waste_Click()
FrmInventoryWaste.Show
End Sub

Private Sub Item_Bins_Setup_Click()
frmSitesBins.Show
End Sub

Private Sub Item_Category_Setup_Click()
frmpasswordform.Show
End Sub

Private Sub Item_Class_Setup_Click()
frmItemClass.Show
End Sub

Private Sub Item_Setup_Click()
frmItemstp.Show
End Sub

Private Sub Item_Sites_Setup_Click()
frmSites.Show
End Sub
Private Sub ItemRecipeformulaSetup_Click()
frmRecipeformula.Show
End Sub
Public Sub KimAr_Ret_Click()
   MDIForm1.Caption = "E-Counts 2.0"
   MDIForm1.StatusBar1.Panels(2).Text = ""
   MDIForm1.StatusBar1.Panels(4).Text = ""
   MDIForm1.StatusBar1.Panels(6).Text = ""
   Para_Rs.Filter = adFilterNone
   MDIForm1.Toolbar1.Visible = False
   MDIForm1.Toolbar2.Visible = True
   ParaCntr_Rs.Close
   Unload Me
End Sub

Private Sub List_of_Items_Click()
frmicreport10.Caption = List_of_Items.Caption
frmicreport10.Show
End Sub

Private Sub List_of_Items_COR_Click()
frmicreport10.Caption = List_of_Items_COR.Caption
frmicreport10.Show
End Sub

Private Sub List_of_Items_for_Adjustment_Click()
frmicreport10.Caption = List_of_Items_for_Adjustment.Caption
frmicreport10.Show
End Sub

Private Sub Manufacturer_Setup_Click()
frmSupplier.Show
End Sub

Private Sub Material_Setup_Click()
frmItemMaterial.Show
End Sub

Private Sub Measurement_Unit_Setup_Click()
frmunits.Show
End Sub

Private Sub Post_Adjustmnet_Voucher_Click()
frmICPostInvntory.Caption = "Post Adjustment Voucher"
frmICPostInvntory.Label1.Caption = "Post Adjustment Note"
frmICPostInvntory.Label3.Caption = "Adjustment Note #:"
frmICPostInvntory.Command1.Caption = "Post Adjustment Note"
frmICPostInvntory.Show
End Sub

Private Sub Post_Consumption_Voucher_Click()
frmICPostInvntory.Caption = "Post Inventory Consumption Voucher"
frmICPostInvntory.Label1.Caption = "Post Issue Note"
frmICPostInvntory.Label3.Caption = "Issue Note #:"
frmICPostInvntory.Command1.Caption = "Post Issue Note"
frmICPostInvntory.Show
End Sub

Private Sub Post_Issue_Return_Voucher_Click()
frmICPostInvntory.Caption = "Post Issue Return Voucher"
frmICPostInvntory.Label1.Caption = "Post Issue Return Note"
frmICPostInvntory.Label3.Caption = "Issue Return Note #:"
frmICPostInvntory.Command1.Caption = "Post Issue Return Note"
frmICPostInvntory.Show
End Sub

Private Sub Print_Note_Click()
frmICNoteReport.Show
End Sub

Private Sub Print_Note_Register_Click()
frmICNoteRegisterReport.Show
End Sub
Private Sub Printinvoice_Click()

End Sub

Private Sub Purchase_List_Rate_Change_Periodic_Click()
frmicreport10.Caption = Purchase_List_Rate_Change_Periodic.Caption
frmicreport10.Show

End Sub

Private Sub Return_Home_Click()
KimAr_Ret_Click
End Sub

Private Sub Saleregister_Click()
frmicreport.Caption = "Sale Register Report"
frmicreport.codeid = "C"
frmicreport.Show
End Sub

Private Sub Set_Sale_Cost_Click()
frmavgRate.Show
End Sub

Private Sub setactiveyear_Click()
FrmFnYear.Show
End Sub

Private Sub Stock_Ledger_Balance_Click()
frmicreport7.Show
End Sub

Public Sub Stock_Ledger_Click()
frmicreport2.Show
End Sub

Private Sub Stock_Ledger_Detail_Click()
frmicreport4.Caption = "Stock Ledger Detail"
frmicreport4.Show
End Sub

Private Sub Stock_Ledger_Periodic_Click()
frmicreport4.Show
End Sub

Private Sub Stock_Zero_Form_Click()
If UCase(Gc_UserId) = UCase("Admin") Then
frmicreport12.Show
End If
End Sub

Private Sub Sub_Category_Setup_Click()
frmItemSubCategory.Show
End Sub

Private Sub Tax_Schedule_Setup_Click()
frmTaxSchedule.Show
End Sub

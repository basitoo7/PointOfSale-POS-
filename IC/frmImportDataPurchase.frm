VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmImportDataPurchase 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Data"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10995
   Icon            =   "frmImportDataPurchase.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   510
      Left            =   -45
      TabIndex        =   7
      Top             =   4935
      Width           =   10995
      Begin VB.Label lblStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   90
         TabIndex        =   8
         Top             =   180
         Width           =   3195
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   9780
      TabIndex        =   6
      Top             =   5475
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   345
      Left            =   8580
      TabIndex        =   5
      Top             =   5475
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   1080
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   10920
      Begin VB.CommandButton Command4 
         Caption         =   "Re-Load Import Data"
         Height          =   360
         Left            =   5145
         TabIndex        =   14
         Top             =   585
         Width           =   1800
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Import Data"
         Height          =   360
         Left            =   3270
         TabIndex        =   13
         Top             =   585
         Width           =   1800
      End
      Begin MSComCtl2.DTPicker DtpFrom 
         Height          =   345
         Left            =   1500
         TabIndex        =   1
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         Format          =   63176705
         CurrentDate     =   40600
      End
      Begin MSComCtl2.DTPicker DtpTo 
         Height          =   345
         Left            =   1500
         TabIndex        =   2
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         Format          =   63176705
         CurrentDate     =   40600
      End
      Begin Crystal.CrystalReport rptLedger 
         Left            =   5355
         Top             =   105
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
      Begin VB.Label Label2 
         Caption         =   "To Date :"
         Height          =   240
         Left            =   750
         TabIndex        =   4
         Top             =   645
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "From Date :"
         Height          =   240
         Left            =   615
         TabIndex        =   3
         Top             =   225
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3960
      Left            =   30
      TabIndex        =   9
      Top             =   975
      Width           =   4695
      Begin MSFlexGridLib.MSFlexGrid GrdGRN 
         Height          =   3780
         Left            =   45
         TabIndex        =   11
         Top             =   135
         Width           =   4590
         _ExtentX        =   8096
         _ExtentY        =   6668
         _Version        =   393216
         Rows            =   1
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3960
      Left            =   4755
      TabIndex        =   10
      Top             =   975
      Width           =   6195
      Begin MSFlexGridLib.MSFlexGrid Grid2 
         Height          =   3780
         Left            =   30
         TabIndex        =   12
         Top             =   135
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   6668
         _Version        =   393216
         Rows            =   1
      End
   End
End
Attribute VB_Name = "frmImportDataPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ls_sql As String

Private Sub ImportCashSale()
Dim PR_Dumy As New Recordset
Dim ln_mincnt As Double
Dim ln_maxcnt As Double
Dim ln_srno As Double
Dim ln_TotalAmount As Double
Dim ln_Amount As Double

Dim ls_transcode As String
lblStatus.Caption = "Cash Sale in progress..."
Me.Refresh

ls_sql = "DELETE FROM IC_TransAutoSale WHERE((Compcode + BranchCode + TransCode) IN"
ls_sql = ls_sql & "(SELECT     compcode + branchcode + transcode From IC_TransMasterAutoSale"
ls_sql = ls_sql & " WHERE      IC_TransMasterAutoSale.saletype = 003))"
gc_dbcon.Execute ls_sql
ls_sql = "delete from IC_TransMasterAutoSale where saletype = 003"
gc_dbcon.Execute ls_sql


ls_sql = "TRUNCATE TABLE IC_TransCashMaster"
gc_dbcon.Execute ls_sql
ls_sql = "TRUNCATE TABLE IC_TransCash"
gc_dbcon.Execute ls_sql


ls_sql = "insert into  IC_TransCashMaster(Compcode, BranchCode, TransCode, TransType, PONO, DCNO, TruckNo, LotNo, InvoiceNo, TransDate, AccountCode, SiteID, BinID, Remarks, SubTotal,"
ls_sql = ls_sql & " TradeDiscount, Freight, Miscellaneous, TaxAmount, TotalAmount, MiscRemarks,  Discount,   SaleType, SedAmount, RecAmount, UserID, Adddate, AddTime)"
ls_sql = ls_sql & " SELECT '001' AS Compcode, '001' AS BranchCode, SUBSTRING('0000000000', 1, 10 - LEN(RTRIM(LTRIM(SaleInvcode))))"
ls_sql = ls_sql & " + RTRIM(LTRIM(SaleInvcode)) AS Transcode, 'S' AS TransType, 'NA' AS PONO, 'NA' AS DCNO, 'NA' AS TruckNo, 'NA' AS LotNo,"
ls_sql = ls_sql & " '0000000001' AS InvoiceNo, [date] AS TransDate, SUBSTRING('000000', 1, 6 - LEN(RTRIM(LTRIM(CustCode))))+RTRIM(LTRIM(CustCode)) AS AccountCode, '001' AS SiteID, '001' AS BinID, 'Sale' AS Remarks, InvTotal AS SubTotal,"
ls_sql = ls_sql & " 0 AS TradeDiscount, 0 AS Freight, 0 AS Miscellaneous, 0 AS TaxAmount, InvTotal AS TotalAmount, 'Sale' AS MiscRemarks, FlatDisc AS Discount,"
ls_sql = ls_sql & " SUBSTRING('000', 1, 3 - LEN(RTRIM(LTRIM(SaleCatCode))))+RTRIM(LTRIM(SaleCatCode))  AS SaleType, 0 AS SedAmount, 0 AS RecAmount, 'admin' AS UserID, [date] AS Adddate, '08:00:00' AS AddTime"
ls_sql = ls_sql & " FROM RahatDeptStoresDataBaseV3.dbo.SaleLedger where [date] >='" & Format(DtpFrom, "YYYY/MM/DD") & "' and [date] <='" & Format(DtpTo, "YYYY/MM/DD") & "' and salecatcode = 3"
gc_dbcon.Execute ls_sql

ls_sql = "insert into  IC_TransCash (Compcode, BranchCode, TransCode, ItemCode, Quantity, ItemRate, Amount)"
ls_sql = ls_sql & " SELECT '001' AS Compcode, '001' AS BranchCode, SUBSTRING('0000000000', 1, 10 - LEN(RTRIM(LTRIM(Saledetail.SaleInvcode))))"
ls_sql = ls_sql & " + RTRIM(LTRIM(Saledetail.SaleInvcode)) AS Transcode, SUBSTRING('000000', 1, 6 - LEN(RTRIM(LTRIM(Saledetail.ICode))))"
ls_sql = ls_sql & " + RTRIM(LTRIM(Saledetail.ICode)) AS ItemCode, Saledetail.LooseQty, Saledetail.SalePrice as Rate ,Saledetail.LooseQty* Saledetail.SalePrice as Amount"
ls_sql = ls_sql & " FROM RahatDeptStoresDatabaseV3.dbo.SaleLedger  SaleLedger INNER JOIN"
ls_sql = ls_sql & " RahatDeptStoresDatabaseV3.dbo.Saledetail Saledetail ON SaleLedger.SaleInvCode = Saledetail.SaleInvcode"
ls_sql = ls_sql & " WHERE (SaleLedger.[date] >= '" & Format(DtpFrom, "YYYY/MM/DD") & "' and (SaleLedger.[date] <= '" & Format(DtpTo, "YYYY/MM/DD") & "') and SaleLedger.salecatcode = 3)"
gc_dbcon.Execute ls_sql

PR_Dumy.Open "Select max(srno) as MaxNo from IC_TransCashmaster", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not PR_Dumy.EOF Then
ln_maxcnt = Val(0 & PR_Dumy("MaxNo"))
End If
PR_Dumy.Close
ln_mincnt = 1
ln_TotalAmount = 0
Do While Not Val(txtamount) <= ln_TotalAmount
ln_srno = CInt(Int((Val(ln_maxcnt) * Rnd()) + ln_mincnt))

PR_Dumy.Open "Select sum(Totalamount) as TotalAmount from IC_TransCashmaster where srno = " & ln_srno & " ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not PR_Dumy.EOF Then
ln_Amount = Val(0 & PR_Dumy("TotalAmount"))
End If
PR_Dumy.Close

PR_Dumy.Open "Select Transcode  from IC_TransCashmaster where srno = " & ln_srno & " ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not PR_Dumy.EOF Then
ls_transcode = Trim(PR_Dumy("TransCode") & "")
End If
PR_Dumy.Close

PR_Dumy.Open "Select Transcode  from IC_Transmasterautosale where Transcode = " & ls_transcode & " ", gc_dbcon, adOpenStatic, adLockReadOnly, 1

If PR_Dumy.EOF Then

    ls_sql = "Insert into Ic_transMasterAutoSale(Compcode, BranchCode, TransCode, TransType, PONO, DCNO, TruckNo, LotNo, InvoiceNo, TransDate, AccountCode, SiteID, BinID, Remarks, SubTotal,"
    ls_sql = ls_sql & " TradeDiscount, Freight, Miscellaneous, TaxAmount, TotalAmount, MiscRemarks,  Discount,   SaleType, SedAmount, RecAmount, UserID, Adddate, AddTime)"
    ls_sql = ls_sql & " select Compcode, BranchCode, TransCode, TransType, PONO, DCNO, TruckNo, LotNo, InvoiceNo, TransDate, AccountCode, SiteID, BinID, Remarks, SubTotal,"
    ls_sql = ls_sql & " TradeDiscount, Freight, Miscellaneous, TaxAmount, TotalAmount, MiscRemarks,  Discount,   SaleType, SedAmount, RecAmount, UserID, Adddate, AddTime  from Ic_transCashMaster where transcode = '" & ls_transcode & "'"
    gc_dbcon.Execute ls_sql

    ls_sql = "insert into  IC_TransAutoSale (Compcode, BranchCode, TransCode, ItemCode, Quantity, ItemRate, Amount)"
    ls_sql = ls_sql & " Select Compcode, BranchCode, TransCode, ItemCode, Quantity, ItemRate, Amount from Ic_transCash where transcode = '" & ls_transcode & "'"
    gc_dbcon.Execute ls_sql

    ln_TotalAmount = ln_TotalAmount + ln_Amount

End If
PR_Dumy.Close
lblStatus = "Total amount in process = " + str(ln_TotalAmount)
Me.Refresh


If ln_TotalAmount >= Val(txtamount) Then
Exit Do
End If

Loop

lblStatus = "Work Status"
Me.Refresh

Call MsgBox("Data for Option2 [Data Import-2] Successfully Imported", vbInformation)


End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
lblStatus.Caption = "Purchase in progress..."
DoEvents
ls_sql = "delete from IC_TransMasterAutoPurchase"
gc_dbcon.Execute ls_sql
ls_sql = "delete from IC_TransAutoPurchase"
gc_dbcon.Execute ls_sql
ls_sql = "insert into  IC_TransMasterAutoPurchase(Compcode, BranchCode, TransCode, TransType, PONO, DCNO, TruckNo, LotNo, InvoiceNo, TransDate, AccountCode, SiteID, BinID, Remarks, SubTotal,"
ls_sql = ls_sql & " TradeDiscount, Freight, Miscellaneous, TaxAmount, TotalAmount, MiscRemarks,  Discount,   SaleType, SedAmount, RecAmount, UserID, Adddate, AddTime)"
ls_sql = ls_sql & " SELECT '001' AS Compcode, '001' AS BranchCode, SUBSTRING('0000000000', 1, 10 - LEN(RTRIM(LTRIM(PurInvcode)))) + RTRIM(LTRIM(PurInvcode)) AS Transcode, 'P' AS TransType, Suppinvcode AS PONO, 'NA' AS DCNO, 'NA' AS TruckNo, 'NA' AS LotNo, '0000000001' AS InvoiceNo, [date] AS TransDate, SUBSTRING('000000', 1, 6 - LEN(RTRIM(LTRIM(SuppCode))))+RTRIM(LTRIM(SuppCode)) AS AccountCode, '001' AS SiteID, '001' AS BinID,  Remarks, InvTotal AS SubTotal, 0 AS TradeDiscount, 0 AS Freight, 0 AS Miscellaneous, 0 AS TaxAmount, InvTotal AS TotalAmount, Remarks AS MiscRemarks, FlatDisc AS Discount, SUBSTRING('000', 1, 3 - LEN(RTRIM(LTRIM(purCatCode))))+RTRIM(LTRIM(purCatCode))  AS SaleType, 0 AS SedAmount,"
ls_sql = ls_sql & " 0 AS RecAmount, 'admin' AS UserID, [date] AS Adddate, '08:00:00' AS AddTime "
ls_sql = ls_sql & " FROM RahatDeptStoresDataBaseV3.dbo.PurLedger where convert(varchar,[date],111) >='" & Format(DtpFrom, "YYYY/MM/DD") & "' and convert(varchar,[date],111) <='" & Format(DtpTo, "YYYY/MM/DD") & "' "

gc_dbcon.Execute ls_sql

ls_sql = "insert into  IC_TransAutoPurchase (Compcode, BranchCode, TransCode, ItemCode, Quantity, ItemRate, Amount,DiscPerc,DiscAmount,TaxPerc,TaxAmount)"
ls_sql = ls_sql & " SELECT '001' AS Compcode, '001' AS BranchCode, SUBSTRING('0000000000', 1, 10 - LEN(RTRIM(LTRIM(purdetail.purInvcode)))) + RTRIM(LTRIM(purdetail.purInvcode)) AS Transcode, SUBSTRING('000000', 1, 6 - LEN(RTRIM(LTRIM(purdetail.ICode)))) + RTRIM(LTRIM(purdetail.ICode)) AS ItemCode, case when purdetail.LooseQty > 0  then purdetail.LooseQty else purdetail.packQty end as Qty  , purdetail.purPrice as Rate ,case when purdetail.LooseQty > 0 then  purdetail.LooseQty* purdetail.purPrice else  purdetail.packQty* purdetail.purPrice end as Amount,"
ls_sql = ls_sql & " Purdetail.discperc, case when purdetail.LooseQty > 0 then  purdetail.LooseQty * purdetail.purPrice else  purdetail.packQty * purdetail.purPrice end * Purdetail.discperc/100 as DiscAmount,"
ls_sql = ls_sql & " Purdetail.gstperc, case when purdetail.LooseQty > 0 then  purdetail.LooseQty * purdetail.purPrice else  purdetail.packQty * purdetail.purPrice end * Purdetail.GSTperc/100 as GStAmount FROM RahatDeptStoresDatabaseV3.dbo.PurLedger  PurLedger INNER JOIN RahatDeptStoresDatabaseV3.dbo.Purdetail Purdetail ON PurLedger.PurInvCode =Purdetail.PurInvcode "
ls_sql = ls_sql & " WHERE convert(varchar,PurLedger.[date],111) >= '" & Format(DtpFrom, "YYYY/MM/DD") & "' and (convert(varchar,PurLedger.[date],111) <= '" & Format(DtpTo, "YYYY/MM/DD") & "')"
gc_dbcon.Execute ls_sql

LoadSupplier
InitializeGrid1
lblStatus = ""
DoEvents
Me.Refresh

End Sub
Private Sub InitializeGrid()
    With GrdGRN
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Vendor Code|<Vendor Name"
        .ColWidth(1) = 1000
        .ColWidth(2) = 3000
        .Redraw = True
    End With
   
End Sub
Private Sub InitializeGrid1()
    With Grid2
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<InvoiceNo|<Supplier Invoice|<SubAmount|<DiscAmount|<TaxAmount|<Total Amount "
        .ColWidth(1) = 1200
        .ColWidth(2) = 1200
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1200
        
        .Redraw = True
    End With
    
End Sub
Private Sub LoadSupplier()
Dim Pr_LoadTrans As New Recordset
Dim ls_sql As String

InitializeGrid

ls_sql = "SELECT IC_TransMasterAutoPurchase.AccountCode, IC_Supplier.Description"
ls_sql = ls_sql & " FROM IC_TransMasterAutoPurchase INNER JOIN"
ls_sql = ls_sql & " IC_Supplier ON IC_TransMasterAutoPurchase.Compcode = IC_Supplier.Compcode AND"
ls_sql = ls_sql & " IC_TransMasterAutoPurchase.AccountCode = IC_Supplier.SupplierCode"
ls_sql = ls_sql & " GROUP BY IC_TransMasterAutoPurchase.AccountCode, IC_Supplier.Description"
    
Pr_LoadTrans.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not Pr_LoadTrans.EOF Then
        With GrdGRN
            Do While Not Pr_LoadTrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = Trim(Pr_LoadTrans("AccountCode") & "")
                .TextMatrix(.Row, 2) = Pr_LoadTrans("Description")
                .Rows = .Rows + 1
                Pr_LoadTrans.MoveNext
                If Pr_LoadTrans.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
        
    End If
Pr_LoadTrans.Close
Exit Sub
LocalErr:
Call MsgBox(Err.Description)


End Sub

Private Sub Command4_Click()
LoadSupplier
End Sub

Private Sub Form_Load()
DtpFrom = Date
DtpTo = Date
End Sub
Private Sub LoadInvoiceData(ls_Suppcode As String)
Dim Pr_LoadTrans As New Recordset
Dim ls_sql As String

InitializeGrid1

ls_sql = "SELECT IC_TransMasterAutoPurchase.TransCode, IC_TransMasterAutoPurchase.PONO, SUM(IC_TransAutoPurchase.Amount) AS Amount,"
ls_sql = ls_sql & " SUM(IC_TransAutoPurchase.TaxAmount) AS TaxAmount, SUM(IC_TransAutoPurchase.DiscAmount) AS DiscAmount,"
ls_sql = ls_sql & " SUM(IC_TransAutoPurchase.Amount) - SUM(IC_TransAutoPurchase.DiscAmount) + SUM(IC_TransAutoPurchase.TaxAmount) AS TotalAmount"
ls_sql = ls_sql & " FROM IC_TransMasterAutoPurchase INNER JOIN IC_TransAutoPurchase ON IC_TransMasterAutoPurchase.Compcode = IC_TransAutoPurchase.Compcode AND"
ls_sql = ls_sql & " IC_TransMasterAutoPurchase.BranchCode = IC_TransAutoPurchase.BranchCode AND"
ls_sql = ls_sql & " IC_TransMasterAutoPurchase.TransCode = IC_TransAutoPurchase.TransCode"
ls_sql = ls_sql & " WHERE (IC_TransMasterAutoPurchase.AccountCode = '" & ls_Suppcode & "')"
ls_sql = ls_sql & " GROUP BY IC_TransMasterAutoPurchase.TransCode, IC_TransMasterAutoPurchase.PONO"

Pr_LoadTrans.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not Pr_LoadTrans.EOF Then
        With Grid2
            Do While Not Pr_LoadTrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = Trim(Pr_LoadTrans("TransCode") & "")
                .TextMatrix(.Row, 2) = Trim(Pr_LoadTrans("PONO") & "")
                .TextMatrix(.Row, 3) = Pr_LoadTrans("Amount")
                .TextMatrix(.Row, 4) = Pr_LoadTrans("TaxAmount")
                .TextMatrix(.Row, 5) = Pr_LoadTrans("DiscAmount")
                .TextMatrix(.Row, 6) = Pr_LoadTrans("Totalamount")
                .Rows = .Rows + 1
                Pr_LoadTrans.MoveNext
                If Pr_LoadTrans.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
        
    End If
Pr_LoadTrans.Close
Exit Sub
LocalErr:
Call MsgBox(Err.Description)


End Sub

Private Sub GrdGRN_Click()
Call GrdGRN_DblClick
End Sub

Private Sub GrdGRN_DblClick()
With GrdGRN
.CellForeColor = QBColor(12)
 LoadInvoiceData (.TextMatrix(.Row, 1))

End With
End Sub

Private Sub GrdGRN_KeyDown(KeyCode As Integer, Shift As Integer)
With GrdGRN
        If KeyCode = vbKeyDelete Then
          
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
                  ls_sql = "delete from IC_TransAutoPurchase where Transcode in (select Transcode from IC_TransMasterAutoPurchase where accountcode =  '" & .TextMatrix(.Row, 1) & "') "
                  gc_dbcon.Execute ls_sql
                  ls_sql = "delete from IC_TransMasterAutoPurchase where AccountCode = '" & .TextMatrix(.Row, 1) & "'  "
                  gc_dbcon.Execute ls_sql
                 .RemoveItem .Row
    
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                .TextMatrix(.Row, 0) = ""
                PI_SrNo = 0
            End If
        End If
    End With
End Sub

Private Sub Grid2_DblClick()
With Grid2
    Printinvoice (.TextMatrix(.Row, 1))
End With
End Sub

Private Sub Grid2_KeyDown(KeyCode As Integer, Shift As Integer)
    With Grid2
        If KeyCode = vbKeyDelete Then
          
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
                  ls_sql = "delete from IC_TransMasterAutoPurchase where Transcode = '" & .TextMatrix(.Row, 1) & "'  "
                  gc_dbcon.Execute ls_sql
                  ls_sql = "delete from IC_TransAutoPurchase where Transcode = '" & .TextMatrix(.Row, 1) & "' "
                  gc_dbcon.Execute ls_sql
                 .RemoveItem .Row
    
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                .TextMatrix(.Row, 0) = ""
                PI_SrNo = 0
            End If
        End If
    End With
End Sub

Private Sub Printinvoice(ls_invoice As String)
With rptLedger
        .DiscardSavedData = True
        .SelectionFormula = ""
        .WindowTitle = ls_Caption
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = 'Purchase Invoice'"
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "AutoPurchaseInvoice.rpt"
        .SelectionFormula = "{PO_TransMaster.Compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {PO_TransMaster.Transcode} ='" & ls_invoice & "'"
        .Action = 1
        
End With
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmImportData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Data"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3390
   Icon            =   "frmImportData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox CheckOption2 
      Caption         =   "Cash Sale"
      Height          =   345
      Left            =   60
      TabIndex        =   12
      Top             =   1740
      Width           =   1500
   End
   Begin VB.CheckBox CheckOption1 
      Caption         =   "Credit Sale"
      Height          =   345
      Left            =   60
      TabIndex        =   11
      Top             =   1425
      Value           =   1  'Checked
      Width           =   1500
   End
   Begin VB.Frame Frame2 
      Height          =   510
      Left            =   30
      TabIndex        =   9
      Top             =   1995
      Width           =   3330
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
         TabIndex        =   10
         Top             =   180
         Width           =   3195
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   2100
      TabIndex        =   8
      Top             =   2595
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Generate"
      Height          =   345
      Left            =   900
      TabIndex        =   7
      Top             =   2595
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   1470
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   3315
      Begin VB.TextBox txtAmount 
         Height          =   375
         Left            =   1500
         TabIndex        =   1
         Top             =   990
         Width           =   1680
      End
      Begin MSComCtl2.DTPicker DtpFrom 
         Height          =   345
         Left            =   1500
         TabIndex        =   2
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         Format          =   24772609
         CurrentDate     =   40600
      End
      Begin MSComCtl2.DTPicker DtpTo 
         Height          =   345
         Left            =   1500
         TabIndex        =   3
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         Format          =   24772609
         CurrentDate     =   40600
      End
      Begin VB.Label Label4 
         Caption         =   "Amount :"
         Height          =   240
         Left            =   795
         TabIndex        =   6
         Top             =   1035
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "To Date :"
         Height          =   240
         Left            =   750
         TabIndex        =   5
         Top             =   645
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "From Date :"
         Height          =   240
         Left            =   615
         TabIndex        =   4
         Top             =   225
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmImportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ls_sql As String

Private Sub Command1_Click()
lblStatus.Caption = "Credit Sale in progress..."
Me.Refresh
DoEvents
If CheckOption1.Value = 1 Then
ls_sql = "delete from IC_TransMasterAutoSale"
gc_dbcon.Execute ls_sql
ls_sql = "delete from IC_TransAutoSale"
gc_dbcon.Execute ls_sql
ls_sql = "insert into  IC_TransMasterAutoSale(Compcode, BranchCode, TransCode, TransType, PONO, DCNO, TruckNo, LotNo, InvoiceNo, TransDate, AccountCode, SiteID, BinID, Remarks, SubTotal,"
ls_sql = ls_sql & " TradeDiscount, Freight, Miscellaneous, TaxAmount, TotalAmount, MiscRemarks,  Discount,   SaleType, SedAmount, RecAmount, UserID, Adddate, AddTime)"
ls_sql = ls_sql & " SELECT '001' AS Compcode, '001' AS BranchCode, SUBSTRING('0000000000', 1, 10 - LEN(RTRIM(LTRIM(SaleInvcode))))"
ls_sql = ls_sql & " + RTRIM(LTRIM(SaleInvcode)) AS Transcode, 'S' AS TransType, 'NA' AS PONO, 'NA' AS DCNO, 'NA' AS TruckNo, 'NA' AS LotNo,"
ls_sql = ls_sql & " '0000000001' AS InvoiceNo, [date] AS TransDate, SUBSTRING('000000', 1, 6 - LEN(RTRIM(LTRIM(CustCode))))+RTRIM(LTRIM(CustCode)) AS AccountCode, '001' AS SiteID, '001' AS BinID, 'Sale' AS Remarks, InvTotal AS SubTotal,"
ls_sql = ls_sql & " 0 AS TradeDiscount, 0 AS Freight, 0 AS Miscellaneous, 0 AS TaxAmount, InvTotal AS TotalAmount, 'Sale' AS MiscRemarks, FlatDisc AS Discount,"
ls_sql = ls_sql & " SUBSTRING('000', 1, 3 - LEN(RTRIM(LTRIM(SaleCatCode))))+RTRIM(LTRIM(SaleCatCode))  AS SaleType, 0 AS SedAmount, 0 AS RecAmount, 'admin' AS UserID, [date] AS Adddate, '08:00:00' AS AddTime"
ls_sql = ls_sql & " FROM RahatDeptStoresDataBaseV3.dbo.SaleLedger where convert(varchar,[date],111) >='" & Format(DtpFrom, "YYYY/MM/DD") & "' and convert(varchar,[date],111) <='" & Format(DtpTo, "YYYY/MM/DD") & "' and salecatcode = 2"
gc_dbcon.Execute ls_sql
DoEvents
ls_sql = "insert into  IC_TransAutoSale (Compcode, BranchCode, TransCode, ItemCode, Quantity, ItemRate, Amount,DiscPerc,DiscAmount,TaxPerc,TaxAmount)"
ls_sql = ls_sql & " SELECT '001' AS Compcode, '001' AS BranchCode, SUBSTRING('0000000000', 1, 10 - LEN(RTRIM(LTRIM(Saledetail.SaleInvcode))))"
ls_sql = ls_sql & " + RTRIM(LTRIM(Saledetail.SaleInvcode)) AS Transcode, SUBSTRING('000000', 1, 6 - LEN(RTRIM(LTRIM(Saledetail.ICode))))"
ls_sql = ls_sql & " + RTRIM(LTRIM(Saledetail.ICode)) AS ItemCode, Saledetail.LooseQty, Saledetail.SalePrice as Rate ,Saledetail.LooseQty* Saledetail.SalePrice as Amount ,"
ls_sql = ls_sql & " Saledetail.Discperc AS Discperc, (((Saledetail.LooseQty* Saledetail.SalePrice)*Saledetail.Discperc)/100) as DiscAmount,0.75 as Taxperc ,(((Saledetail.LooseQty* Saledetail.SalePrice)*0.75)/100) as Taxamount "

ls_sql = ls_sql & " FROM RahatDeptStoresDatabaseV3.dbo.SaleLedger  SaleLedger INNER JOIN"
ls_sql = ls_sql & " RahatDeptStoresDatabaseV3.dbo.Saledetail Saledetail ON SaleLedger.SaleInvCode = Saledetail.SaleInvcode"
ls_sql = ls_sql & " WHERE (convert(varchar , SaleLedger.[date],111) >= '" & Format(DtpFrom, "YYYY/MM/DD") & "' and (convert(varchar , SaleLedger.[date],111) <= '" & Format(DtpTo, "YYYY/MM/DD") & "') and SaleLedger.salecatcode = 2)"
gc_dbcon.Execute ls_sql
DoEvents
Call MsgBox("Credit Sale Successfully Imported", vbInformation)
lblStatus = ""
Me.Refresh
DoEvents

End If


If CheckOption2.Value = 1 Then
 Call ImportCashSale
End If

lblStatus = ""
Me.Refresh

End Sub
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
ls_sql = ls_sql & " FROM RahatDeptStoresDataBaseV3.dbo.SaleLedger where convert(varchar,[date],111) >='" & Format(DtpFrom, "YYYY/MM/DD") & "' and convert(varchar,[date],111) <='" & Format(DtpTo, "YYYY/MM/DD") & "' and salecatcode = 3"
gc_dbcon.Execute ls_sql

ls_sql = "insert into  IC_TransCash (Compcode, BranchCode, TransCode, ItemCode, Quantity, ItemRate, Amount)"
ls_sql = ls_sql & " SELECT '001' AS Compcode, '001' AS BranchCode, SUBSTRING('0000000000', 1, 10 - LEN(RTRIM(LTRIM(Saledetail.SaleInvcode))))"
ls_sql = ls_sql & " + RTRIM(LTRIM(Saledetail.SaleInvcode)) AS Transcode, SUBSTRING('000000', 1, 6 - LEN(RTRIM(LTRIM(Saledetail.ICode))))"
ls_sql = ls_sql & " + RTRIM(LTRIM(Saledetail.ICode)) AS ItemCode, Saledetail.LooseQty, Saledetail.SalePrice as Rate ,Saledetail.LooseQty* Saledetail.SalePrice as Amount"
ls_sql = ls_sql & " FROM RahatDeptStoresDatabaseV3.dbo.SaleLedger  SaleLedger INNER JOIN"
ls_sql = ls_sql & " RahatDeptStoresDatabaseV3.dbo.Saledetail Saledetail ON SaleLedger.SaleInvCode = Saledetail.SaleInvcode"
ls_sql = ls_sql & " WHERE convert(varchar,SaleLedger.[date],111) >= '" & Format(DtpFrom, "YYYY/MM/DD") & "' and convert(varchar,SaleLedger.[date],111) <= '" & Format(DtpTo, "YYYY/MM/DD") & "' and SaleLedger.salecatcode = 3"
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
DoEvents
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
DoEvents
Loop

lblStatus = ""
Me.Refresh

Call MsgBox("Cash Sale Successfully Imported", vbInformation)


End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
DtpFrom = Date
DtpTo = Date
End Sub

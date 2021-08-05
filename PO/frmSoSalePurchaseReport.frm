VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSoSalePurchaseReport 
   Caption         =   "Purchase Sale Report"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmSoSalePurchaseReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkFullDetails 
      Caption         =   "Full Details Reports"
      Height          =   300
      Left            =   1440
      TabIndex        =   32
      Top             =   3600
      Width           =   2145
   End
   Begin VB.CheckBox chkgraph 
      Caption         =   "Graph Only"
      Height          =   300
      Left            =   1440
      TabIndex        =   21
      Top             =   3240
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   4755
      TabIndex        =   10
      Top             =   3585
      Width           =   1005
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3735
      TabIndex        =   9
      Top             =   3585
      Width           =   1005
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3615
      Left            =   30
      TabIndex        =   6
      Top             =   -45
      Width           =   5730
      Begin VB.ComboBox txtgroupon 
         Height          =   330
         ItemData        =   "frmSoSalePurchaseReport.frx":030A
         Left            =   1485
         List            =   "frmSoSalePurchaseReport.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2925
         Width           =   2745
      End
      Begin VB.TextBox txtclasscode 
         Height          =   315
         Left            =   1485
         MaxLength       =   3
         TabIndex        =   27
         Top             =   630
         Width           =   615
      End
      Begin VB.TextBox txtClassDesc 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   630
         Width           =   3210
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2115
         Picture         =   "frmSoSalePurchaseReport.frx":0339
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   630
         Width           =   315
      End
      Begin VB.TextBox txtsuppliercode 
         Height          =   315
         Left            =   1485
         MaxLength       =   6
         TabIndex        =   24
         Top             =   1005
         Width           =   615
      End
      Begin VB.TextBox txtSupplierdesc 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1020
         Width           =   3210
      End
      Begin VB.CommandButton Command9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2115
         Picture         =   "frmSoSalePurchaseReport.frx":04AB
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1020
         Width           =   315
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   2115
         Picture         =   "frmSoSalePurchaseReport.frx":061D
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   315
      End
      Begin VB.TextBox txtdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2460
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   255
         Width           =   3210
      End
      Begin VB.TextBox txtselectedcode 
         Height          =   315
         Left            =   1485
         TabIndex        =   0
         Top             =   255
         Width           =   600
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2370
         Picture         =   "frmSoSalePurchaseReport.frx":078F
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1425
         Width           =   315
      End
      Begin VB.TextBox txtitemdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2715
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1425
         Width           =   2955
      End
      Begin VB.TextBox txtitemcode 
         Height          =   315
         Left            =   1485
         TabIndex        =   1
         Top             =   1425
         Width           =   885
      End
      Begin VB.CheckBox chkcash 
         Caption         =   "Cash"
         Height          =   270
         Left            =   4590
         TabIndex        =   13
         Top             =   1785
         Width           =   1020
      End
      Begin VB.CheckBox chkcredit 
         Caption         =   "Credit"
         Height          =   270
         Left            =   3780
         TabIndex        =   12
         Top             =   1785
         Width           =   1020
      End
      Begin VB.ComboBox txtpstype 
         Height          =   330
         ItemData        =   "frmSoSalePurchaseReport.frx":0901
         Left            =   1485
         List            =   "frmSoSalePurchaseReport.frx":091A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2535
         Width           =   2745
      End
      Begin MSComCtl2.DTPicker dtpto 
         Height          =   315
         Left            =   1485
         TabIndex        =   3
         Top             =   2190
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62914561
         CurrentDate     =   37309
      End
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   1485
         TabIndex        =   2
         Top             =   1830
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62914561
         CurrentDate     =   37309
      End
      Begin Crystal.CrystalReport rptLedger 
         Left            =   0
         Top             =   0
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
         AutoSize        =   -1  'True
         Caption         =   "Report Group on :"
         Height          =   210
         Left            =   150
         TabIndex        =   31
         Top             =   2970
         Width           =   1290
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Category Code :"
         Height          =   210
         Left            =   285
         TabIndex        =   29
         Top             =   660
         Width           =   1170
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Supplier Code :"
         Height          =   210
         Left            =   330
         TabIndex        =   28
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Department Code :"
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label txtselectiveaccount 
         Height          =   300
         Left            =   2745
         TabIndex        =   17
         Top             =   675
         Visible         =   0   'False
         Width           =   2520
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Item Code :"
         Height          =   210
         Left            =   645
         TabIndex        =   16
         Top             =   1455
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Report Type :"
         Height          =   210
         Left            =   450
         TabIndex        =   11
         Top             =   2580
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "From Date :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   615
         TabIndex        =   8
         Top             =   1875
         Width           =   825
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "To Date :"
         Height          =   210
         Left            =   780
         TabIndex        =   7
         Top             =   2220
         Width           =   645
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   3975
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "Description"
            TextSave        =   "Description"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSoSalePurchaseReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PR_Item As New Recordset
Public PO_DESC As Object
Public PO_CODE As Object
Dim pr_dumy As New Recordset
Dim ls_sql As String
Dim ld_netday

Private Sub chkcredit_Click()
If chkcredit.Value = 1 Then
chkcash.Value = 0
End If
End Sub
Private Sub chkcash_Click()
If chkcash.Value = 1 Then
chkcredit.Value = 0
End If
End Sub

Private Sub cmdGenerate_Click()
On Error GoTo LocalErr

MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."



If Me.Caption = "Sale Report Item History With Cost" Then


With rptLedger
        .DiscardSavedData = True
        .WindowTitle = Me.Caption
        
        .SQLQuery = "SELECT SaleReportCostH.TransCode, SaleReportCostH.TransDate, SaleReportCostH.ItemCode, SaleReportCostH.Quantity, SaleReportCostH.NetAmount,"
        .SQLQuery = .SQLQuery & " SaleReportCostH.AvgRate, SaleReportCostH.CostAmount, IC_Item.Description FROM   SaleReportCostHistory SaleReportCostH LEFT OUTER JOIN   IC_Item IC_Item ON SaleReportCostH.ItemCode = IC_Item.ItemCode"
        .SQLQuery = .SQLQuery & " where SaleReportCostH.TransDate >='" & Format(dtpfrom, "YYYY/MM/DD") & "'"
        .SQLQuery = .SQLQuery & " and SaleReportCostH.TransDate <='" & Format(DTPTo, "YYYY/MM/DD") & "'"
        
        If txtitemcode <> "" Then
        .SQLQuery = .SQLQuery & " and SaleReportCostH.ItemCode ='" & txtitemcode & "'"
        End If
        
        .SQLQuery = .SQLQuery & " ORDER BY SaleReportCostH.ItemCode, SaleReportCostH.TransDate"
        
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "SaleReportCostBase.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = '" & Me.Caption & " '"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
End With

Else
With rptLedger
        .Formulas(0) = ""
        .Formulas(1) = ""
        .Formulas(2) = ""
        .Formulas(3) = ""
End With
 
Call ChkTempTables("Tmp_SalePurchase", True)

If txtpstype.ListIndex = 2 Then
ls_sql = "SELECT IC_Item.ItemCode, IC_Item.Description, SUM(PO_POGRNDetail.Quantity) AS Qty, (SUM(PO_POGRNDetail.Amount)+sum(PO_POGRNDetail.GSTAmount)) - SUM(PO_POGRNDetail.DiscAmount) AS Amount, PO_POGRN.Accountcode,IC_Supplier.Description AS Supplier Into Tmp_SalePurchase"
ls_sql = ls_sql & " FROM PO_POGRN INNER JOIN"
ls_sql = ls_sql & " PO_POGRNDetail ON PO_POGRN.Compcode = PO_POGRNDetail.Compcode AND PO_POGRN.TransCode = PO_POGRNDetail.TransCode INNER JOIN"
ls_sql = ls_sql & " IC_Item ON PO_POGRNDetail.Compcode = IC_Item.Compcode AND PO_POGRNDetail.ItemCode = IC_Item.ItemCode INNER JOIN"
ls_sql = ls_sql & " IC_Supplier ON PO_POGRN.Compcode = IC_Supplier.Compcode AND PO_POGRN.AccountCode = IC_Supplier.SupplierCode"
ls_sql = ls_sql & " WHERE (PO_POGRN.TransDate >= '" & Format(dtpfrom, "YYYY/MM/DD") & "') AND (PO_POGRN.TransDate <=  '" & Format(DTPTo, "YYYY/MM/DD") & "')"
If chkcredit.Value = 1 Then
ls_sql = ls_sql & " and (PO_POGRN.Type = 1)"
ElseIf chkcash.Value = 1 Then
ls_sql = ls_sql & "  and (PO_POGRN.Type <> 1)"
End If

If txtselectedcode <> "" Then
    ls_sql = ls_sql & "  and IC_Item.CatCode = '" & txtselectedcode & "'"

End If


If txtSuppliercode <> "" Then
    ls_sql = ls_sql & "  and PO_POGRN.Accountcode = '" & txtSuppliercode & "'"

End If
If txtitemcode <> "" Then
    If txtitemcode = "Selective" Then
    ls_sql = ls_sql & "  and PO_POGRNDetail.itemcode in (" & txtselectiveaccount & ")"
    Else
    ls_sql = ls_sql & "  and PO_POGRNDetail.itemcode = '" & txtitemcode & "'"
    End If
End If

ls_sql = ls_sql & " GROUP BY IC_Item.ItemCode, IC_Item.Description,PO_POGRN.Accountcode, IC_Supplier.Description"

gc_dbcon.Execute ls_sql

With rptLedger
        .DiscardSavedData = True
        .WindowTitle = Me.Caption
        If ChkFullDetails.Value = 1 Then
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "PurchasereportitemwiseDetails.rpt"
        Else
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "Purchasereportitemwise.rpt"
        End If
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        
        If txtselectedcode <> "" Then
         .Formulas(1) = "Reportname = 'Purchase Report For. " & txtdesc & "'"
        ElseIf txtSuppliercode <> "" Then
         .Formulas(1) = "Reportname = 'Purchase Report For. " & txtSupplierdesc & "'"
        ElseIf txtitemcode <> "" Then
        .Formulas(1) = "Reportname = 'Purchase Report For. " & txtitemdesc & "'"
        Else
        .Formulas(1) = "Reportname = 'Purchase Report'"
        End If
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
End With

ElseIf txtpstype.ListIndex = 3 Or txtpstype.ListIndex = 4 Or txtpstype.ListIndex = 5 Then
ls_sql = "SELECT IC_Item.ItemCode, IC_Item.Description, SUM(PO_POGRNDetail.Quantity) AS PQty, SUM(PO_POGRNDetail.Amount) - SUM(PO_POGRNDetail.DiscAmount) AS PAmount,0 as SaleQty,0 as SaleAmount,IC_Item.catcode Into Tmp_SalePurchase "
ls_sql = ls_sql & " FROM PO_POGRN INNER JOIN "
ls_sql = ls_sql & " PO_POGRNDetail ON PO_POGRN.Compcode = PO_POGRNDetail.Compcode AND PO_POGRN.TransCode = PO_POGRNDetail.TransCode INNER JOIN"
ls_sql = ls_sql & " IC_Item ON PO_POGRNDetail.Compcode = IC_Item.Compcode AND PO_POGRNDetail.ItemCode = IC_Item.ItemCode "
ls_sql = ls_sql & " WHERE (PO_POGRN.TransDate >= '" & Format(dtpfrom, "YYYY/MM/DD") & "') AND (PO_POGRN.TransDate <=  '" & Format(DTPTo, "YYYY/MM/DD") & "')"
If chkcredit.Value = 1 Then
ls_sql = ls_sql & " and (PO_POGRN.Type = 1)"
ElseIf chkcash.Value = 1 Then
ls_sql = ls_sql & "  and (PO_POGRN.Type <> 1)"
End If



If txtselectedcode <> "" Then
    ls_sql = ls_sql & "  and IC_Item.catcode = '" & txtselectedcode & "'"
End If
If txtclasscode <> "" Then
    ls_sql = ls_sql & "  and IC_Item.Classid = '" & txtclasscode & "'"
End If

If txtSuppliercode <> "" Then
    ls_sql = ls_sql & "  and IC_Item.ManuCode = '" & txtSuppliercode & "'"
End If

If txtitemcode <> "" Then
    ls_sql = ls_sql & "  and PO_POGRNDetail.itemcode = '" & txtitemcode & "'"
End If



ls_sql = ls_sql & " GROUP BY IC_Item.CatCode,IC_Item.ItemCode, IC_Item.Description"


ls_sql = ls_sql & " Union All "
'--
ls_sql = ls_sql & " SELECT SaleReport.itemcode, IC_Item.Description,0 as PQty,0 as PAmount, SUM(SaleReport.SaleQty) AS SaleQty, SUM(SaleReport.Amount) - SUM(SaleReport.DiscAmount) AS SaleAmount,IC_Item.catcode "
ls_sql = ls_sql & " FROM SaleReport LEFT OUTER JOIN  IC_Item ON SaleReport.Compcode = IC_Item.Compcode AND SaleReport.itemcode = IC_Item.ItemCode"
ls_sql = ls_sql & " WHERE (SaleReport.TransDate >= '" & Format(dtpfrom, "YYYY/MM/DD") & "') AND (SaleReport.TransDate <=  '" & Format(DTPTo, "YYYY/MM/DD") & "')"

If chkcredit.Value = 1 Then
ls_sql = ls_sql & " and (SaleReport.salestatus = 1)"
ElseIf chkcash.Value = 1 Then
ls_sql = ls_sql & "  and (SaleReport.salestatus <> 1)"
End If

If txtselectedcode <> "" Then
ls_sql = ls_sql & "  and IC_Item.catcode = '" & txtselectedcode & "'"
End If

If txtitemcode <> "" Then
    ls_sql = ls_sql & "  and SaleReport.ItemCode = '" & txtitemcode & "'"
End If

If txtclasscode <> "" Then
    ls_sql = ls_sql & "  and IC_Item.Classid = '" & txtclasscode & "'"
End If

If txtSuppliercode <> "" Then
    ls_sql = ls_sql & "  and IC_Item.ManuCode = '" & txtSuppliercode & "'"
End If


ls_sql = ls_sql & " GROUP BY IC_Item.CatCode,SaleReport.itemcode, IC_Item.Description"




gc_dbcon.Execute ls_sql




With rptLedger
        .DiscardSavedData = True
        .WindowTitle = Me.Caption
        If txtpstype.ListIndex = 4 Then
        ' gc_dbcon.Execute ("Select itemcode, IC_Item.Description,IC_Item.CatCode,Sum(PQty) as Pqty,Sum(PAmount) as PAmount,Sum(SQty) as SQty,Sum(SAmount) as SAmount From Tmp_SalePurchase Group by SaleReport.itemcode, IC_Item.Description,IC_Item.CatCode")
        '.ReportFileName = App.Path & Gs_ICRepoPath & "\" & "SalerPurchaseeportitemwise1.rpt"
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "SalerPurchaseeportitemwise3.rpt"
        ElseIf txtpstype.ListIndex = 5 Then
        ls_sql = "delete from Tmp_SalePurchase where pqty > 0 "
        gc_dbcon.Execute ls_sql
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "SalerPurchaseeportitemwise2.rpt"
        ElseIf txtpstype.ListIndex = 6 Then
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "SalerPurchaseeportSupplierWise.rpt"
        Else
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "SalerPurchaseeportitemwise.rpt"
        End If
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = 'Sale Purchase Report'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .Formulas(3) = "Groupon = " & txtgroupon.ListIndex & ""
        .Connect = "DNS=Censoft;UID=Sa"
        
        .Action = 1
End With

ElseIf txtpstype.ListIndex = 1 Then

ls_sql = "SELECT SaleReport.transdate, SaleReport.itemcode, SUM(CASE WHEN SaleReport.salestatus = 0 THEN SaleReport.amount-SaleReport.DiscAmount ELSE 0 END) AS Amount,"
ls_sql = ls_sql & " ic_item.catcode, SUM(CASE WHEN SaleReport.salestatus = 1 THEN SaleReport.amount-SaleReport.DiscAmount ELSE 0 END) AS CreditAmount into Tmp_SalePurchase"
ls_sql = ls_sql & " FROM SaleReport LEFT OUTER JOIN  ic_item ON SaleReport.compcode = ic_item.compcode AND SaleReport.itemcode = ic_item.itemcode"
ls_sql = ls_sql & " WHERE (SaleReport.TransDate >= '" & Format(dtpfrom, "YYYY/MM/DD") & "') AND (SaleReport.TransDate <=  '" & Format(DTPTo, "YYYY/MM/DD") & "' )"

If chkcredit.Value = 1 Then
ls_sql = ls_sql & " and (SaleReport.salestatus = 1)"
ElseIf chkcash.Value = 1 Then
ls_sql = ls_sql & "  and (SaleReport.salestatus <> 1)"
End If


If txtselectedcode <> "" Then
ls_sql = ls_sql & "  and IC_Item.catcode = '" & txtselectedcode & "'"
End If


If txtitemcode <> "" Then
    If txtitemcode = "Selective" Then
    ls_sql = ls_sql & "  and SaleReport.itemcode in (" & txtselectiveaccount & ")"
    Else
    ls_sql = ls_sql & "  and SaleReport.itemcode = '" & txtitemcode & "'"
    End If
End If


ls_sql = ls_sql & " GROUP BY SaleReport.transdate, ic_item.catcode, SaleReport.itemcode"

gc_dbcon.Execute ls_sql


With rptLedger
        .DiscardSavedData = True
        .WindowTitle = Me.Caption
        If chkgraph.Value = 1 Then
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "SalereportDaywisegraph.rpt"
        Else
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "SalereportDaywise.rpt"
        End If
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = 'Sale Report'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
End With


'--
Else

ls_sql = "SELECT SaleReport.itemcode, IC_Item.catcode,IC_Item.Description, SUM(SaleReport.SaleQty) AS Qty, SUM(SaleReport.Amount) - SUM(SaleReport.DiscAmount) AS Amount Into Tmp_SalePurchase"
ls_sql = ls_sql & " FROM SaleReport LEFT OUTER JOIN  IC_Item ON SaleReport.Compcode = IC_Item.Compcode AND SaleReport.itemcode = IC_Item.ItemCode"
ls_sql = ls_sql & " WHERE (SaleReport.TransDate >= '" & Format(dtpfrom, "YYYY/MM/DD") & "') AND (SaleReport.TransDate <=  '" & Format(DTPTo, "YYYY/MM/DD") & "' )"

If chkcredit.Value = 1 Then
ls_sql = ls_sql & " and (SaleReport.salestatus = 1)"
ElseIf chkcash.Value = 1 Then
ls_sql = ls_sql & "  and (SaleReport.salestatus <> 1)"
End If

If txtitemcode <> "" Then
    If txtitemcode = "Selective" Then
    ls_sql = ls_sql & "  and SaleReport.customcode in (" & txtselectiveaccount & ")"
    Else
    ls_sql = ls_sql & "  and SaleReport.itemcode = '" & txtitemcode & "'"
    End If
End If


If txtselectedcode <> "" Then
    If txtselectedcode = "Selective" Then
    ls_sql = ls_sql & "  and SaleReport.customcode in (" & txtselectiveaccount & ")"
    Else
    ls_sql = ls_sql & "  and IC_Item.Catcode = '" & txtselectedcode & "'"
    End If
End If

If txtSuppliercode <> "" Then
    If txtSuppliercode = "Selective" Then
    ls_sql = ls_sql & "  and SaleReport.customcode in (" & txtselectiveaccount & ")"
    Else
    ls_sql = ls_sql & "  and IC_Item.Manucode = '" & txtSuppliercode & "'"
    End If
End If



ls_sql = ls_sql & " GROUP BY SaleReport.itemcode, IC_Item.Catcode, IC_Item.Description"

gc_dbcon.Execute ls_sql


With rptLedger
         ld_netday = DateDiff("D", dtpfrom, DTPTo) + 1
        .DiscardSavedData = True
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "Salereportitemwise.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = 'Sale Report'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .Formulas(3) = "NetDays = " & ld_netday & ""
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
End With

End If
End If
MDIForm1.StatusBar1.Panels(7).Text = ""
Exit Sub
LocalErr:
Call SetErr(Err.Description, vbCritical)
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtitemcode
    Set PO_DESC = txtitemdesc
    Gs_SQL = "Select IC_Item.ItemCode,   IC_Item.Description, IC_ItemCategory.Description as Category,IC_Item.SaleCost from IC_Item left outer join IC_ItemCategory on IC_Item.compcode = IC_ItemCategory.compcode and   IC_Item.catcode = IC_ItemCategory.catcode "
    Gs_FindFld = "IC_Item.Description"
    Gs_OrderBy = "Order by IC_Item.Description"
    Gs_OtherPara = " where IC_Item.compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Items"
    MyLookupOLDB.Show 1

    If Len(txtitemcode) > 0 Then txtItemcode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtselectedcode
    Set PO_DESC = txtdesc
    Gs_SQL = "SELECT CatCode, Description  FROM IC_ItemCategory "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' "
    MyLookupOLDB.Caption = "Department"
    MyLookupOLDB.Show 1
    
    If txtselectedcode <> "" Then Call txtselectedcode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtclasscode
    Set PO_DESC = txtClassDesc
    Gs_SQL = "Select ClassCode,   Description from IC_ItemClass "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Item Class"
    MyLookupOLDB.Show 1
    
    If txtclasscode <> "" Then Call txtclassCode_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command9_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtSuppliercode
    Set PO_DESC = txtSupplierdesc
    
    Gs_SQL = "Select SupplierCode,   Description from IC_Supplier "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Supplier"
    MyLookupOLDB.Show 1
    
    If txtSuppliercode <> "" Then Call txtSuppliercode_KeyDown(vbKeyReturn, vbKeyShift)


End Sub


Private Sub txtSuppliercode_Change()
If txtSuppliercode = "" Then
txtSupplierdesc = ""
End If
End Sub

Private Sub txtSuppliercode_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(txtSuppliercode) <> "" And KeyCode = vbKeyReturn Then
        txtSuppliercode.Text = DoPad(txtSuppliercode.Text, 6)
        pr_dumy.Open "Select * from IC_Supplier where Suppliercode = '" & txtSuppliercode & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Supplier code not found !!!", vbCritical)
            txtSuppliercode = ""
            txtSupplierdesc = ""
            txtSuppliercode.SetFocus
        Else
            txtSupplierdesc = pr_dumy("Description")
            If txtitemcode.Enabled Then txtitemcode.SetFocus
            
        End If
        pr_dumy.Close
ElseIf Trim(txtSuppliercode) = "" And KeyCode = vbKeyReturn Then
        txtSuppliercode = ""
        txtSupplierdesc = ""
        Command9_Click
End If

End Sub
       

Private Sub txtclassCode_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtclasscode) <> "" And KeyCode = vbKeyReturn Then
        txtclasscode = DoPad(txtclasscode, 3)
       If pr_dumy.State = 1 Then pr_dumy.Close
        pr_dumy.Open "Select * from Ic_ItemClass where Classcode = '" & txtclasscode & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("ClassID not found !!!", vbCritical)
            txtclasscode = ""
            txtClassDesc = ""
            txtclasscode.SetFocus
        Else
            txtClassDesc = pr_dumy("Description")
            If txtSuppliercode.Enabled Then txtSuppliercode.SetFocus
            
        End If
        pr_dumy.Close
        
ElseIf Trim(txtclasscode) = "" And KeyCode = vbKeyReturn Then
        txtclasscode = ""
        txtClassDesc = ""
        Command4_Click
End If
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdGenerate.SetFocus
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then DTPTo.SetFocus
End Sub
Private Sub Form_Load()
dtpfrom = Date
DTPTo = Date
txtpstype.ListIndex = 3

End Sub

Private Sub txtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
If txtitemcode <> "" And KeyCode = vbKeyReturn Then
    If txtitemcode <> "Selective" Then
    ls_sql = "Select itemcode,Description from IC_Item where compcode = '" & Gs_compcode & "' and Itemcode = '" & txtitemcode & "' "
    
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Item Code not found", vbCritical)
            Else
                 txtitemdesc = pr_dumy("description")
                 dtpfrom.SetFocus
            End If
         pr_dumy.Close
    End If
End If

End Sub

Private Sub txtpstype_Click()
Me.Caption = txtpstype.Text
End Sub

Private Sub txtselectedcode_KeyDown(KeyCode As Integer, Shift As Integer)
If txtselectedcode <> "" And KeyCode = vbKeyReturn Then
    If txtselectedcode <> "Selective" Then
        txtselectedcode = DoPad(txtselectedcode, 3)
    ls_sql = "SELECT CatCode, Description  FROM IC_ItemCategory where compcode = '" & Gs_compcode & "' and CatCode = '" & txtselectedcode & "' "
    
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Department Code not found", vbCritical)
            Else
                 txtdesc = pr_dumy("description")
                 txtclasscode.SetFocus
            End If
         pr_dumy.Close
    End If
End If
End Sub

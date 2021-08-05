VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmicreport7 
   Caption         =   "Stock Ledger Balance"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmicreports7.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2565
      Width           =   6765
      _ExtentX        =   11933
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
            Object.Width           =   105833
            MinWidth        =   105833
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
   Begin VB.Frame Frame1 
      Height          =   2595
      Left            =   30
      TabIndex        =   4
      Top             =   -45
      Width           =   6675
      Begin Crystal.CrystalReport crrpt 
         Left            =   3255
         Top             =   1005
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
      Begin VB.Frame Frame3 
         ForeColor       =   &H00000080&
         Height          =   2205
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   6675
         Begin VB.ComboBox txtcustomReport 
            Height          =   330
            ItemData        =   "frmicreports7.frx":030A
            Left            =   1680
            List            =   "frmicreports7.frx":0320
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1740
            Width           =   4920
         End
         Begin VB.TextBox txtitemcode 
            Height          =   315
            Left            =   1680
            MaxLength       =   6
            TabIndex        =   16
            Top             =   570
            Width           =   690
         End
         Begin VB.TextBox txtitemdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   2715
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   570
            Width           =   3840
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   2355
            Picture         =   "frmicreports7.frx":03A2
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   570
            Width           =   315
         End
         Begin VB.CommandButton Command5 
            Height          =   315
            Left            =   2370
            Picture         =   "frmicreports7.frx":0514
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   210
            Width           =   315
         End
         Begin VB.TextBox txtdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   2715
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   210
            Width           =   3840
         End
         Begin VB.TextBox txtselectedcode 
            Height          =   315
            Left            =   1680
            TabIndex        =   10
            Top             =   210
            Width           =   690
         End
         Begin MSComCtl2.DTPicker dtpfrom 
            Height          =   315
            Left            =   1680
            TabIndex        =   1
            Top             =   945
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   16449537
            CurrentDate     =   37293
         End
         Begin MSComCtl2.DTPicker DTPTo 
            Height          =   315
            Left            =   1680
            TabIndex        =   8
            Top             =   1335
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   16449537
            CurrentDate     =   37293
         End
         Begin VB.Label txtselectiveaccount 
            Height          =   300
            Left            =   3855
            TabIndex        =   20
            Top             =   930
            Visible         =   0   'False
            Width           =   2520
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Report Type :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   600
            TabIndex        =   19
            Top             =   1785
            Width           =   1230
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Item Code :"
            Height          =   210
            Left            =   720
            TabIndex        =   17
            Top             =   600
            Width           =   795
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Department Code :"
            Height          =   210
            Left            =   210
            TabIndex        =   13
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "To Date :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   915
            TabIndex        =   9
            Top             =   1365
            Width           =   645
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "From Date :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   735
            TabIndex        =   6
            Top             =   975
            Width           =   825
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   5550
         TabIndex        =   3
         Top             =   2205
         Width           =   1035
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4440
         MaskColor       =   &H00000000&
         TabIndex        =   2
         Top             =   2205
         Width           =   1050
      End
      Begin VB.TextBox txtVchrDesc 
         Height          =   315
         Left            =   435
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmicreport7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Pb_BlnkVchr As Boolean
Public PO_CODE As Object
Public PO_DESC As Object
Dim pr_dumy As New Recordset
Public codeid As String
Public Reporttype As String
Dim ls_sql As String





Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
On Error GoTo LocalErr

If txtcustomReport.ListIndex = 1 Then

ElseIf txtcustomReport.ListIndex = 1 Then
   
    With crrpt
        .SQLQuery = ""
        .ReportFileName = App.Path & Gs_ICRepoPath & "\StockledgerAmountSummarysep.RPT"
        
        .SQLQuery = " SELECT StockLedgerAmountSEP.PurchaseAmount, StockLedgerAmountSEP.PurchaseReturnAmount, StockLedgerAmountSEP.AdjInAmount,  StockLedgerAmountSEP.AdjOutAmount, StockLedgerAmountSEP.SaleAmount, StockLedgerAmountSEP.CostAmount, StockLedgerAmountSEP.SaleReturn,"
        .SQLQuery = .SQLQuery & " StockLedgerAmountSEP.OpeningAmount, IC_Item.catcode, IC_ItemCategory.Description FROM  StockLedgerAmountSEP StockLedgerAmountSEP LEFT OUTER JOIN   IC_Item IC_Item ON StockLedgerAmountSEP.Compcode = IC_Item.Compcode AND"
        .SQLQuery = .SQLQuery & " StockLedgerAmountSEP.ItemCode = IC_Item.ItemCode LEFT OUTER JOIN     IC_ItemCategory IC_ItemCategory ON IC_Item.Compcode = IC_ItemCategory.Compcode AND IC_Item.catcode = IC_ItemCategory.CatCode"
        .SQLQuery = .SQLQuery & " where StockLedgerAmountSEP.Compcode = '" & Gs_compcode & "' and StockLedgerAmountsep.transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "'  and StockLedgerAmountsep.transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' "

        If txtselectedcode <> "" Then
         .SQLQuery = .SQLQuery & " and IC_ItemCategory.CatCode = '" & txtselectedcode & "'"
        End If
        
        If txtitemcode <> "" Then
         .SQLQuery = .SQLQuery & " and IC_Item.itemcode = '" & txtitemcode & "'"
        End If
        
        
        .SQLQuery = .SQLQuery & " Union all "

        .SQLQuery = .SQLQuery & " SELECT 0  as PurchaseAmount, 0 as PurchaseReturnAmount,0 as AdjInAmount,  0 as AdjOutAmount, 0 as SaleAmount, 0 as CostAmount,0 as SaleReturn,"
        .SQLQuery = .SQLQuery & " (StockLedgerAmountSEP.PurchaseAmount+StockLedgerAmountSEP.AdjInAmount+StockLedgerAmountSEP.SaleReturn)-(StockLedgerAmountSEP.PurchaseReturnAmount+StockLedgerAmountSEP.AdjOutAmount+StockLedgerAmountSEP.CostAmount) as OpeningAmount, IC_Item.catcode, IC_ItemCategory.Description FROM  StockLedgerAmountSEP StockLedgerAmountSEP LEFT OUTER JOIN   IC_Item IC_Item ON StockLedgerAmountSEP.Compcode = IC_Item.Compcode AND"
        .SQLQuery = .SQLQuery & " StockLedgerAmountSEP.ItemCode = IC_Item.ItemCode LEFT OUTER JOIN     IC_ItemCategory IC_ItemCategory ON IC_Item.Compcode = IC_ItemCategory.Compcode AND IC_Item.catcode = IC_ItemCategory.CatCode"
        .SQLQuery = .SQLQuery & " where StockLedgerAmountSEP.Compcode = '" & Gs_compcode & "' and StockLedgerAmountsep.transdate < '" & Format(dtpfrom, "YYYY/MM/DD") & "' "

        If txtselectedcode <> "" Then
         .SQLQuery = .SQLQuery & " and IC_Item.catcode= '" & txtselectedcode & "'"
        End If
        
        
        If txtitemcode <> "" Then
         .SQLQuery = .SQLQuery & " and IC_Item.itemcode = '" & txtitemcode & "'"
        End If
        
        ' .SQLQuery = .SQLQuery & " Group by IC_Item.catcode , IC_ItemCategory.Description"
      
        
        .WindowTitle = "" & Me.Caption & ""
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = 'Stock Ledger Balance'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
        .PageZoom 120
    End With
    
ElseIf txtcustomReport.ListIndex = 4 Then
   
    With crrpt
        .SQLQuery = ""
        .ReportFileName = App.Path & Gs_ICRepoPath & "\Categorywisesalepurchase.rpt"
        
        .SQLQuery = " SELECT StockLedgerAmountSEP.PurchaseAmount, StockLedgerAmountSEP.PurchaseReturnAmount, StockLedgerAmountSEP.AdjInAmount,  StockLedgerAmountSEP.AdjOutAmount, StockLedgerAmountSEP.SaleAmount, StockLedgerAmountSEP.CostAmount, StockLedgerAmountSEP.SaleReturn,"
        .SQLQuery = .SQLQuery & " StockLedgerAmountSEP.OpeningAmount, IC_Item.catcode, IC_ItemCategory.Description FROM  StockLedgerAmountSEP StockLedgerAmountSEP LEFT OUTER JOIN   IC_Item IC_Item ON StockLedgerAmountSEP.Compcode = IC_Item.Compcode AND"
        .SQLQuery = .SQLQuery & " StockLedgerAmountSEP.ItemCode = IC_Item.ItemCode LEFT OUTER JOIN     IC_ItemCategory IC_ItemCategory ON IC_Item.Compcode = IC_ItemCategory.Compcode AND IC_Item.catcode = IC_ItemCategory.CatCode"
        .SQLQuery = .SQLQuery & " where StockLedgerAmountSEP.Compcode = '" & Gs_compcode & "' and StockLedgerAmountsep.transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "'  and StockLedgerAmountsep.transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' "
        
        If txtselectedcode <> "" Then
            If txtselectedcode = "Selective" Then
            .SQLQuery = .SQLQuery & "  and IC_ItemCategory.CatCode  in (" & txtselectiveaccount & ")"
            Else
            .SQLQuery = .SQLQuery & "  and IC_ItemCategory.CatCode  = '" & txtselectedcode & "'"
            End If
        End If

        If txtitemcode <> "" Then
         .SQLQuery = .SQLQuery & " and IC_Item.itemcode = '" & txtitemcode & "'"
        End If
        .WindowTitle = "" & Me.Caption & ""
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = 'Sale Purchase Report'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
        .PageZoom 120
    End With
    

ElseIf txtcustomReport.ListIndex = 2 Or txtcustomReport.ListIndex = 3 Then
   
        Call ChkTempTables("Tmp_Stockledger", True)
        ls_sql = " SELECT StockLedgerAmount.Compcode, StockLedgerAmount.TransDate, StockLedgerAmount.Remarks, "
        ls_sql = ls_sql & " StockLedgerAmount.ItemCode, StockLedgerAmount.ReceiveQty, StockLedgerAmount.IssueQty, StockLedgerAmount.Rate, StockLedgerAmount.Amount,"
        ls_sql = ls_sql & " IC_Item.CatCode into Tmp_Stockledger FROM StockLedgerAmount INNER JOIN"
        ls_sql = ls_sql & " IC_Item ON StockLedgerAmount.Compcode = IC_Item.Compcode AND StockLedgerAmount.ItemCode = IC_Item.ItemCode "
        ls_sql = ls_sql & "  where StockLedgerAmount.Compcode = '" & Gs_compcode & "' and StockLedgerAmount.transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "'  and StockLedgerAmount.transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' "
        
        If txtselectedcode <> "" Then
        ls_sql = ls_sql & " and IC_Item.catcode= '" & txtselectedcode & "'"
        End If
    
        If txtitemcode <> "" Then
        ls_sql = ls_sql & "  and IC_Item.itemcode = '" & txtitemcode & "'"
        End If
        
        gc_dbcon.Execute ls_sql
    
     'Opening balance
         Call ChkTempTables("Tmp_StockledgerOPAmount", True)
         ls_sql = "SELECT StockLedgerAmount.Compcode, IC_Item.itemCode,'" & Format(dtpfrom, "DD/MM/YYYY") & "' as TransDate, 'OP' AS printtranscode, 'Opening Balance' AS Remarks, SUM(StockLedgerAmount.Amount)"
         ls_sql = ls_sql & " AS OpeningAmount,SUM(StockLedgerAmount.ReceiveQty)-SUM(StockLedgerAmount.IssueQty) as OpeningQty into Tmp_StockledgerOPAmount FROM StockLedgerAmount INNER JOIN IC_Item ON StockLedgerAmount.Compcode = IC_Item.Compcode AND StockLedgerAmount.ItemCode = IC_Item.ItemCode INNER JOIN"
         ls_sql = ls_sql & " IC_ItemCategory ON IC_Item.Compcode = IC_ItemCategory.Compcode AND IC_Item.CatCode = IC_ItemCategory.CatCode"
         ls_sql = ls_sql & " where StockLedgerAmount.Compcode = '" & Gs_compcode & "' and StockLedgerAmount.transdate <'" & Format(dtpfrom, "YYYY/MM/DD") & "' "
    
         If txtselectedcode <> "" Then
           ls_sql = ls_sql & " and IC_ItemCategory.Catcode = '" & txtselectedcode & "'"
         End If
    
         If txtitemcode <> "" Then
            ls_sql = ls_sql & "  and IC_Item.itemcode = '" & txtitemcode & "'"
         End If
    
         ls_sql = ls_sql & " GROUP BY StockLedgerAmount.Compcode, IC_Item.itemcode"
         gc_dbcon.Execute ls_sql

        ls_sql = "insert into Tmp_StockledgerOPAmount (Compcode , itemCode, TransDate, printtranscode, Remarks, OpeningAmount)"
        ls_sql = ls_sql & " SELECT Compcode, itemCode, '" & Format(dtpfrom, "DD/MM/YYYY") & "' AS TransDate, 'OP' AS printtranscode, 'Opening Balance' AS Remarks, 0 AS OpeningAmount"
        ls_sql = ls_sql & "  From Tmp_Stockledger WHERE ((LTRIM(RTRIM(Compcode)) + itemCode) NOT IN  (SELECT     LTRIM(RTRIM(Compcode)) + itemcode FROM  Tmp_StockledgerOPAmount )) group by compcode,itemcode"
        gc_dbcon.Execute ls_sql
    
        If txtselectedcode <> "" Then
        'ls_sql = "Delete from Tmp_StockledgerOPAmount WHERE itemcode <> '" & txtitemcode & "'"
       ' gc_dbcon.Execute ls_sql
        End If
        
       ' ls_Sql = "Delete from Tmp_StockledgerOPAmount Where (OpeningAmount = 0)"
        'gc_dbcon.Execute ls_Sql
        
        
    With crrpt
        .SQLQuery = ""
        If txtcustomReport.ListIndex = 3 Then
        .ReportFileName = App.Path & Gs_ICRepoPath & "\StockledgerAmountItemGWisesum.RPT"
        Else
        .ReportFileName = App.Path & Gs_ICRepoPath & "\StockledgerAmountItemGWise.RPT"
        End If
        .WindowTitle = "" & Me.Caption & ""
       
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = 'Stock Ledger Balance'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
    End With
    
ElseIf txtcustomReport.ListIndex = 0 Or txtcustomReport.ListIndex = 5 Then


    If txtcustomReport.ListIndex = 5 Then
        Call ChkTempTables("Tmp_Stockledger", True)
        ls_sql = " SELECT StockLedgerAmount.Compcode, StockLedgerAmount.TransDate, StockLedgerAmount.TransCode, StockLedgerAmount.Remarks, "
        ls_sql = ls_sql & " StockLedgerAmount.ItemCode, StockLedgerAmount.ReceiveQty, StockLedgerAmount.IssueQty, StockLedgerAmount.Rate, StockLedgerAmount.Amount,"
        ls_sql = ls_sql & " IC_Item.catCode into Tmp_Stockledger FROM StockLedgerAmount INNER JOIN"
        ls_sql = ls_sql & " IC_Item ON StockLedgerAmount.Compcode = IC_Item.Compcode AND StockLedgerAmount.ItemCode = IC_Item.ItemCode "
        ls_sql = ls_sql & "  where StockLedgerAmount.Compcode = '" & Gs_compcode & "' and StockLedgerAmount.transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "'  and StockLedgerAmount.transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' "
        
        If txtselectedcode <> "" Then
        ls_sql = ls_sql & " and IC_Item.Catcode = '" & txtselectedcode & "'"
        End If
    
        If txtitemcode <> "" Then
        ls_sql = ls_sql & "  and IC_Item.itemcode = '" & txtitemcode & "'"
        End If
        
        
        gc_dbcon.Execute ls_sql
    End If
    
    
     'Opening balance
    Call ChkTempTables("Tmp_StockledgerOPAmount", True)
    ls_sql = "SELECT StockLedgerAmount.Compcode, IC_ItemCategory.catCode,'" & Format(dtpfrom, "DD/MM/YYYY") & "' as TransDate, 'OP' AS transcode, 'Opening Balance' AS Remarks, SUM(StockLedgerAmount.Amount)"
    ls_sql = ls_sql & " AS OpeningAmount into Tmp_StockledgerOPAmount FROM StockLedgerAmount INNER JOIN IC_Item ON StockLedgerAmount.Compcode = IC_Item.Compcode AND StockLedgerAmount.ItemCode = IC_Item.ItemCode INNER JOIN"
    ls_sql = ls_sql & " IC_ItemCategory ON IC_Item.Compcode = IC_ItemCategory.Compcode AND IC_Item.catCode = IC_ItemCategory.catCode"
    
    
    If txtcustomReport.ListIndex = 0 Then
        ls_sql = ls_sql & " where StockLedgerAmount.Compcode = '" & Gs_compcode & "' and StockLedgerAmount.transdate <'" & Format(dtpfrom, "YYYY/MM/DD") & "' "
    Else
        ls_sql = ls_sql & " where StockLedgerAmount.Compcode = '" & Gs_compcode & "' and StockLedgerAmount.transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' "
    End If
    
    If txtselectedcode <> "" Then
        ls_sql = ls_sql & " and IC_ItemCategory.catcode = '" & txtselectedcode & "'"
    End If
    
    ls_sql = ls_sql & " GROUP BY StockLedgerAmount.Compcode, IC_ItemCategory.catcode"
    gc_dbcon.Execute ls_sql

    ls_sql = "insert into Tmp_StockledgerOPAmount (Compcode , catCode, TransDate, transcode, Remarks, OpeningAmount)"
    ls_sql = ls_sql & " SELECT Compcode, catCode, '" & Format(dtpfrom, "DD/MM/YYYY") & "' AS TransDate, 'OP' AS transcode, 'Opening Balance' AS Remarks, 0 AS OpeningAmount"
    ls_sql = ls_sql & "  From IC_ItemCategory WHERE ((LTRIM(RTRIM(Compcode)) + catCode) NOT IN  (SELECT     LTRIM(RTRIM(Compcode)) + catcode FROM  Tmp_StockledgerOPAmount))"
    gc_dbcon.Execute ls_sql
    
    If txtselectedcode <> "" Then
        ls_sql = "Delete from Tmp_StockledgerOPAmount WHERE catCode <> '" & txtselectedcode & "'"
        gc_dbcon.Execute ls_sql
    End If
    
    
    With crrpt
          .SQLQuery = ""
       If txtcustomReport.ListIndex = 0 Then
           gc_dbcon.Execute "delete from Tmp_StockledgerOPAmount where openingamount = 0"
            .ReportFileName = App.Path & Gs_ICRepoPath & "\StockledgerAmountSummary.RPT"
         Else
            .ReportFileName = App.Path & Gs_ICRepoPath & "\StockledgerAmount.RPT"
         End If
        .WindowTitle = "" & Me.Caption & ""
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = 'Stock Ledger Balance'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
    End With
    MDIForm1.StatusBar1.Panels(7).Text = ""
End If
Exit Sub

LocalErr:

Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub Command1_Click()

    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtitemcode
    Set PO_DESC = txtitemdesc
    Gs_SQL = "SELECT itemCode, Description from IC_Item "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by itemcode"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' and catcode = '" & txtselectedcode & "' "
 
    MyLookupOLDB.Caption = "Items "
    MyLookupOLDB.Show 1
   If txtitemcode <> "" Then Call txtItemcode_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtselectiveaccount
    Set PO_DESC = txtdesc
    txtselectedcode = "Selective"
    txtdesc = "Selective Category"
    Gs_SQL = "SELECT CatCode, Description  FROM IC_ItemCategory "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' "
    MyLookupOLDBSelective.txtsearchbase.Clear
    MyLookupOLDBSelective.txtsearchbase.AddItem "CatCode"
    MyLookupOLDBSelective.txtsearchbase.AddItem "Description"
    MyLookupOLDBSelective.txtsearchbase.ListIndex = 1
    MyLookupOLDBSelective.Caption = "Item Category"
    MyLookupOLDBSelective.Show 1
    
End Sub



Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    DTPTo.SetFocus
End If
End Sub
Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    cmdGenerate.SetFocus
End If
End Sub

Private Sub Form_Load()
  
  dtpfrom = Date
  DTPTo = Date
  
End Sub

Private Sub txtItemcode_Change()
If txtitemcode = "" Then
txtitemdesc = ""
End If

End Sub

Private Sub txtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
If txtitemcode <> "" And KeyCode = vbKeyReturn Then
    txtitemcode = DoPad(txtitemcode, txtitemcode.MaxLength)
    ls_sql = "Select itemcode,Description from IC_Item where compcode = '" & Gs_compcode & "' and itemcode = '" & txtitemcode & "' "
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Item Code not found", vbCritical)
            Else
                txtitemdesc = pr_dumy("description")
               dtpfrom.SetFocus
            End If
         pr_dumy.Close

End If
End Sub

Private Sub txtselectedcode_Change()
If txtselectedcode = "" Then
txtdesc = ""
End If
End Sub

Private Sub txtselectedcode_KeyDown(KeyCode As Integer, Shift As Integer)
If txtselectedcode <> "" And KeyCode = vbKeyReturn Then
    txtselectedcode = DoPad(txtselectedcode, txtselectedcode.MaxLength)
    ls_sql = "Select Catcode,Description from IC_ItemCategory where compcode = '" & Gs_compcode & "' and Catcode = '" & txtselectedcode & "' "
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Material Code not found", vbCritical)
            Else
                txtdesc = pr_dumy("description")
               txtitemcode.SetFocus
            End If
         pr_dumy.Close

End If
End Sub

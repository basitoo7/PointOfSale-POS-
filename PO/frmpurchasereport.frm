VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPOPurchaseReport 
   Caption         =   "Purchase Report"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5490
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
   Icon            =   "frmpurchasereport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   5490
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox txttype 
      Height          =   330
      ItemData        =   "frmpurchasereport.frx":030A
      Left            =   1185
      List            =   "frmpurchasereport.frx":031A
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   885
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.CheckBox ChkDetail 
      Caption         =   "Detail :"
      Height          =   210
      Left            =   30
      TabIndex        =   8
      Top             =   930
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   4395
      TabIndex        =   7
      Top             =   855
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
      Left            =   3375
      TabIndex        =   6
      Top             =   855
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
      Height          =   900
      Left            =   30
      TabIndex        =   1
      Top             =   -45
      Width           =   5445
      Begin MSComCtl2.DTPicker dtpto 
         Height          =   315
         Left            =   1155
         TabIndex        =   2
         Top             =   510
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   54394881
         CurrentDate     =   37309
      End
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   1155
         TabIndex        =   3
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   54394881
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "From Date :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   285
         TabIndex        =   5
         Top             =   195
         Width           =   825
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "To Date :"
         Height          =   210
         Left            =   450
         TabIndex        =   4
         Top             =   540
         Width           =   645
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1260
      Width           =   5490
      _ExtentX        =   9684
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
Attribute VB_Name = "frmPOPurchaseReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PR_Item As New Recordset
Public PO_DESC As Object
Public PO_CODE As Object

Dim ls_sql As String

Private Sub cmdGenerate_Click()
On Error GoTo LocalErr
If Me.Caption = "Income/Expense Report" And ChkDetail.Value = 1 Then
With rptLedger
        .SQLQuery = ""
        .DiscardSavedData = True
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "IncomeexpenseRPTDetail.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = 'Income/Expense Report Detail'"
        .SQLQuery = "Transdate, SaleAmount, Payments, Expense, Drawing, Remarks, Accountno FROM   IncomeExpenseRPTDetail"
        .SQLQuery = .SQLQuery & " where compcode = '" & Gs_compcode & "' and Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpfrom, "YYYY/MM/DD") & "' "
        .SQLQuery = .SQLQuery & " and transid = " & txttype.ListIndex + 1 & ""
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
        
End With
Exit Sub
End If

If Me.Caption = "Income/Expense Report" Then
With rptLedger
        .SQLQuery = ""
        .DiscardSavedData = True
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "IncomeexpenseRPT.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = 'Income/Expense Report'"
        .SQLQuery = "SELECT Transdate, SaleAmount, Payments, Expense, Drawing FROM IncomeExpenseRPTSum IncomeExpenseRPTSum"
        .SQLQuery = .SQLQuery & " where compcode = '" & Gs_compcode & "' and Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpfrom, "YYYY/MM/DD") & "' "
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
        
End With
Exit Sub
End If

With rptLedger
        .DiscardSavedData = True
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "PurchaseReportsummaryDept.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = 'Purchase Report'"
        .SQLQuery = ""
        .SQLQuery = " SELECT PurchaseReport.Amount, PurchaseReport.DiscAmount, PurchaseReport.GSTAmount, PurchaseReport.Pqty, PurchaseReport.PrAmount,"
        .SQLQuery = .SQLQuery & " PurchaseReport.PrQty, IC_Item.CatCode, IC_ItemCategory.Description FROM  PurchaseReport PurchaseReport LEFT OUTER JOIN"
        .SQLQuery = .SQLQuery & " IC_Item IC_Item ON PurchaseReport.Compcode = IC_Item.Compcode AND PurchaseReport.ItemCode = IC_Item.ItemCode LEFT OUTER JOIN"
        .SQLQuery = .SQLQuery & " IC_ItemCategory IC_ItemCategory ON IC_Item.Compcode = IC_ItemCategory.Compcode AND IC_Item.CatCode = IC_ItemCategory.CatCode"
        .SQLQuery = .SQLQuery & " where PurchaseReport.compcode = '" & Gs_compcode & "'  "
        .SQLQuery = .SQLQuery & " and PurchaseReport.transdate >= '" & Format(dtpfrom.Value, "YYYY/MM/DD") & "' "
        .SQLQuery = .SQLQuery & " and  PurchaseReport.transdate <= '" & Format(DTPTo.Value, "YYYY/MM/DD") & "' "
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
End With
Exit Sub
LocalErr:
Call SetErr(Err.Description, vbCritical)
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Private Sub Command3_Click()
Unload Me
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
End Sub


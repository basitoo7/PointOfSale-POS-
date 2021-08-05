VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSoSalePurchaseReportCat 
   Caption         =   "Purchase Sale Report Category Wise "
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
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
   Icon            =   "frmSoSalePurchaseReportCat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   4740
      TabIndex        =   7
      Top             =   1440
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
      Left            =   3720
      TabIndex        =   6
      Top             =   1440
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
      Height          =   1425
      Left            =   30
      TabIndex        =   1
      Top             =   -45
      Width           =   5730
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2235
         Picture         =   "frmSoSalePurchaseReportCat.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox txtitemdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2565
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   270
         Width           =   3045
      End
      Begin VB.TextBox txtitemcode 
         Height          =   315
         Left            =   1155
         TabIndex        =   10
         Top             =   270
         Width           =   1050
      End
      Begin VB.CheckBox chkcash 
         Caption         =   "Cash"
         Height          =   270
         Left            =   4680
         TabIndex        =   9
         Top             =   1080
         Width           =   1020
      End
      Begin VB.CheckBox chkcredit 
         Caption         =   "Credit"
         Height          =   270
         Left            =   3870
         TabIndex        =   8
         Top             =   1080
         Width           =   1020
      End
      Begin MSComCtl2.DTPicker dtpto 
         Height          =   315
         Left            =   1155
         TabIndex        =   2
         Top             =   1005
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   64290817
         CurrentDate     =   37309
      End
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   1155
         TabIndex        =   3
         Top             =   645
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   64290817
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
      Begin VB.Label txtselectiveaccount 
         Height          =   300
         Left            =   2745
         TabIndex        =   14
         Top             =   675
         Visible         =   0   'False
         Width           =   2520
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Item Code :"
         Height          =   210
         Left            =   180
         TabIndex        =   13
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "From Date :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   285
         TabIndex        =   5
         Top             =   690
         Width           =   825
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "To Date :"
         Height          =   210
         Left            =   450
         TabIndex        =   4
         Top             =   1035
         Width           =   645
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1875
      Width           =   5805
      _ExtentX        =   10239
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
Attribute VB_Name = "frmSoSalePurchaseReportCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PR_Item As New Recordset
Public PO_DESC As Object
Public PO_CODE As Object

Dim ls_sql As String

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
 
Call ChkTempTables("Tmp_SalePurchase", True)

ls_sql = "SELECT PO_POGRNDetail.ItemCode, IC_Item.Description, SUM(PO_POGRNDetail.Quantity) AS Qty, 0 AS SaleAmount, SUM(PO_POGRNDetail.Amount) - SUM(PO_POGRNDetail.DiscAmount) AS Amount Into Tmp_SalePurchase "
ls_sql = ls_sql & " FROM  IC_Item INNER JOIN    PO_POGRNDetail ON IC_Item.ItemCode = PO_POGRNDetail.ItemCode"
ls_sql = ls_sql & " Inner Join IC_ItemCategory ON IC_Item.CatCode = IC_ItemCategory.CatCode"
ls_sql = ls_sql & " WHERE (SaleReport.TransDate >= '" & Format(dtpfrom, "YYYY/MM/DD") & "') AND (SaleReport.TransDate <=  '" & Format(dtpto, "YYYY/MM/DD") & "')"


If txtitemcode <> "" Then
    If txtitemcode = "Selective" Then
    ls_sql = ls_sql & "  and PO_POGRNDetail.customcode in (" & txtselectiveaccount & ")"
    Else
    ls_sql = ls_sql & "  and PO_POGRNDetail.customcode = '" & txtitemcode & "'"
    End If
End If



ls_sql = ls_sql & " GROUP BY PO_POGRNDetail.itemcode, IC_Item.Description"


ls_sql = ls_sql & " SELECT SaleReport.itemcode, IC_Item.Description, SUM(SaleReport.SaleQty) AS Qty, SUM(SaleReport.Amount) - SUM(SaleReport.DiscAmount) AS Amount,0 as Pamount "
ls_sql = ls_sql & " FROM SaleReport LEFT OUTER JOIN  IC_Item ON SaleReport.Compcode = IC_Item.Compcode AND SaleReport.itemcode = IC_Item.ItemCode"
ls_sql = ls_sql & " Inner Join IC_ItemCategory ON IC_Item.CatCode = IC_ItemCategory.CatCode"

ls_sql = ls_sql & " WHERE (SaleReport.TransDate >= '" & Format(dtpfrom, "YYYY/MM/DD") & "') AND (SaleReport.TransDate <=  '" & Format(dtpto, "YYYY/MM/DD") & "')"

If chkcredit.Value = 1 Then
ls_sql = ls_sql & " and (SaleReport.salestatus = 1)"
ElseIf chkcash.Value = 1 Then
ls_sql = ls_sql & "  and (SaleReport.salestatus <> 1)"
End If

If txtitemcode <> "" Then
    If txtitemcode = "Selective" Then
    ls_sql = ls_sql & "  and SaleReport.customcode in (" & txtselectiveaccount & ")"
    Else
    ls_sql = ls_sql & "  and SaleReport.customcode = '" & txtitemcode & "'"
    End If
End If


ls_sql = ls_sql & " GROUP BY SaleReport.itemcode, IC_Item.Description"

gc_dbcon.Execute ls_sql

With rptLedger
        .DiscardSavedData = True
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "Salereportitemwise.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = 'Sale Report'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & dtpto & "'"
        
        .Action = 1
End With


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
    Set PO_CODE = txtselectiveaccount
    Set PO_DESC = txtitemdesc
    txtitemcode = "Selective"
    txtitemdesc = "Selective Items"
    Gs_SQL = "SELECT customCode,Description,Salecost FROM IC_Item "
    Gs_FindFld = "customCode"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' "
    MyLookupOLDBSelective.txtsearchbase.Clear
    MyLookupOLDBSelective.txtsearchbase.AddItem "Customcode"
    MyLookupOLDBSelective.txtsearchbase.AddItem "Description"
    MyLookupOLDBSelective.txtsearchbase.ListIndex = 0
    MyLookupOLDBSelective.Caption = "Items."
    MyLookupOLDBSelective.Show 1
    
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdGenerate.SetFocus
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then dtpto.SetFocus
End Sub
Private Sub Form_Load()
dtpfrom = Date
dtpto = Date
txtpstype.ListIndex = 1

End Sub

Private Sub txtpstype_Click()
Me.Caption = txtpstype.Text
End Sub


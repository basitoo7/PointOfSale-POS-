VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmicreport2 
   Caption         =   "Stock Ledger Report"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmicreports2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5640
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      Height          =   960
      Left            =   15
      TabIndex        =   9
      Top             =   -75
      Width           =   5580
      Begin VB.TextBox txtbranchname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2460
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   15
         Tag             =   "SKIP"
         Top             =   180
         Width           =   3045
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2115
         Picture         =   "frmicreports2.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   180
         Width           =   315
      End
      Begin VB.TextBox txtbranchcode 
         Height          =   315
         Left            =   1545
         MaxLength       =   3
         TabIndex        =   13
         Top             =   210
         Width           =   525
      End
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   2235
         Picture         =   "frmicreports2.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   555
         Width           =   315
      End
      Begin VB.TextBox txtdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2565
         MaxLength       =   50
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   555
         Width           =   2940
      End
      Begin VB.TextBox txtselectedcode 
         Height          =   315
         Left            =   1545
         MaxLength       =   6
         TabIndex        =   0
         Top             =   570
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Branch Code :"
         Height          =   210
         Left            =   465
         TabIndex        =   16
         Top             =   225
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Item Code :"
         Height          =   210
         Left            =   690
         TabIndex        =   12
         Top             =   615
         Width           =   795
      End
   End
   Begin VB.Frame Frame3 
      ForeColor       =   &H00000080&
      Height          =   1005
      Left            =   15
      TabIndex        =   4
      Top             =   780
      Width           =   5580
      Begin VB.ComboBox txtgroupon 
         Height          =   330
         ItemData        =   "frmicreports2.frx":05EE
         Left            =   1545
         List            =   "frmicreports2.frx":05F8
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   570
         Width           =   1860
      End
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   1545
         TabIndex        =   6
         Top             =   180
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
         Format          =   63242241
         CurrentDate     =   37293
      End
      Begin Crystal.CrystalReport crrpt 
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
         Caption         =   "As on Date :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   600
         TabIndex        =   8
         Top             =   210
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Items Grouping On :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   75
         TabIndex        =   7
         Top             =   600
         Width           =   1425
      End
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
      Left            =   3465
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   1845
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4545
      TabIndex        =   2
      Top             =   1845
      Width           =   1035
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2220
      Width           =   5640
      _ExtentX        =   9948
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
End
Attribute VB_Name = "frmicreport2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pb_BlnkVchr As Boolean
Public PO_CODE As Object
Public PO_DESC As Object
Dim pr_dumy As New Recordset
Dim PR_Branch As New Recordset
Public codeid As String
Public Reporttype As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
'On Error GoTo LocalErr
Dim ls_sql As String
Dim ls_branchdesc As String

MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
If txtbranchname <> "" Then
    ls_branchdesc = "-(" + txtbranchname + ")"
Else
    ls_branchdesc = ""
End If

    With crrpt
        
        .ReportFileName = App.Path & Gs_ICRepoPath & "\Stockledger.RPT"
        .WindowTitle = Me.Caption
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = '" & Me.Caption + ls_branchdesc & "' "
        .Formulas(2) = "Period = '" & " As on date " & dtpfrom & "'"
        .Formulas(3) = "Groupon = " & txtgroupon.ListIndex + 1 & ""
        .RetrieveSQLQuery
    
        '.SQLQuery = "SELECT StockLedgerSummary.Qty, IC_Item.ItemCode, IC_Item.CustomCode, IC_Item.Description, IC_Item.ClassID, IC_Item.PurchaseCost, IC_Item.AvgRate,"
        '.SQLQuery = .SQLQuery & " IC_ItemClass.Description AS ItemClass, IC_ItemUM.Description AS UOM, IC_Sites.Description AS ItemSite FROM  StockLedgerSummary StockLedgerSummary INNER JOIN"
        '.SQLQuery = .SQLQuery & " IC_Item IC_Item ON StockLedgerSummary.compcode = IC_Item.Compcode AND StockLedgerSummary.ItemCode = IC_Item.ItemCode LEFT OUTER JOIN"
        '.SQLQuery = .SQLQuery & " IC_ItemUM IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode LEFT OUTER JOIN  IC_Sites IC_Sites ON IC_Item.Compcode = IC_Sites.CompCode AND IC_Item.SiteId = IC_Sites.SiteCode LEFT OUTER JOIN"
        '.SQLQuery = .SQLQuery & " IC_ItemClass IC_ItemClass ON IC_Item.Compcode = IC_ItemClass.Compcode AND IC_Item.ClassID = IC_ItemClass.ClassCode"
'
        .SQLQuery = .SQLQuery & " where StockLedgerSummary.Compcode = '" & Gs_compcode & "'"
        .SQLQuery = .SQLQuery & " and StockLedgerSummary.transdate <= '" & Format(dtpfrom, "YYYY/MM/DD") & "' "

        If Not Trim(txtselectedcode) = "" Then
        .SQLQuery = .SQLQuery & " and StockLedgerSummary.Itemcode = '" & txtselectedcode & "' "
        End If
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
    End With
MDIForm1.StatusBar1.Panels(7).Text = ""
Exit Sub

LocalErr:
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Private Sub Command5_Click()

    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtselectedcode
    Set PO_DESC = txtdesc
    
    Gs_SQL = "SELECT  Itemcode,Description FROM IC_Item "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    MyLookupOLDB.Caption = "Item Code"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    MyLookupOLDB.Caption = "Item Code"
    
    
    MyLookupOLDB.Show 1
    SendKeys "{Tab}"
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub Form_Load()
   PR_Branch.Open "Select * From SysBranch Where compcode = '" & Gs_compcode & "' Order By BranchCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
   txtbranchcode = Gs_BranchCode
  If MySeek(txtbranchcode, "Branchcode", PR_Branch) Then
   txtbranchname = PR_Branch("BranchDesc")
  End If
  dtpfrom = Date
  txtgroupon.ListIndex = 0

  
End Sub

Private Sub Form_Unload(Cancel As Integer)
PR_Branch.Close
End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbranchcode
    Set PO_DESC = txtbranchname
    
    GoTop PR_Branch
    MyLookup.Caption = "Company Branches"
    MyLookup.FillGrid PR_Branch, "BranchCode", "BranchDesc", txtbranchcode.MaxLength
    MyLookup.Show 1

    If Len(txtbranchcode) > 0 Then txtBranchCode_KeyDown vbKeyReturn, vbKeyShift
End Sub
Private Sub txtBranchCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) And txtbranchcode <> "" Then
     txtbranchcode = DoPad(txtbranchcode, txtbranchcode.MaxLength)
     
     If Not MySeek(txtbranchcode, "Branchcode", PR_Branch) Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtbranchcode.SetFocus
     Else
        txtbranchname = PR_Branch("BranchDesc")
        txtselectedcode.SetFocus
     End If
  ElseIf KeyCode = vbKeyF12 Then

     Command1_Click
  ElseIf KeyCode = vbKeyReturn And txtbranchcode = "" Then
      txtbranchname = ""
  End If
End Sub


Private Sub txtbranchcode_LostFocus()
If txtbranchcode = "" Then
      txtbranchname = ""
End If
End Sub

Private Sub txtselectedcode_LostFocus()
If Trim(txtselectedcode) = "" Then
    txtdesc = ""
End If
End Sub

Private Sub txtselectedcode_Validate(Cancel As Boolean)
If txtselectedcode <> "" Then
         txtselectedcode = DoPad(txtselectedcode, txtselectedcode.MaxLength)
         pr_dumy.Open "Select itemcode,Description from Ic_item where itemcode = '" & txtselectedcode & "' and compcode = '" & Gs_compcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Item code not found", vbCritical)
                'Cancel = True
            Else
                txtdesc = pr_dumy("description")
            End If
         pr_dumy.Close

ElseIf txtselectedcode = "" Then
        txtdesc = ""
End If
End Sub

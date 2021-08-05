VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmadjustment1 
   Caption         =   "Inventory Adjustments"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdjustment1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6225
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   0
      TabIndex        =   1
      Top             =   660
      Width           =   6180
      Begin MSComCtl2.DTPicker txtvaluedate 
         Height          =   315
         Left            =   4905
         TabIndex        =   33
         Top             =   150
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58327041
         CurrentDate     =   37580
      End
      Begin VB.TextBox txtAdjCode 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "XXX"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   900
         Width           =   405
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   1845
         Picture         =   "frmAdjustment1.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   915
         Width           =   315
      End
      Begin VB.TextBox txtAdjDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2175
         MaxLength       =   64
         TabIndex        =   29
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   900
         Width           =   3900
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3060
         MaxLength       =   50
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   180
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox TxtRemarks 
         Height          =   495
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1260
         Width           =   4635
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   1845
         Picture         =   "frmAdjustment1.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   540
         Width           =   315
      End
      Begin VB.TextBox txtLocCode 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "XXX"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   540
         Width           =   405
      End
      Begin VB.TextBox txtLocDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2175
         MaxLength       =   64
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   540
         Width           =   3900
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   315
         Left            =   2520
         Picture         =   "frmAdjustment1.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   180
         Width           =   315
      End
      Begin VB.TextBox txtTransNo 
         BackColor       =   &H00FFFF00&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "XXX"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Adjustment Type :"
         Height          =   255
         Left            =   60
         TabIndex        =   32
         Top             =   900
         Width           =   1305
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks :"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1260
         Width           =   1125
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Location Code :"
         Height          =   255
         Left            =   60
         TabIndex        =   22
         Top             =   540
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Reference #  :"
         Height          =   255
         Left            =   90
         TabIndex        =   21
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label label2 
         Caption         =   "Value Date :"
         Height          =   255
         Left            =   3810
         TabIndex        =   20
         ToolTipText     =   "Enter Value Date"
         Top             =   180
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2475
      Left            =   0
      TabIndex        =   2
      Top             =   2520
      Width           =   6180
      Begin VB.CommandButton Command6 
         Height          =   315
         Left            =   2430
         Picture         =   "frmAdjustment1.frx":0760
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox txtbatchno 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "XXX"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1500
         MaxLength       =   50
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Item Code"
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox txtItemDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         MaxLength       =   64
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2100
         Width           =   2535
      End
      Begin VB.TextBox TxtGrnTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4785
         MaxLength       =   11
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2100
         Width           =   1275
      End
      Begin VB.TextBox TxtUM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0;(""$""#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3195
         MaxLength       =   11
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   390
         Width           =   735
      End
      Begin VB.TextBox TxtFactor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0;(""$""#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2775
         MaxLength       =   11
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   390
         Width           =   435
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   1170
         Picture         =   "frmAdjustment1.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   375
         Width           =   315
      End
      Begin VB.TextBox TxtItemCode 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "XXX"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   390
         Width           =   1050
      End
      Begin VB.TextBox txtqty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   3960
         MaxLength       =   11
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   390
         Width           =   1095
      End
      Begin VB.TextBox txtUnitPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.000;(""$""#,##0.000)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         Height          =   315
         Left            =   5100
         MaxLength       =   11
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   390
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid GrdGRN 
         Height          =   1335
         Left            =   120
         TabIndex        =   16
         Top             =   780
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   2355
         _Version        =   393216
         Rows            =   1
      End
      Begin VB.Label Label14 
         Caption         =   "Batch No :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1545
         TabIndex        =   36
         Top             =   150
         Width           =   825
      End
      Begin VB.Label Label11 
         Caption         =   "Adjusment Total :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3465
         TabIndex        =   26
         Top             =   2160
         Width           =   1245
      End
      Begin VB.Label Label10 
         Caption         =   "U / M :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3495
         TabIndex        =   25
         Top             =   165
         Width           =   465
      End
      Begin VB.Label Label5 
         Caption         =   "Factor :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2775
         TabIndex        =   23
         Top             =   165
         Width           =   585
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Item Code :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   135
         TabIndex        =   19
         Top             =   150
         Width           =   1260
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Quantity :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3975
         TabIndex        =   18
         Top             =   165
         Width           =   1065
      End
      Begin VB.Label Label9 
         Caption         =   "Unit Price :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5220
         TabIndex        =   17
         Top             =   165
         Width           =   825
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   1005
      ButtonWidth     =   1217
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&New"
            Description     =   "Add"
            Object.ToolTipText     =   "Add new record"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Edit"
            Description     =   "Edit"
            Object.ToolTipText     =   "Edit an existing record"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Delete"
            Description     =   "Remove "
            Object.ToolTipText     =   "Remove an existing record."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Save"
            Description     =   "Save a new Record"
            Object.ToolTipText     =   "Save on disk"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Slip"
            Description     =   "Print Listing."
            Object.ToolTipText     =   "Print listing."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Find"
            Description     =   "Find a Record."
            Object.ToolTipText     =   "Find a record."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancel"
            Description     =   "Cancel Operation"
            Object.ToolTipText     =   "Cancel operation mode"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   14
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4920
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdjustment1.frx":0A44
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdjustment1.frx":0E98
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdjustment1.frx":12EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdjustment1.frx":1740
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdjustment1.frx":1B94
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdjustment1.frx":1FE8
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdjustment1.frx":273C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmadjustment1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim PB_BlnkAdj As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object

Dim PI_CurRow    As Integer
Dim PI_SrNo     As Integer
Dim PS_RowClicked As String

Dim cntsql As New ADODB.Command
Dim ls_TransType As String

Dim PR_ICItmLoc As New Recordset
Dim PR_ICAdJt As New Recordset
Dim PR_IcItem As New Recordset
Dim PR_ICAdjCode As New Recordset
Dim PR_ICGRNSUM As New Recordset
Dim PR_BatchQty As New Recordset

Private Sub cboIssueType_Click()
    cboIssueID.ListIndex = cboIssueType.ListIndex
End Sub

Private Sub cboIssueType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then TxtPartyCode.SetFocus
End Sub

Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtTransNo
    Set PO_DESC = Text1
    GoTop PR_ICAdJt
    MyLookup.Caption = "Adjustment References "
    MyLookup.FillGrid PR_ICAdJt, "Transc_No", "Value_Date", 10
    MyLookup.Show 1
    
    If Len(txtTransNo) > 0 Then txtTransNo_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtLocCode
    Set PO_DESC = txtLocDesc
    GoTop PR_ICItmLoc
    MyLookup.Caption = "Locations. "
    MyLookup.FillGrid PR_ICItmLoc, "LocationCode", "Description", 6
    MyLookup.Show 1
    
    If Len(txtLocCode) > 0 Then txtLocCode_KeyDown vbKeyReturn, vbKeyShift
End Sub


Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = TxtItemCode
    Set PO_DESC = txtItemDesc
    
    PR_IcItem.Filter = "LocationCode = '" & txtLocCode.Text & " '"
    If PR_IcItem.EOF Then
       Call SetErr("No Item has been found.", vbCritical)
       Exit Sub
    End If
    GoTop PR_IcItem
    MyLookup.Caption = "Items. "
    MyLookup.FillGrid PR_IcItem, "ItemId", "Description", 15
    MyLookup.Show 1
    PR_IcItem.Filter = adFilterNone
    
    If Len(TxtItemCode) > 0 Then txtItemCode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command4_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtAdjCode
    Set PO_DESC = txtAdjDesc
    
    GoTop PR_ICAdjCode
    MyLookup.Caption = "Receving Locations"
    MyLookup.FillGrid PR_ICAdjCode, "AdjCode", "Description", 5
    MyLookup.Show 1
    
    If Len(txtAdjCode) > 0 Then txtAdjCode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command6_Click()
  Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbatchno
    Set PO_DESC = Text1
    
    PR_BatchQty.Filter = "ItemCode = '" & TxtItemCode & "' and  Quantity > 0 "
    If PR_BatchQty.EOF Then
       Call SetErr("No Item has been found.", vbCritical)
       Exit Sub
    End If
    GoTop PR_BatchQty
    MyLookup.Caption = "Batch Qty. "
    MyLookup.FillGrid PR_BatchQty, "Batchno", "Quantity", 8
    MyLookup.Show 1
    PR_BatchQty.Filter = adFilterNone
    If Len(txtbatchno) > 0 Then txtbatchno_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Form_Load()
 cntsql.ActiveConnection = gc_dbcon
 cntsql.CommandType = adCmdText
 ls_TransType = "A"
 
  SetToolBar(1) = chkRights("ICADJSTP01")
  SetToolBar(2) = False
  SetToolBar(3) = False
  SetToolBar(4) = chkRights("ICADJSTP04")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  Call ChkTempTables("Tmp_GrnSm", True)
  gc_dbcon.Execute ("Select SUM(Round(Quantity*UnitCost,0)) As TGRNValue, SUM(Quantity) AS TGRNQty,LocationCode1+ItemClass+ItemCode as ItemGRN into Tmp_GRNSM from Ic_Trans where compcode ='" & Gs_compcode & "' and TransType = 'G' AND Value_Date >= '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' AND Value_Date <= '" & Format(Gs_FnEndPeriod, "YYYY/MM/DD") & "' group by locationcode1,itemclass,itemcode")
  
  PR_ICAdjCode.Open "Select * from Ic_AdjType", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_ICGRNSUM.Open "Select * From Tmp_GRNSM Order By ItemGRN", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_ICItmLoc.Open "Select * from Ic_Locations where compcode ='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_IcItem.Open "Select *,(ltrim(rtrim(itemclass))+ltrim(rtrim(itemcode))) as ItemID,(ltrim(rtrim(locationcode))+ltrim(rtrim(itemclass))+ltrim(rtrim(itemcode))) as ItemFind from Ic_Item where compcode ='" & Gs_compcode & "' order by itemfind", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_ICAdJt.Open "Select *,(Compcode+TransType+Transc_No) As AdjFind from Ic_Trans where compcode ='" & Gs_compcode & "' and TransType = 'A' AND (Value_Date >= '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' AND Value_Date <= '" & Format(Gs_FnEndPeriod, "YYYY/MM/DD") & "') order by AdjFind", gc_dbcon, adOpenDynamic, adLockOptimistic
  PR_BatchQty.Open "SELECT LTRIM(RTRIM(ItemClass)) + LTRIM(RTRIM(ItemCode)) AS ItemCode, BatchNo,LTRIM(RTRIM(ItemClass)) + LTRIM(RTRIM(ItemCode))+LTRIM(RTRIM(BatchNo)) AS Findfld, BatchNo, SUM(Quantity) AS Quantity From IC_Trans where compcode = '" & Gs_compcode & "' GROUP BY ItemClass,ItemCode,BatchNo ", gc_dbcon, adOpenStatic, adLockOptimistic, 1
  
  PB_BlnkAdj = IIf(PR_ICAdJt.EOF, True, False)
  InitializeGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_ICAdjCode.Close
    PR_ICItmLoc.Close
    PR_ICAdJt.Close
    PR_IcItem.Close
    PR_BatchQty.Close
    Set PR_ICGRNSUM = Nothing
    gc_dbcon.Execute ("Drop Table Tmp_GrnSm;")
End Sub

Private Sub Text1_GotFocus()
    txtvaluedate.SetFocus
End Sub

Private Sub txtbatchno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
        PR_BatchQty.Filter = "ItemCode = '" & TxtItemCode & "' and  Quantity > 0 "
        If Not MySeek(Trim(TxtItemCode.Text) + Trim(txtbatchno), "Findfld", PR_BatchQty) Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtbatchno.SetFocus
         Else
             ln_balqty = Val(0 & PR_BatchQty("Quantity"))
             txtqty.SetFocus
         End If
ElseIf KeyCode = vbKeyF12 Then
        Command6_Click
ElseIf KeyCode = vbKeyPageUp Then
        TxtItemCode.SetFocus
 End If
End Sub

'Private Sub txtIssueID_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim lb_found As Boolean
'
' If LastKey(KeyCode) And Len(txtIssueID.Text) > 0 Then
'         txtIssueID = UCase(txtIssueID)
'         lb_found = MySeek(txtIssueID.Text, "IssueID", PR_IssueType)
'         If Not lb_found Then
'             Call SetErr(Gs_RecNFMsg, vbCritical)
'             txtIssueID.SetFocus
'             txtIssueDesc.Text = ""
'         Else
'             txtIssueDesc.Text = PR_IssueType("Description")
'             txtJobCode.SetFocus
'         End If
' End If
'End Sub

'Private Sub txtJobCode_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim lb_found As Boolean
'
' If LastKey(KeyCode) And Len(txtJobCode.Text) > 0 Then
'         txtJobCode.Text = DoPad(txtJobCode.Text, 10)
'         lb_found = MySeek(txtJobCode.Text, "JobCode", PR_JobCode)
'         If Not lb_found Then
'             Call SetErr(Gs_RecNFMsg, vbCritical)
'             txtJobCode.SetFocus
'             txtJobDesc.Text = ""
'         Else
'             txtJobDesc.Text = PR_IssueType("Description")
'             TxtPartyCode.SetFocus
'         End If
' End If
'
'End Sub

Private Sub txtqty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If Not Range(Val(0 & txtqty.Text), Val(0 & PR_IcItem.Fields("ItemBalQty")) * (-1), Val(0 & PR_IcItem.Fields("ItemBalQty"))) Then
          Call SetErr("Invalid Quantity Value.", vbCritical)
          txtqty.SetFocus
       Else
          txtqty.Text = Round(txtqty.Text * TxtFactor, 0)
          Call txtUnitPrice_KeyDown(vbKeyReturn, vbKeyShift)
          TxtItemCode.SetFocus
       End If
    End If
End Sub

Private Sub txtAdjCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 
 If KeyCode = vbKeyReturn And Len(txtAdjCode.Text) > 0 Then
         txtAdjCode.Text = UCase(txtAdjCode.Text)
         lb_found = MySeek(txtAdjCode.Text, "AdjCode", PR_ICAdjCode)
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtAdjCode.SetFocus
             txtAdjDesc.Text = ""
         Else
             txtAdjDesc.Text = PR_ICAdjCode("Description")
             TxtRemarks.SetFocus
         End If
 ElseIf KeyCode = vbKeyF12 Then
        Command4_Click
 End If
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
   If Lastkey(KeyCode) Then
      TxtItemCode.SetFocus
   End If
End Sub

Private Sub txtTransNo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn And Len(txtTransNo.Text) > 0 Then
         
         txtTransNo.Text = IIf(IsNumeric(LTrim(Str(txtTransNo.Text))), DoPad(UCase(txtTransNo.Text), 10), UCase(txtTransNo.Text))
         lb_found = MySeek(Gs_compcode & ls_TransType & txtTransNo.Text, "AdjFind", PR_ICAdJt)
         
         If Gl_Demo Then
            If PR_ICAdJt.RecordCount > gn_MaxVchrs Then
                Call SetErr("This is Demo software you cannot add more than " & LTrim(Str(gn_MaxVchrs)) & " Transactions.", vbCritical)
                Unload Me
                Exit Sub
            End If
         End If
         InitializeGrid
         
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   Cancel = True
                   Call ClearVal
                   txtTransNo.SetFocus
                Else
                   txtLocCode.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   Cancel = True
                   Call ClearVal
                   txtLocCode.SetFocus
                Else
                   Call SetVal
                   LoadGRNTrans
                   If Mode <> "D" Then
                      txtTransNo.SetFocus
                   End If
                End If
            End Select
    ElseIf KeyCode = vbKeyF12 Then
        cmdLookup_Click
    End If
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
       cmdLookup.Enabled = False
    Else
       cmdLookup.Enabled = True
    End If
    If Button.Index = 7 Then InitializeGrid
    
    If PB_BlnkAdj And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_ICAdJt, Me, txtTransNo, txtvaluedate, Para_Rs, "IC_AdjCnt", 10, "txtTransNo", "text1", 0, False, Toolbar1)
    End If
    
End Sub


Public Sub SaveValues()
Dim cntsql As New ADODB.Command
Dim Ln_Cnt As Integer
Dim ln_AvgPrice As Integer

PB_BlnkAdj = False

cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText

On Error GoTo RollBack
gc_dbcon.BeginTrans

     Select Case Mode
           Case "D"
'              cntsql.CommandText = "DELETE FROM IC_Trans WHERE CompCode = '" & Gs_compcode & "' AND Transc_No = '" & Trim(txtTransNo) & "' AND TransType = '" & ls_TransType & "' AND Value_Date = '" & Format(txtvaluedate, "YYYY/MM/DD") & "' AND LocationCode1 = '" & Trim(txtLocCode) & "'"
'              cntsql.Execute
           Case Else
'                If Mode = "E" Then
'                    cntsql.CommandText = "DELETE FROM IC_Trans WHERE CompCode = '" & Gs_compcode & "' AND Transc_No = '" & Trim(txtTransNo) & "' AND TransType = '" & ls_TransType & "' AND Value_Date = '" & Format(txtvaluedate, "YYYY/MM/DD") & "' AND RTrim(LTrim(LocationCode1)) = '" & Trim(txtLocCode) & "'"
'                    cntsql.Execute
'                End If
              
                With GrdGRN
                    For Ln_Cnt = 1 To .Rows - 1
                      If txtAdjCode <> "G" Then
                           'Transaction and Ledger Posting of Issuance
                            If MySeek(txtLocCode.Text & .TextMatrix(Ln_Cnt, 1), "ItemFind", PR_IcItem) Then
                               PR_IcItem.Fields("ItemBalQty") = PR_IcItem.Fields("ItemBalQty") - Val(.TextMatrix(Ln_Cnt, 2))
                               PR_IcItem.Update
                            End If
                            cntsql.CommandText = "INSERT into IC_Trans(compcode,Transc_No, TransType, Value_Date, LocationCode1, SerialNo, ItemClass, ItemCode,batchno,Quantity,UnitCost,Remarks) VALUES ('" & Gs_compcode & "','" & Trim(txtTransNo) & "','" & ls_TransType & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & Trim(txtLocCode) & "'," & .TextMatrix(Ln_Cnt, 0) & ",'" & .TextMatrix(Ln_Cnt, 6) & "','" & .TextMatrix(Ln_Cnt, 7) & "'," & Val(.TextMatrix(Ln_Cnt, 2)) * -1 & "," & Val(.TextMatrix(Ln_Cnt, 3)) & ",'" & Trim(TxtRemarks) & "')"
                            cntsql.Execute
                       Else
                            ' Updation of Value & Quantity in GRN Ledgers
                            If MySeek(txtLocCode.Text & .TextMatrix(Ln_Cnt, 1), "ItemGRN", PR_ICGRNSUM) Then
                               PR_ICGRNSUM.Fields("TGRNValue") = PR_ICGRNSUM.Fields("TGRNValue") + Val(.TextMatrix(Ln_Cnt, 4))
                               PR_ICGRNSUM.Fields("TGRNQty") = PR_ICGRNSUM.Fields("TGRNQty") + Val(.TextMatrix(Ln_Cnt, 2))
                            End If
                            
                            'Transaction and Ledger Posting of Receipt
                            If MySeek(txtLocCode.Text & .TextMatrix(Ln_Cnt, 1), "ItemFind", PR_IcItem) Then
                               ln_OpnQty = PR_IcItem.Fields("ItemOpenQty")
                               ln_OpnValue = IIf(ln_OpnQty <= 0, 0, PR_IcItem.Fields("ItemOpenValue"))
                               
                               If PR_ICGRNSUM.EOF Then
                                  ln_AvgPrice = Round((ln_OpnValue + Val(.TextMatrix(Ln_Cnt, 4))) / (Val(.TextMatrix(Ln_Cnt, 2)) + ln_OpnQty), 2)
                               Else
                                  ln_AvgPrice = Round((ln_OpnValue + PR_ICGRNSUM.Fields("TGRNValue")) / (ln_OpnQty + PR_ICGRNSUM.Fields("TGRNQty")), 2)
                               End If
                               PR_IcItem.Fields("ItemAvgPrice") = ln_AvgPrice
                               PR_IcItem.Fields("ItemUnitPrice") = Val(.TextMatrix(Ln_Cnt, 3))
                               PR_IcItem.Fields("ItemBalQty") = PR_IcItem.Fields("ItemBalQty") + Val(.TextMatrix(Ln_Cnt, 2))
                               PR_IcItem.Update
                            Else
                               Call SetErr("Receiving Location could not be found, Process Terminated.", vbCritical)
                               GoTo RollBack
                            End If
                            cntsql.CommandText = "INSERT into IC_Trans(compcode,Transc_No, TransType, Value_Date, LocationCode1, SerialNo, ItemClass, ItemCode,Quantity,UnitCost,Remarks) VALUES ('" & Gs_compcode & "','" & Trim(txtTransNo) & "','" & ls_TransType & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & Trim(txtLocCode) & "'," & .TextMatrix(Ln_Cnt, 0) & ",'" & .TextMatrix(Ln_Cnt, 6) & "','" & .TextMatrix(Ln_Cnt, 7) & "'," & Val(.TextMatrix(Ln_Cnt, 2)) & "," & Val(.TextMatrix(Ln_Cnt, 3)) & ",'" & Trim(TxtRemarks) & "')"
                            cntsql.Execute
                    End If
                   Next
                 txtTransNo.Text = DoPad(LTrim(Str(Para_Rs.Fields("IC_AdjCnt") + 1)), 10)
                 End With
                 
     End Select
gc_dbcon.CommitTrans
PR_ICAdJt.Requery
InitializeGrid
On Error GoTo 0
Exit Sub

RollBack:
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
gc_dbcon.RollbackTrans

If Mode = "A" Then
    Para_Rs.Fields("Ic_Adjcnt") = Para_Rs.Fields("Ic_Adjcnt") - 1
    Para_Rs.Update
End If
On Error GoTo 0
End Sub
Public Sub ClearVal()
     txtLocDesc = ""
     txtLocCode = ""
     txtAdjCode = ""
     txtAdjDesc = ""
     TxtRemarks = ""
     txtvaluedate.Value = Date
     TxtGrnTotal = ""
     TxtItemCode = ""
     InitializeGrid
     TxtFactor = ""
     PI_SrNo = 0
     TxtUM = ""
End Sub

Private Sub SetVal()
     txtvaluedate = PR_ICAdJt("Value_Date")
     txtLocCode = PR_ICAdJt("LocationCode1")
     txtLocDesc = IIf(MySeek(PR_ICAdJt("LocationCode1"), "LocationCode", PR_ICItmLoc), PR_ICItmLoc.Fields("Description"), "Not Found.")
     txtAdjCode = PR_ICAdJt("IssueType") & ""
     txtAdjDesc = IIf(MySeek(PR_ICAdJt("IssueType"), "AdjCode", PR_ICAdjCode), PR_ICAdjCode.Fields("Description"), "Not Found.")
     TxtRemarks = PR_ICAdJt("Remarks")
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtTransNo) = txtTransNo.MaxLength And Len(txtLocCode) = txtLocCode.MaxLength And PI_SrNo > 0 Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Private Sub txtLocCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 
 If KeyCode = vbKeyReturn And Len(txtLocCode.Text) > 0 Then
         txtLocCode.Text = IIf(IsNumeric(LTrim(RTrim(txtLocCode.Text))), DoPad(txtLocCode.Text, txtLocCode.MaxLength), UCase(txtLocCode.Text))
         lb_found = MySeek(txtLocCode.Text, "LocationCode", PR_ICItmLoc)
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtLocCode.SetFocus
             txtLocDesc.Text = ""
         Else
             txtLocDesc.Text = PR_ICItmLoc("Description")
             txtAdjCode.SetFocus
         End If
 ElseIf KeyCode = vbKeyF12 Then
   Command2_Click
 End If
End Sub

Private Sub txtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
Dim Ln_Cnt As Integer
 
 If KeyCode = vbKeyReturn And Len(TxtItemCode.Text) > 0 Then
         With GrdGRN
               For Ln_Cnt = 1 To .Rows - 1
                   If .TextMatrix(Ln_Cnt, 1) = TxtItemCode Then
                      Call SetErr("Item ID already exists.", vbCritical)
                      TxtItemCode.SetFocus
                      Exit Sub
                   End If
                Next
         End With
         TxtItemCode.Text = UCase(TxtItemCode.Text)
         lb_found = MySeek(txtLocCode.Text & TxtItemCode.Text, "ItemFind", PR_IcItem)
         
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             TxtItemCode.SetFocus
         Else
             txtItemDesc = PR_IcItem.Fields("Description")
             TxtFactor = PR_IcItem.Fields("ItemFactor")
             TxtUM = PR_IcItem.Fields("ItemUM")
             txtunitprice = IIf(gs_ICBase = "A", Val(0 & PR_IcItem.Fields("ItemAvgPrice")), Val(0 & PR_IcItem.Fields("ItemUnitPrice")))
             txtbatchno.SetFocus
         End If
 ElseIf KeyCode = vbKeyF12 Then
        Command1_Click
 End If
End Sub

Private Sub InitializeGrid()
    With GrdGRN
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Item Code|<Quantity|<Unit Price|<Amount  |<Location|<Item Class|<Item Code"
        .ColWidth(1) = 1200
        .ColWidth(2) = 900
        .ColWidth(3) = 900
        .ColAlignment(3) = 7
        .ColWidth(4) = 1440
        .ColAlignment(4) = 7
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .Redraw = True
    End With
End Sub

Private Sub grdGrn_KeyDown(KeyCode As Integer, Shift As Integer)
    With GrdGRN
        If KeyCode = vbKeyDelete Then
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            .RemoveItem .Row
            TotalGRN
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                .TextMatrix(.Row, 0) = ""
                PI_SrNo = 0
            End If
        End If
    End With
End Sub

Private Sub txtUnitPrice_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Ln_Cnt As Integer

    If Lastkey(KeyCode) Then
      If (Val(txtqty) > 0 Or Val(txtunitprice) > 0) Then
        
        If PS_RowClicked = "" Then
            If PI_SrNo = 0 Then
                PI_SrNo = 1
            Else
                PI_SrNo = PI_SrNo + 1
            End If
        End If
        
        If Len(TxtItemCode.Text) > 0 Then
            With GrdGRN
                If PS_RowClicked = "" Then
                    If Not PI_SrNo = 1 Then .Rows = .Rows + 1
                    .Row = .Rows - 1
                Else
                    .Row = PI_CurRow
                End If
                If PS_RowClicked = "" Then
                    .TextMatrix(.Row, 0) = PI_SrNo
                Else
                    .TextMatrix(.Row, 0) = PI_CurRow
                End If
                .TextMatrix(.Row, 1) = Trim(TxtItemCode)
                .TextMatrix(.Row, 2) = Val(txtqty)
                .TextMatrix(.Row, 3) = Val(txtunitprice)
                .TextMatrix(.Row, 4) = Round(Val(txtqty) * Val(txtunitprice), 0)
                
                .TextMatrix(.Row, 5) = PR_IcItem("LocationCode")
                .TextMatrix(.Row, 6) = PR_IcItem("ItemClass")
                .TextMatrix(.Row, 7) = PR_IcItem("ItemCode")
                
                TxtItemCode.Text = ""
                TxtFactor.Text = ""
                TxtUM.Text = ""
                txtqty = ""
                txtunitprice = ""
                TxtItemCode.SetFocus
                PS_RowClicked = ""
            End With
        End If
            TotalGRN
    Else
               txtLocCode.SetFocus
    End If
    End If
End Sub

Private Sub TotalGRN()
    Dim Ln_Cnt As Integer
    TxtGrnTotal = ""
    With GrdGRN
        For Ln_Cnt = 1 To .Rows - 1
            .TextMatrix(Ln_Cnt, 0) = Ln_Cnt
            TxtGrnTotal = Val(TxtGrnTotal) + Val(.TextMatrix(Ln_Cnt, 4))
            PI_SrNo = Ln_Cnt
        Next
    End With
End Sub

Private Sub LoadGRNTrans()
Dim lb_found As Boolean
Dim Ln_Cnt   As Integer
Dim temp As String
    
InitializeGrid
    
    lb_found = MySeek(Gs_compcode & ls_TransType & txtTransNo, "AdjFind", PR_ICAdJt)
   
    If lb_found Then
        With GrdGRN
            Do While Gs_compcode & ls_TransType & txtTransNo = PR_ICAdJt("AdjFind")
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = PR_ICAdJt("SerialNo")
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = Trim(PR_ICAdJt("LocationCode1")) & Trim(PR_ICAdJt("ItemClass")) & Trim(PR_ICAdJt("ItemCode"))
                .TextMatrix(.Row, 2) = Abs(PR_ICAdJt("Quantity") & "")
                .TextMatrix(.Row, 3) = PR_ICAdJt("UnitCost")
                .TextMatrix(.Row, 4) = Abs(Round(PR_ICAdJt("Quantity") * PR_ICAdJt("UnitCost"), 0))
                .TextMatrix(.Row, 5) = Trim(PR_ICAdJt("LocationCode1"))
                .TextMatrix(.Row, 6) = PR_ICAdJt("ItemClass")
                .TextMatrix(.Row, 7) = PR_ICAdJt("ItemCode")
                .Rows = .Rows + 1
                PR_ICAdJt.MoveNext
                If PR_ICAdJt.EOF Then Exit Do
             Loop
            
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
        TotalGRN
    Else
        Call SetErr("Transaction Issue not found.", vbCritical)
        txtLocCode.SetFocus
    End If
End Sub

Private Sub grdGRN_DblClick()
    With GrdGRN
        If .Row > 0 Then
            PI_CurRow = .Row
        End If
        
        TxtItemCode = .TextMatrix(.Row, 1)
        txtqty = .TextMatrix(.Row, 2)
        txtunitprice = Val(.TextMatrix(.Row, 3))
        If MySeek(txtLocCode.Text & .TextMatrix(.Row, 1), "ItemFind", PR_IcItem) Then
            TxtFactor = PR_IcItem.Fields("ItemFactor")
            TxtUM = PR_IcItem.Fields("ItemUM")
        Else
            TxtFactor = ""
            TxtUM = ""
        End If
        PS_RowClicked = "Y"
    End With
End Sub

Private Sub txtvaluedate_KeyDown(KeyCode As Integer, Shift As Integer)
    If Lastkey(KeyCode) Then
       If txtvaluedate.Value < DateValue(Gs_Fnperiod) Or txtvaluedate.Value > DateValue(Gs_FnEndPeriod) Then
         Call SetErr("Invalid Period.", vbCritical)
         txtvaluedate.SetFocus
       Else
        txtLocCode.SetFocus
       End If
       
    End If
End Sub

Public Sub setfrmenv(ls_mode As String)
    txtLocCode.Enabled = IIf(ls_mode <> "D", True, False)
    txtAdjCode.Enabled = IIf(ls_mode <> "D", True, False)
    TxtRemarks.Enabled = IIf(ls_mode <> "D", True, False)
    Frame2.Enabled = IIf(ls_mode <> "D", True, False)
End Sub


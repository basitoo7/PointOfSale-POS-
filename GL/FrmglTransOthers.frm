VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmglTransOthers 
   Caption         =   "GL Transaction."
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmglTransOthers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   11295
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtActDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "SKIP"
      Top             =   6645
      Width           =   8460
   End
   Begin Crystal.CrystalReport rptVoucher 
      Left            =   -75
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowBorderStyle=   3
      WindowControlBox=   0   'False
      WindowMaxButton =   0   'False
      WindowMinButton =   0   'False
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.TextBox txtTDr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "###,###,###"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
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
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   9750
      MaxLength       =   11
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Debit Amount"
      Top             =   6630
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Height          =   2835
      Left            =   105
      TabIndex        =   1
      Top             =   570
      Width           =   11115
      Begin VB.ComboBox txtaccounttype 
         Height          =   330
         ItemData        =   "FrmglTransOthers.frx":030A
         Left            =   1335
         List            =   "FrmglTransOthers.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2010
         Width           =   1875
      End
      Begin VB.TextBox txtTCr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,###"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
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
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   9585
         MaxLength       =   11
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Debit Amount"
         Top             =   1605
         Width           =   1380
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   15
         MaxLength       =   64
         TabIndex        =   25
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1680
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   3240
         Picture         =   "FrmglTransOthers.frx":0327
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1650
         Width           =   315
      End
      Begin VB.TextBox txtaccountdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3585
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   21
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1650
         Width           =   4875
      End
      Begin VB.TextBox txtAccountNo 
         BackColor       =   &H00FFFF00&
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
         Left            =   1350
         MaxLength       =   13
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         ToolTipText     =   "Enter Voucher Type"
         Top             =   1650
         Width           =   1860
      End
      Begin VB.TextBox txtvchrType 
         BackColor       =   &H00FFFF00&
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
         Left            =   1350
         MaxLength       =   3
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         ToolTipText     =   "Enter Voucher Type"
         Top             =   195
         Width           =   600
      End
      Begin VB.TextBox txtvchrno 
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
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         ToolTipText     =   "Voucher No"
         Top             =   555
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker txtTransDate 
         Height          =   315
         Left            =   9405
         TabIndex        =   11
         Tag             =   "SKIPN"
         Top             =   915
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   64552961
         CurrentDate     =   37404
      End
      Begin VB.TextBox txtRemarks 
         Height          =   315
         Left            =   1335
         MaxLength       =   400
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Voucher Nirration"
         Top             =   2415
         Width           =   9630
      End
      Begin VB.TextBox Txtinstrument 
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
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Account No"
         Top             =   1290
         Width           =   3900
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   180
         MaxLength       =   64
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1680
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.TextBox txtVchrDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2325
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   195
         Width           =   8670
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   315
         Left            =   1980
         Picture         =   "FrmglTransOthers.frx":0499
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   195
         Width           =   315
      End
      Begin MSComCtl2.DTPicker txtvaluedate 
         Height          =   315
         Left            =   1350
         TabIndex        =   14
         Tag             =   "SKIPN"
         Top             =   915
         Width           =   1620
         _ExtentX        =   2858
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
         Format          =   64552961
         CurrentDate     =   37293
      End
      Begin VB.Label Label7 
         Caption         =   "Account Type :"
         Height          =   255
         Left            =   165
         TabIndex        =   29
         Top             =   2040
         Width           =   1125
      End
      Begin VB.Label LBLTopdesc 
         Caption         =   "Credit Total :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   8505
         TabIndex        =   27
         ToolTipText     =   "Enter Value Date"
         Top             =   1650
         Width           =   1065
      End
      Begin VB.Label Label12 
         Caption         =   "Account No :"
         Height          =   255
         Left            =   330
         TabIndex        =   23
         Top             =   1680
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Voucher No :"
         Height          =   255
         Left            =   330
         TabIndex        =   19
         Top             =   570
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Voucher Type :"
         Height          =   255
         Left            =   165
         TabIndex        =   18
         Top             =   210
         Width           =   1125
      End
      Begin VB.Label label2 
         Caption         =   "Value Date :"
         Height          =   255
         Left            =   420
         TabIndex        =   17
         ToolTipText     =   "Enter Value Date"
         Top             =   930
         Width           =   885
      End
      Begin VB.Label Label5 
         Caption         =   "Remarks :"
         Height          =   255
         Left            =   570
         TabIndex        =   16
         ToolTipText     =   "Enter Value Date"
         Top             =   2445
         Width           =   720
      End
      Begin VB.Label Label11 
         Caption         =   "Instrument No :"
         Height          =   255
         Left            =   195
         TabIndex        =   15
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Transaction Date :"
         Height          =   255
         Left            =   8040
         TabIndex        =   10
         ToolTipText     =   "Enter Value Date"
         Top             =   945
         Width           =   1335
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   1005
      ButtonWidth     =   1376
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
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
            Caption         =   "&Print"
            Description     =   "Print Listing."
            Object.ToolTipText     =   "Print listing."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Re&fresh"
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
               Picture         =   "FrmglTransOthers.frx":060B
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmglTransOthers.frx":0A5F
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmglTransOthers.frx":0EB3
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmglTransOthers.frx":1307
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmglTransOthers.frx":175B
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmglTransOthers.frx":1BAF
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmglTransOthers.frx":2303
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   90
      TabIndex        =   2
      Top             =   3330
      Width           =   11130
      Begin VB.TextBox TXTBARCODE 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         MaxLength       =   255
         TabIndex        =   31
         Tag             =   "SKIP"
         Top             =   0
         Visible         =   0   'False
         Width           =   2310
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdGRN 
         Height          =   3000
         Left            =   45
         TabIndex        =   24
         Top             =   165
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   5292
         _Version        =   393216
         BackColor       =   16777215
         RowHeightMin    =   300
         BackColorSel    =   16777215
         ForeColorSel    =   0
         GridColor       =   -2147483632
         FocusRect       =   2
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Label LBLBottomdesc 
      Caption         =   "Debit Total :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   8745
      TabIndex        =   4
      ToolTipText     =   "Enter Value Date"
      Top             =   6675
      Width           =   990
   End
   Begin VB.Menu File_Mainmenu 
      Caption         =   "File"
      Begin VB.Menu New_Record 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu Edit_Record 
         Caption         =   "Edit"
         Shortcut        =   ^E
      End
      Begin VB.Menu Delete_Record 
         Caption         =   "Delete"
         Shortcut        =   ^D
      End
      Begin VB.Menu Save_Record 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu Edit_Mainmenu 
      Caption         =   "Edit"
      Begin VB.Menu Copy_Data 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu Paste_data 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu Add_Row 
         Caption         =   "Add Row"
         Shortcut        =   ^I
      End
      Begin VB.Menu Delete_Row 
         Caption         =   "Delete Row"
         Shortcut        =   ^R
      End
      Begin VB.Menu Move_Back 
         Caption         =   "Move Back"
         Shortcut        =   ^Z
      End
   End
End
Attribute VB_Name = "FrmglTransOthers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PB_BlnkTrns As Boolean
Dim PI_CurRow    As Integer
Dim PI_SrNo      As Integer

Dim Mode As String
Dim PS_RowClicked As String
Dim ls_TFields As String
Dim ld_PrvDate As Date
Dim ln_FrmCount As Integer
Dim PR_Branch As New Recordset
Dim PR_VchCntr As New Recordset
Dim PR_Sub1 As New Recordset
Dim PR_GlTrans  As New Recordset
Dim PR_GlRef As New Recordset
Dim PR_VchType  As New Recordset
Dim PR_ActDetail As New Recordset
Dim pr_dumy As New Recordset

Dim PR_Crncy As New Recordset
Dim ln_OrgVchNo As Double
Dim ls_VchType As String
Dim ld_VluDate As Date
Dim ls_VchNo   As String
Dim ls_VchBase   As String
Dim ls_sql   As String

Dim ls_VchDesc As String
Dim ls_Narration As String
Dim ls_DAccount As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim ld_valueDate As Date
Dim ln_cnt As Double
Dim CX, CY
Dim ClickRow
Dim TboxCol

Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtvchrType
    Set PO_DESC = txtvchrno
    If Mode = "A" Then
        Gs_SQL = "Select VchrType 'Voucher Type',  VchrDescrip 'Description' from GlVchrType"
        Gs_FindFld = "VchrDescrip"
        Gs_Subon = True
        Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
        Gs_OrderBy = "Order by VchrDescrip"
        MyLookupOLDB.Caption = "Voucher Types"
        MyLookupOLDB.Show 1
       If Len(txtvchrType) > 0 Then txtVchrType_KeyDown vbKeyReturn, vbKeyShift

    Else
        Gs_SQL = "SELECT VchrType, Voucher_no, Vchr_Remarks,Amount FROM  Gl_Ref"
        Gs_OtherPara = " WHERE Compcode = '" & Gs_compcode & "' and ventry = 2"
        Gs_FindFld = "GRNCode"
        MyLookupOLDBsearchmultipul.Caption = "Vouchers"
        MyLookupOLDBsearchmultipul.Show 1
        If Len(txtvchrType) > 0 Then txtVchrType_KeyDown vbKeyReturn, vbKeyShift
        If Len(txtvchrno) > 0 Then txtvchrno_KeyDown vbKeyReturn, vbKeyShift
        End If
End Sub

Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtAccountNo
    Set PO_DESC = txtaccountdesc
    Gs_SQL = "Select  AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtAccountNo) > 0 Then Call TxtAccountNo_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Delete_record_Click()
Mode = DentMode(Mode, 3, PR_GlTrans, Me, txtvchrType, txtvchrType, "X", "X", 3, "X", "X", 1, False, Toolbar1)
InitializeGrid
End Sub

Private Sub Edit_record_Click()
Mode = DentMode(Mode, 2, PR_GlTrans, Me, txtvchrType, txtvchrType, "X", "X", 3, "X", "X", 1, False, Toolbar1)
InitializeGrid
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF11 Then
       Mode = DentMode(Mode, 4, PR_GlTrans, Me, txtvchrType, txtvaluedate, "X", "X", 3, "X", "X", 1, False, Toolbar1)
  End If
End Sub

Private Sub Form_Load()
    Dim ln_cnt As Integer
    Dim ls_PrvAlia As String
    Dim ls_PFields As String
    Dim SqlStr As New ADODB.Command
   Me.Caption = Me.Caption + " (" + Gs_CompName + ")"
' Setting up Preveliges
  SetToolBar(1) = chkRights("GLTRANS001")
  SetToolBar(2) = chkRights("GLTRANS002")
  SetToolBar(3) = chkRights("GLTRANS003")
  SetToolBar(4) = chkRights("GLTRANS004")

  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
    
    PR_Branch.Open "Select * from SysBranch Where Compcode = '" & Gs_compcode & "' order by Branchcode", gc_dbcon, adOpenStatic, adLockReadOnly, 1
    PR_VchType.Open "SELECT *,VchrType As FindFld FROM GlVchrType WHERE CompCode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1

    PR_Crncy.Open "Select * From SysCurrency Order By Crncy_Code", gc_dbcon, adOpenStatic, adLockReadOnly, 1

    InitializeGrid
    
    txtvaluedate = Date
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_Branch.Close
    PR_VchType.Close
    PR_Crncy.Close
End Sub
'Private Sub LoadGLSub2()
'If PR_Sub1.State = 1 Then PR_Sub1.Close
'PR_Sub1.Open "Select * from gl_sub2 where compcode = '" & Gs_compcode & "' and acct_sub1+acct_sub2 = '" & Left(txtAccountNo, 9) & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
'If Not PR_Sub1.EOF Then
'    txtsub = Trim(PR_Sub1("Acct_Desc") & "")
'End If
'PR_Sub1.Close
'End Sub


Private Sub Move_Back_Click()
    With GrdGRN
    .SetFocus
    .CellBackColor = vbWindowBackground
    If .Col = 5 Then
    .Col = 4
    TboxCol = True
    GrdGRN_EnterCell
    ElseIf .Col = 4 Then
    .Col = 3
    TboxCol = True
    ElseIf .Col = 3 Then
    .Col = 2
    ElseIf .Col = 2 Then
    .Col = 5
    If .Row >= 2 Then
    .Row = .Row - 1
    End If
    Else
    .Col = 2
    End If
    End With
End Sub

Private Sub New_Record_Click()
Mode = DentMode(Mode, 1, PR_GlTrans, Me, txtvchrType, txtvchrType, "X", "X", 3, "X", "X", 1, False, Toolbar1)
InitializeGrid
End Sub

Private Sub Save_Record_Click()
Mode = DentMode(Mode, 4, PR_GlTrans, Me, txtvchrType, txtvchrType, "X", "X", 3, "X", "X", 1, False, Toolbar1)
End Sub
Private Sub TxtAccountNo_KeyDown(KeyCode As Integer, Shift As Integer)
If txtAccountNo <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  AccountNo = '" & txtAccountNo & "' "
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                txtaccountdesc = pr_dumy("description")
                   If PR_VchType("PRStatus") = 1 Then
                        txtaccounttype.ListIndex = 1
                    Else
                        txtaccounttype.ListIndex = 0
                    End If
                txtaccounttype.SetFocus
            End If
         pr_dumy.Close

ElseIf txtAccountNo = "" And KeyCode = vbKeyReturn Then
    Command2_Click
End If

End Sub

Private Sub txtaccounttype_Click()
If txtaccounttype.ListIndex = 1 Then
    LBLTopdesc = "Credit Total :"
    LBLBottomdesc = "Debit Total :"
Else
    LBLTopdesc = "Debit Total :"
    LBLBottomdesc = "Credit Total :"
End If
End Sub

Private Sub txtaccounttype_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtRemarks.SetFocus
End Sub


Private Sub txtinstrument_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If Txtinstrument = "" Then
          Call SetErr("Cannot be empty", vbCritical)
          Txtinstrument.SetFocus
       Else
          txtAccountNo.SetFocus
       End If
    ElseIf KeyCode = vbKeyPageUp Then
        txtvaluedate.SetFocus
    End If
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtRemarks = Replace(txtRemarks, "'", "")
           GrdGRN.Col = 1
           GrdGRN.SetFocus
    End If
   
End Sub

Private Sub txtRemarks_LostFocus()
If txtRemarks <> "" Then
 txtRemarks = UCase(Replace(txtRemarks, "'", ""))
End If
End Sub

Private Sub txtvaluedate_KeyDown(KeyCode As Integer, Shift As Integer)
       If KeyCode = vbKeyReturn Then
        If (Format(DateValue(txtvaluedate.Value), "YYYY/MM/DD") < Format(DateValue(Gs_Fnperiod), "YYYY/MM/DD")) Or (Format(DateValue(txtvaluedate.Value), "YYYY/MM/DD") > Format(DateValue(Gs_FnEndPeriod), "YYYY/MM/DD")) Then
            Call SetErr("Invalid Period's Transaction.", vbCritical)
            txtvaluedate.SetFocus
            Exit Sub
        Else
           If PR_VchType.Fields("VchrBase") = "B" Then Txtinstrument.SetFocus
           If PR_VchType.Fields("VchrBase") <> "B" Then txtAccountNo.SetFocus
                      
        End If
    End If
End Sub

Private Sub txtvaluedate_LostFocus()
'ld_PrvDate = txtvaluedate.Value
End Sub

Private Sub txtvchrno_KeyDown(KeyCode As Integer, Shift As Integer)
    If Lastkey(KeyCode) Then
        txtvchrno = DoPad(txtvchrno.Text, 10)
        If Val(txtvchrno.Text) = 0 Then
           Call SetErr("Invalid Voucher No.", vbCritical)
           txtvchrType.SetFocus
        Else
          ls_VchNo = txtvchrno
          'If Mode = "E" Then DTPmdate.SetFocus
           LoadVchrRef
        End If
    End If
End Sub

Private Sub txtvchrType_Change()
If txtvchrType = "" Then
    txtVchrDesc = ""
End If
End Sub

Private Sub txtVchrType_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn And txtvchrType <> "" Then
         txtvchrType = UCase(txtvchrType)
      
         lb_found = MySeek(txtvchrType.Text, "FindFld", PR_VchType)
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   txtvchrno = ""
                   txtvchrType.SetFocus
                Else
                    txtVchrDesc = PR_VchType.Fields("VchrDescrip")
                    ls_VchBase = PR_VchType("VchrBase")
                    If PR_VchType("PRStatus") = 1 Then
                        txtaccounttype.ListIndex = 1
                    Else
                        txtaccounttype.ListIndex = 0
                    End If
                    ls_DAccount = Trim(PR_VchType("AccountNo") & "")
                       If Mode = "A" Then
                         If PR_VchType.Fields("VchrFrequency") = "1" Then
                            pr_dumy.Open "select max(voucher_no) as voucherno from gl_ref where compcode = '" & Gs_compcode & "' and branchcode = '" & Gs_BranchCode & "' and vchrtype = '" & txtvchrType & "' and month(value_date) = " & Month(txtvaluedate) & " and year(value_date) = " & Year(txtvaluedate) & "", gc_dbcon, adOpenStatic, adLockReadOnly, 1
                            If Not pr_dumy.EOF Then
                                txtvchrno = DoPad(Trim(str(Val(0 & pr_dumy("voucherno")) + 1)), 10)
                            Else
                                txtvchrno = DoPad(Trim(str(Val(1))), 10)
                            End If
                            pr_dumy.Close
                         Else
                            pr_dumy.Open "select max(voucher_no) as voucherno from gl_ref where compcode = '" & Gs_compcode & "' and branchcode = '" & Gs_BranchCode & "' and vchrtype = '" & txtvchrType & "' and value_date >= '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' and value_date <= '" & Format(Gs_FnEndPeriod, "YYYY/MM/DD") & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
                            If Not pr_dumy.EOF Then
                                txtvchrno = DoPad(Trim(str(Val(0 & pr_dumy("voucherno")) + 1)), 10)
                            Else
                                txtvchrno = DoPad(Trim(str(Val(1))), 10)
                            End If
                            pr_dumy.Close
                            
                         End If
                            ln_OrgVchNo = Val(txtvchrno)
                           txtvaluedate.SetFocus
                         
                      Else
                         txtvchrno.SetFocus
                    End If
                End If
  ElseIf KeyCode = vbKeyReturn And Trim(txtvchrType) = "" Then
    Call cmdLookup_Click
  ElseIf KeyCode = vbKeyF12 Then
   Call cmdLookup_Click
  End If
  End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Index = 5 And Mode <> "D" Then Call setprint
Mode = DentMode(Mode, Button.Index, PR_GlTrans, Me, txtvchrType, txtvchrType, "X", "X", 3, "X", "X", 1, False, Toolbar1)
If Button.Index <> 4 Then InitializeGrid
If Button.Index = 5 And txtvchrType <> "" And txtvchrno <> "" Then Call setprint
    
End Sub
Private Sub setprint()
On Error GoTo LocalErr
Dim ls_BranchName As String
 


If txtvchrno <> "" Then
         If MySeek(Gs_BranchCode, "BranchCode", PR_Branch) Then ls_BranchName = PR_Branch("BranchDesc")
   With rptVoucher
        
        .ReportFileName = App.Path & Gs_GlRepoPath & "\Vchr_Print.RPT"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = '" & txtVchrDesc & "'"
        .Formulas(5) = "BranchName = '" & Gs_BranchCode + "-" + ls_BranchName & "'"
        .SelectionFormula = "{Gl_Trans.Voucher_No} = '" & Trim(txtvchrno) & "' and {Gl_Trans.BranchCode} = '" & Gs_BranchCode & "'"
        .SelectionFormula = .SelectionFormula & " and {Gl_Trans.VchrType} = '" & Trim(txtvchrType) & "'"
        .SelectionFormula = .SelectionFormula & " and {Gl_Trans.CompCode} = '" & Gs_compcode & "'"
        '.SelectionFormula = .SelectionFormula & " and {Gl_Trans.Value_Date} = Date(" & Year(txtvaluedate) & "," & Month(txtvaluedate) & "," & Day(txtvaluedate) & ")"
        .Formulas(2) = "Sig1 = '" & Gc_UserName & "'"
        .Formulas(3) = "Sig2 = '" & Gs_Sign2 & "'"
        .Formulas(4) = "Sig3 = '" & Gs_Sign3 & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
   End With
 End If
Exit Sub
LocalErr:
Call SetErr("Printer Not Ready", vbCritical)
End Sub
Public Sub InitializeGrid()
    With GrdGRN
        .Redraw = False
        .Clear
        .Rows = 2
        .Cols = 2
        .FormatString = "Sr# |<Account No|<Account Title|<Instrument #|<Account Narration|<Amount"
        .ColWidth(1) = 1500
        .ColWidth(2) = 2000
        .ColWidth(3) = 1500
        .ColWidth(4) = 3300
        .ColWidth(5) = 1500
        .ColAlignment(5) = 7
        '.ColWidth(4) = 0
        
        .Redraw = True
    End With
     ClickRow = ""
End Sub


Private Sub TotalAmount()
    Dim ln_cnt As Integer
    txtTDr = 0
    txtTCr = 0
    With GrdGRN
        For ln_cnt = 1 To .Rows - 1
            .TextMatrix(ln_cnt, 0) = ln_cnt
            txtTDr = txtTDr + Val(.TextMatrix(ln_cnt, 5))
            txtTCr = Val(txtTDr)
        Next
    End With
End Sub
Public Sub SaveValues()
Dim lb_Vstat As Boolean
Dim ln_cnt As Integer
Dim ln_Cols As Integer
Dim ls_TotalSubs As String
Dim ln_TmpVchNo As Double
Dim ls_opt As String
Dim ls_sql As String

On Error GoTo RollBack

ln_OrgVchNo = txtvchrno

gc_dbcon.BeginTrans
     Select Case Mode
           Case "D"
              ' Delete Detail of Voucher
              ls_sql = "DELETE FROM Gl_Trans WHERE CompCode = '" & Gs_compcode & "' And BranchCode = '" & Gs_BranchCode & "' AND Voucher_No = '" & txtvchrno & "' AND VchrType = '" & txtvchrType & "' "
              gc_dbcon.Execute ls_sql
              
              ls_sql = "DELETE FROM Gl_Ref WHERE CompCode = '" & Gs_compcode & "' And BranchCode = '" & Gs_BranchCode & "' AND Voucher_No = '" & txtvchrno & "' AND VchrType = '" & txtvchrType & "' "
              gc_dbcon.Execute ls_sql
                
           Case Else
                If Mode = "E" Then
                 ' Delete Reference of Voucher
                   ls_sql = "DELETE FROM Gl_Ref WHERE CompCode = '" & Gs_compcode & "' And BranchCode = '" & Gs_BranchCode & "' AND Voucher_No = '" & txtvchrno & "' AND VchrType = '" & txtvchrType & "'"
                   gc_dbcon.Execute ls_sql
                   
                   ls_sql = "DELETE FROM Gl_Trans WHERE CompCode = '" & Gs_compcode & "' And BranchCode = '" & Gs_BranchCode & "' AND Voucher_No = '" & txtvchrno & "' AND VchrType = '" & txtvchrType & "'"
                   gc_dbcon.Execute ls_sql
                Else
                  
                  If Mode = "A" Then
                         PR_VchType.Filter = "VchrType = '" & txtvchrType & "'"
                        If PR_VchType.Fields("VchrFrequency") = "1" Then
                            pr_dumy.Open "select max(voucher_no) as voucherno from gl_ref where compcode = '" & Gs_compcode & "' and branchcode = '" & Gs_BranchCode & "' and vchrtype = '" & txtvchrType & "' and month(value_date) = " & Month(txtvaluedate) & " and year(value_date) = " & Year(txtvaluedate) & "", gc_dbcon, adOpenStatic, adLockReadOnly, 1
                            If Not pr_dumy.EOF Then
                                txtvchrno = DoPad(Trim(str(Val(0 & pr_dumy("voucherno")) + 1)), 10)
                            Else
                                txtvchrno = DoPad(Trim(str(Val(1))), 10)
                            End If
                            pr_dumy.Close
                        Else
                            pr_dumy.Open "select max(voucher_no) as voucherno from gl_ref where compcode = '" & Gs_compcode & "' and branchcode = '" & Gs_BranchCode & "' and vchrtype = '" & txtvchrType & "' and value_date >= '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' and value_date <= '" & Format(Gs_FnEndPeriod, "YYYY/MM/DD") & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
                            If Not pr_dumy.EOF Then
                                txtvchrno = DoPad(Trim(str(Val(0 & pr_dumy("voucherno")) + 1)), 10)
                            Else
                                txtvchrno = DoPad(Trim(str(Val(1))), 10)
                            End If
                            pr_dumy.Close
                            
                        End If
                         
                        If Val(txtvchrno) <> Val(ln_OrgVchNo) Then
                            lb_Vstat = True
                        End If
                       PR_VchType.Filter = adFilterNone
                     End If
                 
                End If
                ' Save References of Voucher
                
                 ls_sql = "INSERT into Gl_Ref(compcode,BranchCode,Value_Date,Trans_Date, Voucher_No, VchrType, Vchr_Remarks,InstrumentNo,CrncyCode,ExchgRate,userid,adddate,addtime,Accountno,ActType,Amount,ventry) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & Format(txtTransDate, "YYYY/MM/DD") & "','" & txtvchrno & "','" & txtvchrType & "','" & txtRemarks & "','" & Txtinstrument & "','PKR'," & Val(1) & ",'" & Gc_UserId & "','" & Format(Gd_SysDate, "YYYY/MM/DD") & "','" & Time & "','" & txtAccountNo & "'," & txtaccounttype.ListIndex & " ," & Val(txtTCr) & "  ,2 )"
                 gc_dbcon.Execute ls_sql
                 
                ' Save Details of Voucher
                    With GrdGRN
                       For ln_cnt = 1 To .Rows - 1
                          If Len(Trim(.TextMatrix(ln_cnt, 1))) > 0 Then
                           If txtaccounttype.ListIndex = 0 Then
                            ls_sql = "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, ExchgRate,userid,adddate,addtime,Acct_Nirration,AcctName,instrumentno) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & .TextMatrix(ln_cnt, 1) & "'," & .TextMatrix(ln_cnt, 0) & ",'" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & txtvchrno & "','" & txtvchrType & "',0," & Val(.TextMatrix(ln_cnt, 5)) & "," & Val(1) & ",'" & Gc_UserId & "','" & Format(Gd_SysDate, "YYYY/MM/DD") & "','" & Time & "','" & Trim(.TextMatrix(ln_cnt, 4)) & "','" & Trim(.TextMatrix(ln_cnt, 2)) & "','" & Trim(.TextMatrix(ln_cnt, 3)) & "')"
                           Else
                            ls_sql = "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, ExchgRate,userid,adddate,addtime,Acct_Nirration,AcctName,instrumentno) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & .TextMatrix(ln_cnt, 1) & "'," & .TextMatrix(ln_cnt, 0) & ",'" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & txtvchrno & "','" & txtvchrType & "'," & Val(.TextMatrix(ln_cnt, 5)) & ",0," & Val(1) & ",'" & Gc_UserId & "','" & Format(Gd_SysDate, "YYYY/MM/DD") & "','" & Time & "','" & Trim(.TextMatrix(ln_cnt, 4)) & "','" & Trim(.TextMatrix(ln_cnt, 2)) & "','" & Trim(.TextMatrix(ln_cnt, 3)) & "')"
                           End If
                           PI_SrNo = ln_cnt
                           gc_dbcon.Execute ls_sql
                          End If
                       Next
                    End With
                           PI_SrNo = PI_SrNo + 1
                           If txtaccounttype.ListIndex = 0 Then
                            ls_sql = "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, ExchgRate,userid,adddate,addtime,Acct_Nirration,AcctName,instrumentno) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtAccountNo & "'," & PI_SrNo & ",'" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & txtvchrno & "','" & txtvchrType & "'," & Val(txtTCr) & ",0," & Val(1) & ",'" & Gc_UserId & "','" & Format(Gd_SysDate, "YYYY/MM/DD") & "','" & Time & "','" & Trim(txtRemarks) & "','" & Trim(txtaccountdesc) & "','" & Txtinstrument & "')"
                           Else
                            ls_sql = "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, ExchgRate,userid,adddate,addtime,Acct_Nirration,AcctName,instrumentno) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtAccountNo & "'," & PI_SrNo & ",'" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & txtvchrno & "','" & txtvchrType & "',0," & Val(txtTCr) & "," & Val(1) & ",'" & Gc_UserId & "','" & Format(Gd_SysDate, "YYYY/MM/DD") & "','" & Time & "','" & Trim(txtRemarks) & "','" & Trim(txtaccountdesc) & "','" & Txtinstrument & "')"
                           End If
                           gc_dbcon.Execute ls_sql
     
     
     
     
     End Select
Call SetClear(Me)
gc_dbcon.CommitTrans

If lb_Vstat Then Call SetErr("Your New Transaction Voucher No will be " + txtvchrno.Text, vbCritical)
If Mode <> "D" Then
   ls_opt = SetErr("Print Voucher ?.", vbYesNo)
   If ls_opt = vbYes Then Call setprint
End If
Call txtVchrType_KeyDown(vbKeyReturn, vbKeyShift)
txtvchrType.SetFocus


InitializeGrid

Exit Sub

RollBack:
MsgBox Err.Description
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
gc_dbcon.RollbackTrans
ln_FrmCount = 0
On Error GoTo 0
End Sub

Public Function ChkInputs() As Boolean

If Trim(txtvchrType) = "" Then
    Call MsgBox("Enter/Select Voucher Type  !!!", vbCritical)
    ChkInputs = False
ElseIf Trim(txtvchrno) = "" Then
    Call MsgBox("Must enter Voucher No !!!", vbCritical)
    ChkInputs = False
ElseIf Trim(Txtinstrument) = "" And ls_VchBase = "B" Then
    Call MsgBox("Bank Payment/Receive must enter instrument no !!!", vbCritical)
    ChkInputs = False
ElseIf Trim(txtAccountNo) = "" Then
    Call MsgBox("Enter/Select Account no!!!", vbCritical)
    ChkInputs = False
ElseIf Trim(txtRemarks) = "" Then
    Call MsgBox("Enter Voucher Remarks!!!", vbCritical)
    ChkInputs = False
ElseIf Val(txtTDr) = 0 And Val(txtTCr) = 0 Then
    Call MsgBox("Debit and Credit Amount is equal to Zero !!!", vbCritical)
    ChkInputs = False
ElseIf Val(txtTDr) <> Val(txtTCr) Then
    Call MsgBox("Debit and Credit Amount not balance !!!", vbCritical)
    ChkInputs = False
ElseIf GrdGRN.Rows < 2 Then
    Call MsgBox("Enter Grid Entries !!!", vbCritical)
    ChkInputs = False
Else
            ChkInputs = True
End If
End Function

Private Sub LoadVchrRef()
Dim lb_found As Boolean
    If PR_GlRef.State = 1 Then PR_GlRef.Close
    PR_GlRef.Open "Select * from gl_ref where compcode = '" & Gs_compcode & "' and BranchCode+VchrType + Voucher_no+ ltrim(str(month(value_date)))+ ltrim(str(Year(value_date))) = '" & Trim(Gs_BranchCode) & Trim(txtvchrType) & Trim(txtvchrno) + LTrim(str(Month(txtvaluedate.Value))) + LTrim(str(Year(txtvaluedate.Value))) & "' and ventry = 2", gc_dbcon, adOpenStatic, adLockReadOnly, 1
    lb_found = PR_GlRef.EOF
      Select Case Mode
         Case "A"
            If lb_found Then
               Call MsgBox("Voucher already exist !!!", vbCritical)
               txtvchrType.SetFocus
            Else
               If PR_VchType.Fields("VchrBase") = "B" Then Txtinstrument.SetFocus
               If PR_VchType.Fields("VchrBase") <> "B" Then txtRemarks.SetFocus
            End If
         Case Else
              If lb_found Then
                 Call SetErr(Gs_RecNFMsg, vbCritical)
                 InitializeGrid
                 txtvchrno.SetFocus
              Else
                  txtvaluedate = PR_GlRef("Value_Date")
                  txtTransDate = PR_GlRef("Trans_Date")
                  Txtinstrument = Trim(PR_GlRef("InstrumentNo") & "")
                  txtRemarks = Trim(PR_GlRef("Vchr_Remarks") & "")
                  txtAccountNo = Trim(PR_GlRef("Accountno") & "")
                  If txtAccountNo <> "" Then Call TxtAccountNo_KeyDown(vbKeyReturn, vbKeyShift)
                  txtaccounttype.ListIndex = Val(PR_GlRef("ActType"))
                  txtTCr = Val(PR_GlRef("Amount"))
                  LoadVchrTrans
                  
              End If
      End Select
End Sub

Private Sub LoadVchrTrans()
InitializeGrid

If txtaccounttype.ListIndex = 0 Then
ls_sql = "Select * from GL_Trans where compcode = '" & Gs_compcode & "' and BranchCode+VchrType + Voucher_no+ ltrim(str(month(value_date)))+ ltrim(str(Year(value_date))) = '" & Trim(Gs_BranchCode) & Trim(txtvchrType) & Trim(txtvchrno) + LTrim(str(Month(txtvaluedate.Value))) + LTrim(str(Year(txtvaluedate.Value))) & "' and Cr_amount >0 order by Serialno"
Else
ls_sql = "Select * from GL_Trans where compcode = '" & Gs_compcode & "' and BranchCode+VchrType + Voucher_no+ ltrim(str(month(value_date)))+ ltrim(str(Year(value_date))) = '" & Trim(Gs_BranchCode) & Trim(txtvchrType) & Trim(txtvchrno) + LTrim(str(Month(txtvaluedate.Value))) + LTrim(str(Year(txtvaluedate.Value))) & "' and Dr_amount >0 order by Serialno"

End If

PR_GlTrans.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
    If Not PR_GlTrans.EOF Then
        With GrdGRN
            Do While Not PR_GlTrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = PR_GlTrans("SerialNo")
                .TextMatrix(.Row, 1) = PR_GlTrans("AccountNo")
                .TextMatrix(.Row, 3) = PR_GlTrans("Instrumentno") & ""
                .TextMatrix(.Row, 4) = PR_GlTrans("Acct_Nirration") & ""
                If txtaccounttype.ListIndex = 0 Then
                .TextMatrix(.Row, 5) = PR_GlTrans("Cr_Amount")
                Else
                .TextMatrix(.Row, 5) = PR_GlTrans("Dr_Amount")
                End If
                .TextMatrix(.Row, 2) = PR_GlTrans("AcctName") & ""
                .Rows = .Rows + 1
                PR_GlTrans.MoveNext
                If PR_GlTrans.EOF Or PR_GlTrans.BOF Then Exit Do
             Loop
            
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
        TotalAmount
        If txtRemarks.Enabled Then txtRemarks.SetFocus
    Else
        Call SetErr("Voucher Transaction not found.", vbCritical)
        txtvchrno.SetFocus
    End If
        PR_GlTrans.Close
        PR_GlRef.Close
End Sub

Public Sub FrmRefresh()
    PR_Branch.Requery
    PR_VchType.Requery
End Sub
Public Sub GetKeysAdd(argFlexGrid As MSHFlexGrid, KeyAscii As Integer)
'This Procedure is used to display the pressed key into FlexGrid in Addition Mode
'so that when you press Enter Key in the last row then one row will be added.
'When you press the BackSpace Key in an empty Row then a Row will be Removed.
'On Error GoTo ErrHandler

If KeyAscii = 13 Then 'if Enter Key then...
  
  With argFlexGrid
        ' .SelectionMode = flexSelectionByRow
        .Row = .RowSel
      If .Col = 1 Then
        .CellBackColor = vbWindowBackground
       If .TextMatrix(.Row, 1) <> "" Then
          If pr_dumy.State = 1 Then pr_dumy.Close
          
          pr_dumy.Open "Select AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  AccountNo = '" & Trim(.TextMatrix(.Row, 1)) & " ' ", gc_dbcon, adOpenStatic, adLockReadOnly
          
          If pr_dumy.RecordCount <= 0 Then
              Call MsgBox("Account No not found !!!", vbCritical)
             .TextMatrix(.Row, 1) = ""
             
          Else
             .TextMatrix(.Row, 0) = .Row
             .TextMatrix(.Row, 2) = Trim(pr_dumy("Description") & "")
             .Col = .Col + 2
             .CellBackColor = vbHighlight
             txtActDesc = .TextMatrix(.Row, 2)
             If .Rows > 2 Then
              
             
              .TextMatrix(.Row, 3) = .TextMatrix(.Row - 1, 3)
              Else
              .TextMatrix(.Row, 3) = Txtinstrument
              End If
      
             
         pr_dumy.Close
         End If
       Else
           Call GrdGRN_KeyDown(112, vbKeyShift)
       End If
      ElseIf .Col = 3 Then
      .CellBackColor = vbWindowBackground
      .Col = .Col + 1
      If .Rows > 2 Then
              
             
              .TextMatrix(.Row, 4) = .TextMatrix(.Row - 1, 4)
              Else
              .TextMatrix(.Row, 4) = txtRemarks
              End If
            
             TboxCol = 5
             TXTBARCODE.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
             TXTBARCODE.Text = .TextMatrix(.Row, 4)
             TXTBARCODE.Visible = True
             ClickRow = .Row
             TXTBARCODE.SetFocus
       ElseIf .Col = 5 Then
         .CellBackColor = vbWindowBackground
        If .TextMatrix(.Row, 1) <> "" Then
          If .Row = .Rows - 1 Then
           .Rows = .Rows + 1
          End If
          .Col = 1
          .LeftCol = 1
          .Row = .Row + 1
          .SetFocus
        Else
         Call MsgBox("Enter/Select Accountno !!!", vbCritical)
         .Row = .Row
         .Col = 1
        End If
          
        If .RowSel > 10 Then
              .TopRow = .Rows - 1 'To Move the Scrollbar
        End If
            
   End If
   End With
 Exit Sub
End If
      
If KeyAscii = 8 Then  'If BackSpace Key then...
With argFlexGrid
   If .Col = 1 Or .Col = 3 Or .Col = 5 Then
   If Len(Trim(.Text)) <> 0 Then  'If current cell is not empty then...
      .Text = Left(.Text, (Len(.Text) - 1)) 'Removing a character from the right side of the FlexGrid cell's text
   End If
   End If
End With
End If

If KeyAscii <> 27 And KeyAscii <> 8 Then
    With GrdGRN
      
      If .Col = 1 Or .Col = 3 Then
        If .CellBackColor = vbHighlight Then
         .Text = "": .CellBackColor = vbWindowBackground
        End If
        .Text = .Text & Chr(KeyAscii) 'Reset Value in Cell and Append the pressed character to the right.
        
      ElseIf .Col = 4 Then
        TXTBARCODE.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
       .TextMatrix(.Row, 4) = Chr(KeyAscii)
        TXTBARCODE.Text = .TextMatrix(.Row, 4)
        TXTBARCODE.Visible = True
        TXTBARCODE.SelStart = Len(TXTBARCODE)
        ClickRow = .Row
        TXTBARCODE.SetFocus
        '.CellBackColor = vbWindowBackground
        '.Text = .Text & Chr(KeyAscii) 'Reset Value in Cell and Append the pressed character to the right.
      ElseIf .Col = 5 Then
        If .CellBackColor = vbHighlight Then
                .Text = "": .CellBackColor = vbWindowBackground
        End If
         .Text = .Text & Chr(KeyAscii)
          If Not IsNumeric(.Text) Then
          .Text = ""
           Call MsgBox("Enter Numeric entry !!!", vbCritical)
           Exit Sub
          End If
      
      End If
        TotalAmount
      
    End With
  End If
End Sub
Private Sub GrdGRN_Click()
GrdGRN.CellBackColor = vbHighlight
End Sub
Private Sub GrdGRN_EnterCell()
GrdGRN.CellBackColor = vbHighlight
If GrdGRN.Col = 4 Then
With GrdGRN

        TXTBARCODE.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
        TXTBARCODE.Text = .TextMatrix(.Row, 4)
        TXTBARCODE.Visible = True
        ClickRow = .Row
        TXTBARCODE.SetFocus
End With
End If

End Sub
Private Sub GrdGRN_LeaveCell()
With GrdGRN
 .CellBackColor = vbWindowBackground
End With
End Sub
Private Sub GrdGRN_KeyPress(KeyAscii As Integer)
On Error GoTo LocalErr
 Call GetKeysAdd(GrdGRN, KeyAscii)
Exit Sub

LocalErr:
Call MsgBox(Err.Description, vbCritical)
End Sub



Private Sub GrdGRN_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And GrdGRN.Col = 1 Then  ' F1 key pressed
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = Text1
    Set PO_DESC = Text2
    Gs_SQL = "Select  AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    GrdGRN.TextMatrix(GrdGRN.Row, 1) = Text1
    If GrdGRN.TextMatrix(GrdGRN.Row, 1) <> "" Then
        Call GrdGRN_KeyPress(13)
    End If
 ElseIf KeyCode = vbKeyDelete Then 'Delete Key Pressed
    With GrdGRN
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            .RemoveItem .Row
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                InitializeGrid
            End If
            'ResetRowSRNO
            TotalAmount
    End With
 ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then 'key down and keyup
    With GrdGRN
    End With
 End If
End Sub
Private Sub Copy_data_Click()
With GrdGRN
Clipboard.Clear
Clipboard.SetText .TextMatrix(.Row, .Col)
End With
End Sub

Private Sub Delete_row_Click()
   With GrdGRN
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            .RemoveItem .Row
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                InitializeGrid
            End If
            ResetRowSRNO
            TotalAmount
    End With
End Sub
Private Sub ResetRowSRNO()
With GrdGRN
   For ln_cnt = 1 To .Rows - 1
    .TextMatrix(ln_cnt, 0) = ln_cnt
   Next
End With
End Sub
Private Sub TXTBARCODE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyTab Then
TXTBARCODE_LostFocus
ElseIf KeyCode = vbKeyRight Then
TboxCol = True
TXTBARCODE_LostFocus

End If
End Sub

Private Sub TXTBARCODE_LostFocus()
With GrdGRN
If ClickRow <> "" Then
.TextMatrix(ClickRow, 4) = TXTBARCODE.Text
 TXTBARCODE.Text = ""
 ClickRow = ""
 End If
 .CellBackColor = vbWindowBackground
  TXTBARCODE.Visible = False
  If TboxCol = True Then
  .Col = 3
  .SetFocus

  TboxCol = False
  Else
 .SetFocus
 .Col = 5
 End If
 
 
End With


End Sub

Private Sub Paste_data_Click()
With GrdGRN
.TextMatrix(.Row, .Col) = Clipboard.GetText
End With
End Sub
Private Sub Add_Row_Click()
With GrdGRN
If .TextMatrix(.Row, 1) <> "" Then
          .CellBackColor = vbWindowBackground
          If .Row = .Rows - 1 Then
           .Rows = .Rows + 1
          End If
          .Col = 1
          .LeftCol = 1
          .Row = .Row + 1
          .Row = .Rows - 1
          .SetFocus
        Else
         Call MsgBox("Enter/Select Item Code!!!", vbCritical)
         .Row = .Row
         .Col = 1
        End If
          
        If .RowSel > 10 Then
              .TopRow = .Rows - 1 'To Move the Scrollbar
        End If
End With
End Sub





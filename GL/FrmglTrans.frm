VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmglTrans 
   Caption         =   "GL Transaction."
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   12270
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmglTrans.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   12270
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtActDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   150
      Locked          =   -1  'True
      MaxLength       =   64
      TabIndex        =   39
      TabStop         =   0   'False
      Tag             =   "SKIP"
      Top             =   7200
      Width           =   8610
   End
   Begin Crystal.CrystalReport rptVoucher 
      Left            =   2580
      Top             =   480
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
      Left            =   9375
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Debit Amount"
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox txtTCr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0;(""$""#,##0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
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
      Left            =   10785
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Debit Amount"
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   2445
      Left            =   90
      TabIndex        =   1
      Top             =   570
      Width           =   12135
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   45
         MaxLength       =   64
         TabIndex        =   38
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1725
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   2460
         Picture         =   "FrmglTrans.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1245
         Width           =   315
      End
      Begin VB.TextBox txtbranchname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2295
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   32
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   180
         Width           =   5400
      End
      Begin VB.CheckBox chkapproved 
         Alignment       =   1  'Right Justify
         Caption         =   "Approved :"
         Height          =   330
         Left            =   9510
         TabIndex        =   29
         Tag             =   "SKIPN"
         Top             =   180
         Visible         =   0   'False
         Width           =   1110
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
         TabIndex        =   22
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         ToolTipText     =   "Enter Voucher Type"
         Top             =   885
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
         TabIndex        =   21
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         ToolTipText     =   "Voucher No"
         Top             =   1245
         Width           =   1095
      End
      Begin VB.CommandButton CmdCurrency 
         Height          =   315
         Left            =   11085
         Picture         =   "FrmglTrans.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   570
         Width           =   315
      End
      Begin MSComCtl2.DTPicker txtTransDate 
         Height          =   315
         Left            =   6075
         TabIndex        =   15
         Tag             =   "SKIPN"
         Top             =   540
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   63766529
         CurrentDate     =   37404
      End
      Begin VB.TextBox txtRemarks 
         Height          =   375
         Left            =   1350
         MaxLength       =   400
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Voucher Nirration"
         Top             =   1965
         Width           =   10710
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
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Account No"
         Top             =   1605
         Width           =   3900
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   180
         MaxLength       =   64
         TabIndex        =   10
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
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   885
         Width           =   5355
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   315
         Left            =   1980
         Picture         =   "FrmglTrans.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   885
         Width           =   315
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   1965
         Picture         =   "FrmglTrans.frx":0760
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   180
         Width           =   315
      End
      Begin MSMask.MaskEdBox txtbranchcode 
         Height          =   315
         Left            =   1350
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         ToolTipText     =   "Default Currency"
         Top             =   180
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCrncyCode 
         Height          =   315
         Left            =   10455
         TabIndex        =   17
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         ToolTipText     =   "Default Currency"
         Top             =   570
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtexrate 
         Height          =   315
         Left            =   10455
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         ToolTipText     =   "Default Currency"
         Top             =   915
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   8
         Format          =   "##0.0000;(##0.0000)"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker txtvaluedate 
         Height          =   315
         Left            =   1350
         TabIndex        =   23
         Tag             =   "SKIPN"
         Top             =   540
         Width           =   1650
         _ExtentX        =   2910
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
         Format          =   63766529
         CurrentDate     =   37293
      End
      Begin MSComCtl2.DTPicker DTPmdate 
         Height          =   315
         Left            =   6060
         TabIndex        =   30
         Tag             =   "SKIPN"
         Top             =   1230
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         Format          =   63766529
         CurrentDate     =   37404
      End
      Begin MSComCtl2.DTPicker DTPChkDate 
         Height          =   315
         Left            =   10455
         TabIndex        =   33
         Tag             =   "SKIPN"
         Top             =   1260
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         Format          =   63766529
         CurrentDate     =   37404
      End
      Begin VB.Label Label15 
         Caption         =   "Chk Date :"
         Height          =   255
         Left            =   9690
         TabIndex        =   34
         ToolTipText     =   "Enter Value Date"
         Top             =   1305
         Width           =   1065
      End
      Begin VB.Label Label14 
         Caption         =   "Modify Date :"
         Height          =   255
         Left            =   5055
         TabIndex        =   31
         ToolTipText     =   "Enter Value Date"
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Voucher No :"
         Height          =   255
         Left            =   330
         TabIndex        =   28
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Voucher Type :"
         Height          =   255
         Left            =   180
         TabIndex        =   27
         Top             =   900
         Width           =   1125
      End
      Begin VB.Label label2 
         Caption         =   "Value Date :"
         Height          =   255
         Left            =   420
         TabIndex        =   26
         ToolTipText     =   "Enter Value Date"
         Top             =   555
         Width           =   885
      End
      Begin VB.Label Label5 
         Caption         =   "Remarks :"
         Height          =   255
         Left            =   570
         TabIndex        =   25
         ToolTipText     =   "Enter Value Date"
         Top             =   1965
         Width           =   720
      End
      Begin VB.Label Label11 
         Caption         =   "Instrument No :"
         Height          =   255
         Left            =   210
         TabIndex        =   24
         Top             =   1635
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Ex. Rate :"
         Height          =   255
         Left            =   9705
         TabIndex        =   19
         Top             =   960
         Width           =   705
      End
      Begin VB.Label Label12 
         Caption         =   "Currency :"
         Height          =   255
         Left            =   9585
         TabIndex        =   18
         Top             =   585
         Width           =   765
      End
      Begin VB.Label Label18 
         Caption         =   "Branch # :"
         Height          =   255
         Left            =   540
         TabIndex        =   14
         Top             =   195
         Width           =   765
      End
      Begin VB.Label Label4 
         Caption         =   "Transaction Date :"
         Height          =   255
         Left            =   4695
         TabIndex        =   13
         ToolTipText     =   "Enter Value Date"
         Top             =   570
         Width           =   1335
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12270
      _ExtentX        =   21643
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
               Picture         =   "FrmglTrans.frx":08D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmglTrans.frx":0D26
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmglTrans.frx":117A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmglTrans.frx":15CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmglTrans.frx":1A22
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmglTrans.frx":1E76
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmglTrans.frx":25CA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4170
      Left            =   90
      TabIndex        =   2
      Top             =   2940
      Width           =   12135
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
         Left            =   1545
         MaxLength       =   255
         TabIndex        =   37
         Tag             =   "SKIP"
         Top             =   1080
         Visible         =   0   'False
         Width           =   2310
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdGRN 
         Height          =   3885
         Left            =   60
         TabIndex        =   36
         Top             =   210
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   6853
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
   Begin VB.Label Label10 
      Caption         =   "Total :"
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
      Left            =   8865
      TabIndex        =   5
      ToolTipText     =   "Enter Value Date"
      Top             =   7245
      Width           =   525
   End
   Begin VB.Menu File_menu 
      Caption         =   "File"
      Begin VB.Menu New_menu 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu Edit_menu_sub 
         Caption         =   "Edit"
         Shortcut        =   ^E
      End
      Begin VB.Menu Delete_menu 
         Caption         =   "Delete"
         Shortcut        =   ^D
      End
      Begin VB.Menu Save_menu 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "FrmglTrans"
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
Dim ls_VchDesc As String
Dim ls_Narration As String
Dim ls_DAccount As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim ld_valueDate As Date
Dim res
Dim CX, CY
Dim ClickRow
Dim TboxCol



Private Sub CmdCurrency_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCrncyCode
    Set PO_DESC = Text1
    GoTop PR_Crncy
    MyLookup.Caption = "Currency Types"
    MyLookup.FillGrid PR_Crncy, "Crncy_code", "Crncy_Descrip", 5
    MyLookup.Show 1
    If txtCrncyCode.Text <> "" Then txtcrncycode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtvchrno
    Set PO_DESC = Text1
    Gs_SQL = "Select Voucher_no 'Voucher_no',  Vchr_remarks 'Nirration',Value_Date 'Value Date' from gl_ref"
    Gs_FindFld = "Voucher_no"
    Gs_Subon = True
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' and vchrtype = '" & txtvchrType & "'  and Value_date = '" & Format(txtvaluedate, "YYYY/MM/DD") & "' "
    Gs_OrderBy = "Order by Voucher_no"
    MyLookupOLDB.Caption = "Vouchers"
    MyLookupOLDB.Show 1
    If Len(txtvchrno) > 0 Then txtvchrno_KeyDown vbKeyReturn, vbKeyShift
End Sub


Private Sub Delete_menu_Click()
Mode = DentMode(Mode, 3, PR_GlRef, Me, txtvchrType, txtvchrType, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
InitializeGrid

End Sub

Private Sub DTPChkDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then TxtRemarks.SetFocus
End Sub

Private Sub edit_menu_sub_Click()
Mode = DentMode(Mode, 2, PR_GlRef, Me, txtvchrType, txtvchrType, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
InitializeGrid
End Sub

Private Sub New_Menu_Click()
Mode = DentMode(Mode, 1, PR_GlRef, Me, txtvchrType, txtvchrType, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
InitializeGrid
End Sub

Private Sub Save_menu_Click()
Mode = DentMode(Mode, 4, PR_GlRef, Me, txtvchrType, txtvchrType, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)


End Sub

Private Sub txtbranchcode_Change()
If txtbranchcode = "" Then
    txtbranchname = ""
End If
End Sub

Private Sub txtcrncycode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 If Lastkey(KeyCode) And txtCrncyCode.Text <> "" Then
         txtCrncyCode = UCase(txtCrncyCode.Text)
         lb_found = MySeek(txtCrncyCode, "Crncy_Code", PR_Crncy)
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtCrncyCode.SetFocus
         Else
             Text1 = PR_Crncy("Crncy_Descrip")
             If txtCrncyCode <> "PKR" Then
                txtexrate.SetFocus
             Else
                GrdGRN.SetFocus
             End If
         End If
 ElseIf KeyCode = vbKeyPageUp Then
         TxtRemarks.SetFocus
 ElseIf KeyCode = vbKeyF12 Then
    CmdCurrency_Click
 End If
End Sub

Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtvchrType
    Set PO_DESC = txtVchrDesc
    
    Gs_SQL = "Select VchrType 'Voucher Type',  VchrDescrip 'Description' from GlVchrType"
    Gs_FindFld = "VchrDescrip"
    Gs_Subon = True
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by VchrDescrip"
    MyLookupOLDB.Caption = "Voucher Types"
    MyLookupOLDB.Show 1
 
    If Len(txtvchrType) > 0 Then txtVchrType_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command4_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbranchcode
    Set PO_DESC = txtbranchname
    GoTop PR_Branch
    MyLookup.Caption = "Branches"
    MyLookup.FillGrid PR_Branch, "BranchCode", "BranchDesc", txtbranchcode.MaxLength
    MyLookup.Show 1
    
    If Len(txtbranchcode) > 0 Then txtBranchCode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF11 Then
       Mode = DentMode(Mode, 4, PR_GlTrans, FrmglTrans, txtvchrType, txtvaluedate, "X", "X", 3, "X", "X", 1, False, Toolbar1)
  End If
End Sub

Private Sub Form_Load()
    Dim ln_cnt As Integer
    Dim ls_PrvAlia As String
    Dim ls_PFields As String
    Dim SqlStr As New ADODB.Command
    txtbranchcode = Gs_BranchCode
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

    'PR_GlTrans.Open "SELECT Gl_Trans.*, BranchCode+VchrType + Voucher_no+ ltrim(str(month(value_date))) AS FindField FROM Gl_Trans WHERE CompCode = '" & Gs_compcode & "' and (value_date >= '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' and value_date <= '" & Format(Gs_FnEndPeriod, "YYYY/MM/DD") & "')  order by Value_Date,BranchCode,Vchrtype,voucher_no,SerialNo ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
    'PR_GlRef.Open "SELECT Gl_Ref.*, BranchCode+VchrType + Voucher_No + ltrim(str(month(value_date))) AS FindField FROM Gl_Ref WHERE CompCode = '" & Gs_compcode & "' and (value_date >= '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' and value_date <= '" & Format(Gs_FnEndPeriod, "YYYY/MM/DD") & "')  order by value_date,BranchCode,VchrType,Voucher_No", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
    PR_Crncy.Open "Select * From SysCurrency Order By Crncy_Code", gc_dbcon, adOpenStatic, adLockReadOnly, 1

    PI_SrNo = 0
    PI_CurRow = 0
    InitializeGrid
    txtCrncyCode = "PKR"
    txtvaluedate = Date
    DTPmdate = Date
    DTPChkDate = Date
    txtbranchcode = Gs_BranchCode
    
    If MySeek(txtbranchcode, "Branchcode", PR_Branch) Then
     txtbranchname = PR_Branch("BranchDesc")
    End If
    
  If Gs_compcode = "002" Then
    chkapproved.Enabled = False
  Else
    chkapproved.Enabled = True
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_Branch.Close
    PR_VchType.Close
    PR_Crncy.Close
End Sub

Private Sub txtBranchCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If Lastkey(KeyCode) And txtbranchcode.Text <> "" Then
         txtbranchcode = DoPad(txtbranchcode, 3)
         lb_found = MySeek(txtbranchcode.Text, "BranchCode", PR_Branch)
        
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtbranchcode.SetFocus
         Else
         txtbranchname = PR_Branch("BranchDesc")

         If txtvchrType = "" Then txtvchrType.SetFocus
             If txtvchrType <> "" Then txtVchrType_KeyDown vbKeyReturn, vbKeyShift
         End If
 ElseIf KeyCode = vbKeyF12 Then
     Call Command4_Click
 End If
End Sub
Private Sub txtexrate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then GrdGRN.SetFocus
If KeyCode = vbKeyPageUp Then txtCrncyCode.SetFocus
End Sub

Private Sub txtinstrument_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If Txtinstrument = "" Then
          Call SetErr("Cannot be empty", vbCritical)
          Txtinstrument.SetFocus
       Else
          DTPChkDate.SetFocus
       End If
    ElseIf KeyCode = vbKeyPageUp Then
        txtvaluedate.SetFocus
    End If
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TxtRemarks = Replace(TxtRemarks, "'", "")
            GrdGRN.SetFocus
            
    End If
    If KeyCode = vbKeyPageUp Then Txtinstrument.SetFocus
End Sub

Private Sub txtvaluedate_KeyDown(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = vbKeyReturn Then
        If (Format(DateValue(txtvaluedate.Value), "YYYY/MM/DD") < Format(DateValue(Gs_Fnperiod), "YYYY/MM/DD")) Or (Format(DateValue(txtvaluedate.Value), "YYYY/MM/DD") > Format(DateValue(Gs_FnEndPeriod), "YYYY/MM/DD")) Then
            Call SetErr("Invalid Period's Transaction.", vbCritical)
            txtvaluedate.SetFocus
        Else
            If txtbranchcode = "" Then
               txtbranchcode.SetFocus
            Else
              If txtvchrType.Enabled Then txtvchrType.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtvchrno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And txtvchrno <> "" Then
        txtvchrno = DoPad(txtvchrno.Text, 10)
        If Val(txtvchrno.Text) = 0 Then
           Call SetErr("Invalid Voucher No.", vbCritical)
           txtvchrType.SetFocus
        Else
          ls_VchNo = txtvchrno
          If Mode = "E" Then DTPmdate.SetFocus
           LoadVchrRef
        End If
    ElseIf KeyCode = vbKeyReturn And txtvchrno = "" Then
        Command2_Click
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
                    ls_DAccount = Trim(PR_VchType("AccountNo") & "")
                       If Mode = "A" Then
                         If PR_VchType.Fields("VchrFrequency") = "1" Then
                            pr_dumy.Open "select max(voucher_no) as voucherno from gl_ref where compcode = '" & Gs_compcode & "' and branchcode = '" & txtbranchcode & "' and vchrtype = '" & txtvchrType & "' and month(value_date) = " & Month(txtvaluedate) & " and year(value_date) = " & Year(txtvaluedate) & "", gc_dbcon, adOpenStatic, adLockReadOnly, 1
                            If Not pr_dumy.EOF Then
                                txtvchrno = DoPad(Trim(str(Val(0 & pr_dumy("voucherno")) + 1)), 10)
                            Else
                                txtvchrno = DoPad(Trim(str(Val(1))), 10)
                            End If
                            pr_dumy.Close
                         Else
                            pr_dumy.Open "select max(voucher_no) as voucherno from gl_ref where compcode = '" & Gs_compcode & "' and branchcode = '" & txtbranchcode & "' and vchrtype = '" & txtvchrType & "' and value_date >= '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' and value_date <= '" & Format(Gs_FnEndPeriod, "YYYY/MM/DD") & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
                            If Not pr_dumy.EOF Then
                                txtvchrno = DoPad(Trim(str(Val(0 & pr_dumy("voucherno")) + 1)), 10)
                            Else
                                txtvchrno = DoPad(Trim(str(Val(1))), 10)
                            End If
                            pr_dumy.Close
                            
                         End If
                            ln_OrgVchNo = Val(txtvchrno)
                         
                         If PR_VchType.Fields("VchrBase") = "B" Then Txtinstrument.SetFocus
                         If PR_VchType.Fields("VchrBase") <> "B" Then TxtRemarks.SetFocus
                      
                      Else
                         txtvchrno.SetFocus
                    End If
                End If
             
  ElseIf KeyCode = vbKeyReturn And txtvchrType = "" Then
   Call cmdLookup_Click
  End If
  End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 If Val(txtTDr) <> Val(txtTCr) And Button.Index = 4 Then
    Call SetErr("Voucher is Out of Balance.", vbCritical)
 Else
    If PB_BlnkTrns And Range(Button.Index, 2, 3) Then
       Call SetErr("Data not found.", vbCritical)
       Mode = ""
    Else
       
       If Button.Index = 5 And Mode <> "D" Then Call setprint
       If Button.Index = 7 Then txtCrncyCode = "PKR"
    End If
       Mode = DentMode(Mode, Button.Index, PR_GlTrans, FrmglTrans, txtvaluedate, txtvchrType, "X", "X", 3, "X", "X", 1, False, Toolbar1)
       If Button.Index <> 4 Then InitializeGrid
       
 End If
 If Button.Index = 5 And txtvchrType <> "" And txtvchrno <> "" Then Call setprint
    
End Sub
Private Sub setprint()
On Error GoTo LocalErr
Dim ls_BranchName As String
 


If ls_VchNo <> "" Then
         If MySeek(txtbranchcode, "BranchCode", PR_Branch) Then ls_BranchName = PR_Branch("BranchDesc")
   With rptVoucher
        
        .ReportFileName = App.Path & Gs_GlRepoPath & "\Vchr_Print.RPT"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = '" & ls_VchDesc & "'"
        .Formulas(5) = "BranchName = '" & txtbranchcode + "-" + ls_BranchName & "'"
        .SelectionFormula = "{Gl_Trans.Voucher_No} = '" & Trim(txtvchrno) & "' and {Gl_Trans.BranchCode} = '" & txtbranchcode & "'"
        .SelectionFormula = .SelectionFormula & " and {Gl_Trans.VchrType} = '" & Trim(txtvchrType) & "'"
        .SelectionFormula = .SelectionFormula & " and {Gl_Trans.CompCode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {Gl_Trans.Value_Date} = Date(" & Year(ld_VluDate) & "," & Month(ld_VluDate) & "," & Day(ld_VluDate) & ")"
        .Formulas(2) = "Sig1 = '" & Gc_UserName & "'"
        .Formulas(3) = "Sig2 = '" & Gs_Sign2 & "'"
        .Formulas(4) = "Sig3 = '" & Gs_Sign3 & "'"
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
        .FormatString = "Sr# |<Account No|<Account Name|<Account Narration|<Debit Amount|<Credit Amount|<empcode"
        .ColWidth(1) = 1500
        .ColWidth(2) = 3000
        .ColWidth(3) = 4000
        .ColWidth(4) = 1300
        .ColAlignment(4) = 7
        .ColWidth(5) = 1300
        .ColAlignment(5) = 7
        .ColWidth(6) = 0
        .Redraw = True
    End With
    PI_SrNo = 0
    ls_DAccount = ""
End Sub


Private Sub TotalAmount()
    Dim ln_cnt As Integer
    txtTDr = 0
    txtTCr = 0
    With GrdGRN
        For ln_cnt = 1 To .Rows - 1
            .TextMatrix(ln_cnt, 0) = ln_cnt
            txtTDr = txtTDr + Val(.TextMatrix(ln_cnt, 4))
            txtTCr = txtTCr + Val(.TextMatrix(ln_cnt, 5))
            PI_SrNo = ln_cnt
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

If Trim(txtTDr) = "" Or Trim(txtTCr) = "" Then Exit Sub

'ln_FrmCount = ln_FrmCount + 1
'If ln_FrmCount > 1 Then Exit Sub

lb_Vstat = False
ls_VchType = txtvchrType
ls_VchNo = txtvchrno
ld_VluDate = txtvaluedate
ls_VchDesc = txtVchrDesc

ls_TFields = Replace(ls_TFields, "+", ",")
'ld_PrvDate = txtvaluedate.Value

'If Mode <> "A" Then
'    ls_sql = "Select *  FROM Gl_Ref WHERE CompCode = '" & Gs_compcode & "' And BranchCode = '" & txtbranchcode & "' AND Voucher_No = '" & txtvchrno & "' AND VchrType = '" & txtvchrType & "' and value_date = '" & Format(txtvaluedate.Value, "YYYY/MM/DD") & "'"
'     If PR_Dumy.State = 1 Then PR_Dumy.Close
'     PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
'     If PR_Dumy.EOF Then
'     Call MsgBox("You have changed value date of your voucher,voucher not saved", vbCritical)
'     PR_Dumy.Close
'     Exit Sub
'     End If
'     PR_Dumy.Close
'End If

gc_dbcon.BeginTrans
     Select Case Mode
           Case "D"
              ' Delete Detail of Voucher
              ls_sql = "DELETE FROM Gl_Trans WHERE CompCode = '" & Gs_compcode & "' And BranchCode = '" & txtbranchcode & "' AND Voucher_No = '" & txtvchrno & "' AND VchrType = '" & txtvchrType & "' and month(value_date) = " & Month(txtvaluedate.Value) & " and year(value_date) = " & Year(txtvaluedate.Value) & ""
              gc_dbcon.Execute ls_sql
              
              ls_sql = "DELETE FROM Gl_Ref WHERE CompCode = '" & Gs_compcode & "' And BranchCode = '" & txtbranchcode & "' AND Voucher_No = '" & txtvchrno & "' AND VchrType = '" & txtvchrType & "' and month(value_date) = " & Month(txtvaluedate.Value) & " and year(value_date) = " & Year(txtvaluedate.Value) & ""
              gc_dbcon.Execute ls_sql
                
           Case Else
                If Mode = "E" Then
                 ' Delete Reference of Voucher
                   ls_sql = "DELETE FROM Gl_Ref WHERE CompCode = '" & Gs_compcode & "' And BranchCode = '" & txtbranchcode & "' AND Voucher_No = '" & txtvchrno & "' AND VchrType = '" & txtvchrType & "' and month(value_date) = " & Month(txtvaluedate.Value) & " and year(value_date) = " & Year(txtvaluedate.Value) & ""
                   gc_dbcon.Execute ls_sql
                   
                   ls_sql = "DELETE FROM Gl_Trans WHERE CompCode = '" & Gs_compcode & "' And BranchCode = '" & txtbranchcode & "' AND Voucher_No = '" & txtvchrno & "' AND VchrType = '" & txtvchrType & "' and month(value_date) = " & Month(txtvaluedate.Value) & " and year(value_date) = " & Year(txtvaluedate.Value) & ""
                   gc_dbcon.Execute ls_sql
                Else
                  
                  If Mode = "A" Then
                         PR_VchType.Filter = "VchrType = '" & txtvchrType & "'"
                        If PR_VchType.Fields("VchrFrequency") = "1" Then
                            pr_dumy.Open "select max(voucher_no) as voucherno from gl_ref where compcode = '" & Gs_compcode & "' and branchcode = '" & txtbranchcode & "' and vchrtype = '" & txtvchrType & "' and month(value_date) = " & Month(txtvaluedate) & " and year(value_date) = " & Year(txtvaluedate) & "", gc_dbcon, adOpenStatic, adLockReadOnly, 1
                            If Not pr_dumy.EOF Then
                                txtvchrno = DoPad(Trim(str(Val(0 & pr_dumy("voucherno")) + 1)), 10)
                            Else
                                txtvchrno = DoPad(Trim(str(Val(1))), 10)
                            End If
                            pr_dumy.Close
                        Else
                            pr_dumy.Open "select max(voucher_no) as voucherno from gl_ref where compcode = '" & Gs_compcode & "' and branchcode = '" & txtbranchcode & "' and vchrtype = '" & txtvchrType & "' and value_date >= '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' and value_date <= '" & Format(Gs_FnEndPeriod, "YYYY/MM/DD") & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
                            If Not pr_dumy.EOF Then
                                txtvchrno = DoPad(Trim(str(Val(0 & pr_dumy("voucherno")) + 1)), 10)
                            Else
                                txtvchrno = DoPad(Trim(str(Val(1))), 10)
                            End If
                            pr_dumy.Close
                            
                        End If
                         
                        If ln_TmpVchNo > ln_OrgVchNo And ln_OrgVchNo = Val(txtvchrno) Then
                            lb_Vstat = True
                            txtvchrno = DoPad(Trim(str(ln_TmpVchNo)), 10)
                        End If
                PR_VchType.Filter = adFilterNone
                End If
                 
                End If
                ' Save References of Voucher
                If Mode = "E" Then
                    ld_valueDate = DTPmdate
                    ld_VluDate = DTPmdate
                Else
                    ld_valueDate = txtvaluedate
                    ld_VluDate = txtvaluedate
                End If
                
                
                
                
                 ls_sql = "INSERT into Gl_Ref(compcode,BranchCode,Value_Date,Trans_Date, Voucher_No, VchrType, Vchr_Remarks,InstrumentNo,CrncyCode,ExchgRate,userid,adddate,addtime,chkdate) VALUES ('" & Gs_compcode & "','" & txtbranchcode & "','" & Format(ld_valueDate, "YYYY/MM/DD") & "','" & Format(txtTransDate, "YYYY/MM/DD") & "','" & txtvchrno & "','" & txtvchrType & "','" & TxtRemarks & "','" & Txtinstrument & "','" & txtCrncyCode & "'," & Val(txtexrate) & ",'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & Format(DTPChkDate, "YYYY/MM/DD") & "')"
                 gc_dbcon.Execute ls_sql
                 
                ' Save Details of Voucher
                    With GrdGRN
                       For ln_cnt = 1 To .Rows - 1
                          If Len(Trim(.TextMatrix(ln_cnt, 1))) > 0 Then
                           ls_sql = "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, ExchgRate,userid,adddate,addtime,Acct_Nirration,AcctName) VALUES ('" & Gs_compcode & "','" & txtbranchcode & "','" & .TextMatrix(ln_cnt, 1) & "'," & .TextMatrix(ln_cnt, 0) & ",'" & Format(ld_valueDate, "YYYY/MM/DD") & "','" & txtvchrno & "','" & txtvchrType & "'," & Val(.TextMatrix(ln_cnt, 4)) & "," & Val(.TextMatrix(ln_cnt, 5)) & "," & Val(txtexrate) & ",'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & Trim(.TextMatrix(ln_cnt, 3)) & "','" & Trim(.TextMatrix(ln_cnt, 2)) & "')"
                           gc_dbcon.Execute ls_sql
                          End If
                       Next
                     End With
                     
     End Select
     
gc_dbcon.CommitTrans
Call SetClear(Me)
ln_FrmCount = 0

PI_SrNo = 1
PI_CurRow = 1
InitializeGrid
On Error GoTo 0
If lb_Vstat Then Call SetErr("Your New Transaction Voucher No will be " + txtvchrno.Text, vbCritical)
If Mode <> "D" Then
   ls_opt = SetErr("Print Voucher ?.", vbYesNo)
   If ls_opt = vbYes Then Call setprint
End If
If Mode = "A" Then
Call txtVchrType_KeyDown(vbKeyReturn, vbKeyShift)
txtvchrType.SetFocus
Else
    txtvchrno = ""
    txtvchrType = ""
    txtVchrDesc = ""
    InitializeGrid
End If

Exit Sub

RollBack:
MsgBox Err.Description
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
gc_dbcon.RollbackTrans
ln_FrmCount = 0
On Error GoTo 0
End Sub

Public Function ChkInputs() As Boolean
TotalAmount
If Val(0 & txtTDr) > 0 And Val(0 & txtTCr) > 0 Then
    If Val(txtvchrno) > 0 And Len(Trim(TxtRemarks.Text)) > 0 And txtbranchcode <> " " And PI_SrNo > 0 And Val(0 & txtTDr) = Val(0 & txtTCr) Then
        If (Format(DateValue(txtvaluedate.Value), "YYYY/MM/DD") < Format(DateValue(Gs_Fnperiod), "YYYY/MM/DD")) Or (Format(DateValue(txtvaluedate.Value), "YYYY/MM/DD") > Format(DateValue(Gs_FnEndPeriod), "YYYY/MM/DD")) Then
            Call SetErr("Invalid Period's Transaction.", vbCritical)
            ChkInputs = False
        Else
            If Mode = "A" Or Mode = "E" Then
                res = MsgBox("Do You Want To Save Record", vbYesNo + vbInformation)
                If res = vbYes Then
                   ChkInputs = True
                Else
                   ChkInputs = False
                End If
            Else
              ChkInputs = True
            
            End If
        End If
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
Else
    Call SetErr(Gs_InvldMsg, vbCritical)
    ChkInputs = False
End If
End Function

Private Sub LoadVchrRef()
Dim lb_found As Boolean
    If PR_GlRef.State = 1 Then PR_GlRef.Close
    PR_GlRef.Open "Select * from gl_ref where compcode = '" & Gs_compcode & "' and BranchCode+VchrType + Voucher_no+ ltrim(str(month(value_date)))+ ltrim(str(year(value_date))) = '" & Trim(txtbranchcode) & Trim(txtvchrType) & Trim(txtvchrno) + LTrim(str(Month(txtvaluedate.Value))) + LTrim(str(Year(txtvaluedate.Value))) & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
    lb_found = PR_GlRef.EOF
      Select Case Mode
         Case "A"
            If lb_found Then
               Call MsgBox("Voucher already exist !!!", vbCritical)
               txtvchrType.SetFocus
            Else
               If PR_VchType.Fields("VchrBase") = "B" Then Txtinstrument.SetFocus
               If PR_VchType.Fields("VchrBase") <> "B" Then TxtRemarks.SetFocus
            End If
         Case Else
              If lb_found Then
                 Call SetErr(Gs_RecNFMsg, vbCritical)
                 InitializeGrid
                 txtvchrType.SetFocus
              Else
                 
                  txtTransDate = PR_GlRef("Trans_Date")
                  DTPmdate = txtvaluedate
                  If Not IsNull(PR_GlRef("ChkDate")) Then
                  DTPChkDate = PR_GlRef("ChkDate")
                  End If
                  Txtinstrument = PR_GlRef("InstrumentNo") & ""
                  TxtRemarks = PR_GlRef("Vchr_Remarks") & ""
                  txtCrncyCode = PR_GlRef("CrncyCode") & ""
                  txtexrate = Val(0 & PR_GlRef("ExchgRate"))
                  LoadVchrTrans
                  
              End If
      End Select
End Sub

Private Sub LoadVchrTrans()
Dim lb_found As Boolean
Dim ln_cnt   As Integer
Dim temp As String
    
InitializeGrid
'txtAcctName = ""
If Mode = "E" Then
    ls_VchType = txtvchrType
    ls_VchNo = txtvchrno
    ld_VluDate = txtvaluedate
    ls_VchDesc = txtVchrDesc
End If

    PR_GlTrans.Open "Select GL_Trans.*,gl_detail.acct_desc from GL_Trans left outer join gl_detail on gl_trans.compcode+gl_trans.accountno =gl_detail.compcode+gl_detail.accountno where gl_trans.compcode = '" & Gs_compcode & "' and gl_trans.BranchCode+gl_trans.VchrType + gl_trans.Voucher_no+ ltrim(str(month(value_date)))+ ltrim(str(year(value_date))) = '" & Trim(txtbranchcode) & Trim(txtvchrType) & Trim(txtvchrno) + LTrim(str(Month(txtvaluedate.Value))) + LTrim(str(Year(txtvaluedate.Value))) & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
    If Not PR_GlTrans.EOF Then
        With GrdGRN
            Do While Not PR_GlTrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = PR_GlTrans("SerialNo")
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = PR_GlTrans("AccountNo")
                .TextMatrix(.Row, 2) = Trim(PR_GlTrans("acct_desc") & "")
                .TextMatrix(.Row, 3) = PR_GlTrans("Acct_Nirration") & ""
                .TextMatrix(.Row, 4) = PR_GlTrans("Dr_Amount")
                .TextMatrix(.Row, 5) = PR_GlTrans("Cr_Amount")
                
                
                .Rows = .Rows + 1
                PR_GlTrans.MoveNext
                If PR_GlTrans.EOF Or PR_GlTrans.BOF Then Exit Do
             Loop
            
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
        TotalAmount
        If TxtRemarks.Enabled Then TxtRemarks.SetFocus
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
              .TextMatrix(.Row, 3) = TxtRemarks
              End If
      
             
             
         pr_dumy.Close
         End If
       Else
           Call GrdGRN_KeyDown(112, vbKeyShift)
       End If
      ElseIf .Col = 3 Then
      .CellBackColor = vbWindowBackground
      .Col = .Col + 1
       ElseIf .Col = 4 Then
      .CellBackColor = vbWindowBackground
      .Col = .Col + 1
      .CellBackColor = vbHighlight
       
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
   If .Col = 1 Or .Col = 4 Or .Col = 5 Then
   If Len(Trim(.Text)) <> 0 Then  'If current cell is not empty then...
      .Text = Left(.Text, (Len(.Text) - 1)) 'Removing a character from the right side of the FlexGrid cell's text
   End If
   End If
End With
End If

If KeyAscii <> 27 And KeyAscii <> 8 Then
    With GrdGRN
      
      If .Col = 1 Then
        If .CellBackColor = vbHighlight Then
         .Text = "": .CellBackColor = vbWindowBackground
        End If
        .Text = .Text & Chr(KeyAscii) 'Reset Value in Cell and Append the pressed character to the right.
        
      ElseIf .Col = 3 Then
        TXTBARCODE.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
       .TextMatrix(.Row, 3) = Chr(KeyAscii)
        TXTBARCODE.Text = .TextMatrix(.Row, 3)
        TXTBARCODE.Visible = True
        TXTBARCODE.SelStart = Len(TXTBARCODE)
        ClickRow = .Row
        TXTBARCODE.SetFocus
        '.CellBackColor = vbWindowBackground
        '.Text = .Text & Chr(KeyAscii) 'Reset Value in Cell and Append the pressed character to the right.
      ElseIf .Col = 4 Or .Col = 5 Then
        If .CellBackColor = vbHighlight Then
                .Text = "": .CellBackColor = vbWindowBackground
        End If
         .Text = .Text & Chr(KeyAscii)
          If Not IsNumeric(.Text) Then
          .Text = ""
           Call MsgBox("Enter Numeric entry !!!", vbCritical)
           Exit Sub
          End If
          
          If .Col = 4 Then
          .TextMatrix(.Row, 5) = 0
          End If
          If .Col = 5 Then
          .TextMatrix(.Row, 4) = 0
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
If GrdGRN.Col = 3 Then
With GrdGRN

        TXTBARCODE.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
        TXTBARCODE.Text = .TextMatrix(.Row, 3)
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
'On Error GoTo LocalErr
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

Private Sub TXTBARCODE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyTab Then
TXTBARCODE_LostFocus
'ElseIf KeyCode = vbKeyRight Then
'TboxCol = True
'TXTBARCODE_LostFocus

End If
End Sub

Private Sub TXTBARCODE_LostFocus()
With GrdGRN
If ClickRow <> "" Then
.TextMatrix(ClickRow, 3) = TXTBARCODE.Text
 TXTBARCODE.Text = ""
 ClickRow = ""
 End If
 .CellBackColor = vbWindowBackground
  TXTBARCODE.Visible = False
  If TboxCol = True Then
  .Col = 3
  .CellBackColor = vbHighlight
  .SetFocus

  TboxCol = False
  Else
 .SetFocus
 .Col = 4
 .CellBackColor = vbHighlight
 End If
 
 
End With


End Sub


VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInventoryTransfer 
   Caption         =   "Inventory Transfer"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8355
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmInventoryTransfer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8355
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
      Height          =   2205
      Left            =   45
      TabIndex        =   1
      Top             =   570
      Width           =   8280
      Begin VB.TextBox txtbranchname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2370
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   31
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   150
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   2025
         Picture         =   "FrmInventoryTransfer.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   150
         Width           =   315
      End
      Begin VB.TextBox TxtSiteDesc1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2370
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1245
         Width           =   2085
      End
      Begin VB.TextBox txtSiteID1 
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
         MaxLength       =   3
         TabIndex        =   26
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1245
         Width           =   570
      End
      Begin VB.CommandButton Command12 
         Height          =   315
         Left            =   2025
         Picture         =   "FrmInventoryTransfer.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1245
         Width           =   315
      End
      Begin VB.TextBox TxtBinDesc1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   6150
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1215
         Width           =   2055
      End
      Begin VB.TextBox TxtBinID1 
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
         Left            =   5355
         MaxLength       =   3
         TabIndex        =   23
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1230
         Width           =   420
      End
      Begin VB.CommandButton Command11 
         Height          =   315
         Left            =   5805
         Picture         =   "FrmInventoryTransfer.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1200
         Width           =   315
      End
      Begin VB.CommandButton Command8 
         Height          =   315
         Left            =   5805
         Picture         =   "FrmInventoryTransfer.frx":0760
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   855
         Width           =   315
      End
      Begin VB.TextBox txtbinID 
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
         Left            =   5355
         MaxLength       =   3
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   870
         Width           =   420
      End
      Begin VB.TextBox TxtBinDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   6135
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   855
         Width           =   2070
      End
      Begin VB.CommandButton Command7 
         Height          =   315
         Left            =   2025
         Picture         =   "FrmInventoryTransfer.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   870
         Width           =   330
      End
      Begin VB.TextBox TxtSiteID 
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
         MaxLength       =   3
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   870
         Width           =   570
      End
      Begin VB.TextBox TxtsiteDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2370
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   870
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker txtvaluedate 
         Height          =   315
         Left            =   5355
         TabIndex        =   12
         Top             =   150
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   63504385
         CurrentDate     =   37580
      End
      Begin VB.TextBox txtDepartmentdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   6270
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   510
         Width           =   1935
      End
      Begin VB.TextBox txtDepartmentCode 
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
         Left            =   5355
         MaxLength       =   6
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   525
         Width           =   600
      End
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   5970
         Picture         =   "FrmInventoryTransfer.frx":0A44
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   510
         Width           =   285
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   75
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   180
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox TxtRemarks 
         Height          =   510
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1605
         Width           =   6750
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   315
         Left            =   2550
         Picture         =   "FrmInventoryTransfer.frx":0BB6
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   525
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
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   525
         Width           =   1095
      End
      Begin Crystal.CrystalReport rptVoucher 
         Left            =   8370
         Top             =   180
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
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
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   7470
         Top             =   -180
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
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
      Begin MSMask.MaskEdBox txtbranchcode 
         Height          =   315
         Left            =   1440
         TabIndex        =   32
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         ToolTipText     =   "Default Currency"
         Top             =   165
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin VB.Label Label28 
         Caption         =   "Branch # :"
         Height          =   255
         Left            =   675
         TabIndex        =   33
         Top             =   180
         Width           =   765
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Site ID To :"
         Height          =   255
         Left            =   465
         TabIndex        =   29
         Top             =   1275
         Width           =   975
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Bin ID :"
         Height          =   255
         Left            =   4725
         TabIndex        =   28
         Top             =   1260
         Width           =   630
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Bin ID :"
         Height          =   255
         Left            =   4710
         TabIndex        =   21
         Top             =   885
         Width           =   630
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Site ID From :"
         Height          =   255
         Left            =   450
         TabIndex        =   17
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Depertment :"
         Height          =   255
         Left            =   4005
         TabIndex        =   13
         Top             =   525
         Width           =   1350
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks :"
         Height          =   255
         Left            =   300
         TabIndex        =   7
         Top             =   1620
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Transfer #  :"
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   525
         Width           =   1275
      End
      Begin VB.Label label2 
         Caption         =   "Issue Date :"
         Height          =   255
         Left            =   4485
         TabIndex        =   5
         ToolTipText     =   "Enter Value Date"
         Top             =   180
         Width           =   930
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8355
      _ExtentX        =   14737
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
               Picture         =   "FrmInventoryTransfer.frx":0D28
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryTransfer.frx":117C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryTransfer.frx":15D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryTransfer.frx":1A24
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryTransfer.frx":1E78
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryTransfer.frx":22CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmInventoryTransfer.frx":2A20
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3915
      Left            =   45
      TabIndex        =   34
      Top             =   2670
      Width           =   8280
      Begin VB.TextBox txttotalamount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6840
         MaxLength       =   11
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   3495
         Width           =   1320
      End
      Begin VB.TextBox txtnoofitems 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   0
         MaxLength       =   50
         TabIndex        =   36
         Top             =   -360
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtitemname 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   35
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   3510
         Width           =   6225
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdGRN 
         Height          =   3210
         Left            =   45
         TabIndex        =   39
         Top             =   195
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   5662
         _Version        =   393216
         RowHeightMin    =   300
         BackColorSel    =   16777215
         ForeColorSel    =   0
         GridColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
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
      Begin VB.Label Label11 
         Caption         =   " Total :"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6345
         TabIndex        =   38
         Top             =   3525
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmInventoryTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object


Dim ls_transtype As String
Dim ls_transcodeinv As String

Dim pr_dumy As New Recordset
Dim PR_UOM As New Recordset

Dim PR_ICTransfer As New Recordset

Dim PR_IcItem As New Recordset
Dim PR_Branch As New Recordset

Private Function maxtranscode() As String
pr_dumy.Open "select max(transcode) as transcode from Ic_TransferMaster where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 10)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 10)
End If
pr_dumy.Close
End Function
Private Function maxtranscode1() As String
pr_dumy.Open "select max(transcode) as transcode from Ic_TransMaster where compcode = '" & Gs_compcode & "' and transtype = '" & ls_transtype & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode1 = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 10)
Else
    maxtranscode1 = DoPad(Trim(str(Int(1))), 10)
End If
pr_dumy.Close
End Function



Private Sub Command4_Click()
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
        If Mode = "A" Then
            txtDepartmentCode.SetFocus
        Else
            txttransno.SetFocus
        End If
     End If
  ElseIf KeyCode = vbKeyF12 Then

     Command4_Click
  ElseIf KeyCode = vbKeyReturn And txtbranchcode = "" Then
      txtbranchname = ""
  End If
End Sub


Private Sub Command7_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = TxtSiteID
    Set PO_DESC = TxtsiteDesc
    Gs_SQL = "Select SiteCode, Description from IC_Sites "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Company Sites"
    MyLookupOLDB.Show 1
    
    If TxtSiteID <> "" Then Call txtSiteID_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub txtSiteID_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(TxtSiteID) <> "" And KeyCode = vbKeyReturn Then
        TxtSiteID = DoPad(TxtSiteID, 3)
        pr_dumy.Open "Select * from IC_Sites where Compcode  = '" & Gs_compcode & "' and Sitecode = '" & TxtSiteID & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Site Code not found !!!", vbCritical)
            TxtSiteID = ""
            TxtsiteDesc = ""
            TxtSiteID.SetFocus
        Else
            TxtsiteDesc = pr_dumy("Description")
            txtbinID.SetFocus
            
        End If
        pr_dumy.Close

ElseIf Trim(TxtSiteID) = "" And KeyCode = vbKeyReturn Then
        TxtSiteID = ""
        TxtsiteDesc = ""
End If

End Sub
Private Sub Command12_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtSiteID1
    Set PO_DESC = TxtSiteDesc1
    Gs_SQL = "Select SiteCode, Description from IC_Sites "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Company Sites"
    MyLookupOLDB.Show 1
    
    If txtSiteID1 <> "" Then Call txtSiteID1_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub txtSiteID1_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(txtSiteID1) <> "" And KeyCode = vbKeyReturn Then
        txtSiteID1 = DoPad(txtSiteID1, 3)
        pr_dumy.Open "Select * from IC_Sites where Compcode  = '" & Gs_compcode & "' and Sitecode = '" & txtSiteID1 & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Site Code not found !!!", vbCritical)
            txtSiteID1 = ""
            TxtSiteDesc1 = ""
            txtSiteID1.SetFocus
        Else
            TxtSiteDesc1 = pr_dumy("Description")
            TxtBinID1.SetFocus
            
        End If
        pr_dumy.Close

ElseIf Trim(txtSiteID1) = "" And KeyCode = vbKeyReturn Then
        txtSiteID1 = ""
        TxtSiteDesc1 = ""
End If

End Sub
Private Sub Command5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtDepartmentCode
    Set PO_DESC = txtDepartmentdesc
    Gs_SQL = "Select DeptCode, Description from IC_Departments "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Departments"
    MyLookupOLDB.Show 1
    
    If txtDepartmentCode <> "" Then Call txtDepartmentCode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub txtDepartmentCode_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(txtDepartmentCode) <> "" And KeyCode = vbKeyReturn Then
        txtDepartmentCode = DoPad(txtDepartmentCode, 6)
        pr_dumy.Open "Select * from IC_Departments where Compcode  = '" & Gs_compcode & "' and Deptcode = '" & txtDepartmentCode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Site Code not found !!!", vbCritical)
            txtDepartmentCode = ""
            txtDepartmentdesc = ""
            txtDepartmentCode.SetFocus
        Else
            txtDepartmentdesc = pr_dumy("Description")
            TxtSiteID.SetFocus
            
        End If
        pr_dumy.Close

ElseIf Trim(txtDepartmentCode) = "" And KeyCode = vbKeyReturn Then
        txtDepartmentCode = ""
        txtDepartmentdesc = ""
End If

End Sub



Private Sub Command8_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbinID
    Set PO_DESC = TxtBinDesc
    Gs_SQL = "Select BinCode, Description from IC_SitesBins "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where sitecode ='" & TxtSiteID & "' and Compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Company Site Bins"
    MyLookupOLDB.Show 1
    
    If txtbinID <> "" Then Call txtBinID_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub txtBinID_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(txtbinID) <> "" And KeyCode = vbKeyReturn Then
        txtbinID = DoPad(txtbinID, 3)
        pr_dumy.Open "Select * from IC_SitesBins where  sitecode ='" & TxtSiteID & "' and compcode = '" & Gs_compcode & "'  and bincode = '" & txtbinID & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Bin Code not found !!!", vbCritical)
            txtbinID = ""
            TxtBinDesc = ""
            txtbinID.SetFocus
        Else
            TxtBinDesc = pr_dumy("Description")
            txtSiteID1.SetFocus
            
        End If
        pr_dumy.Close

ElseIf Trim(txtbinID) = "" And KeyCode = vbKeyReturn Then
        txtbinID = ""
        TxtBinDesc = ""
End If

End Sub
Private Sub Command11_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = TxtBinID1
    Set PO_DESC = TxtBinDesc1
    Gs_SQL = "Select BinCode, Description from IC_SitesBins "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where sitecode ='" & txtSiteID1 & "' and Compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Company Site Bins"
    MyLookupOLDB.Show 1
    
    If TxtBinID1 <> "" Then Call txtBinID1_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub txtBinID1_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(TxtBinID1) <> "" And KeyCode = vbKeyReturn Then
        TxtBinID1 = DoPad(TxtBinID1, 3)
        pr_dumy.Open "Select * from IC_SitesBins where  sitecode ='" & txtSiteID1 & "' and compcode = '" & Gs_compcode & "'  and bincode = '" & TxtBinID1 & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Bin Code not found !!!", vbCritical)
            TxtBinID1 = ""
            TxtBinDesc1 = ""
            TxtBinID1.SetFocus
        Else
            TxtBinDesc1 = pr_dumy("Description")
            If TxtRemarks.Enabled Then TxtRemarks.SetFocus
            
        End If
        pr_dumy.Close

ElseIf Trim(TxtBinID1) = "" And KeyCode = vbKeyReturn Then
        TxtBinID1 = ""
        TxtBinDesc1 = ""
End If

End Sub





Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txttransno
    Set PO_DESC = Text1
    Gs_SQL = "Select Transcode, Transdate,Accountcode from IC_TransferMaster "
    Gs_FindFld = "Transcode"
    Gs_OrderBy = "Order by Transcode"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Transfer Entery"
    MyLookupOLDB.Show 1
    
    If txttransno <> "" Then Call txtTransNo_KeyDown(vbKeyReturn, vbKeyShift)

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then Mode = DentMode(Mode, 4, PR_ICTransfer, Me, txttransno, txtvaluedate, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 0, False, Toolbar1)
End Sub

Private Sub Form_Load()

 
  SetToolBar(1) = chkRights("ICISUSTP01")
  SetToolBar(2) = chkRights("ICISUSTP02")
  SetToolBar(3) = chkRights("ICISUSTP03")
  SetToolBar(4) = chkRights("ICISUSTP04")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  'PR_IcItem.Open "Select * from Ic_Item where compcode ='" & Gs_compcode & "' ", gc_dbcon, adOpenDynamic, adLockPessimistic, 1
 ' PR_ICTransfer.Open "Select * from Ic_TransferMaster where compcode ='" & Gs_compcode & "' and transtype in ('T')  order by Transcode", gc_dbcon, adOpenDynamic, adLockOptimistic
  

  txtvaluedate.Value = Date
 
  InitializeGrid

  PR_Branch.Open "Select * From SysBranch Where compcode = '" & Gs_compcode & "' Order By BranchCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
  txtbranchcode = Gs_BranchCode
  
  If MySeek(txtbranchcode, "Branchcode", PR_Branch) Then
   txtbranchname = PR_Branch("BranchDesc")
  End If
  
  
End Sub
Private Sub InitializeGrid()
    With GrdGRN
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Item Code|<Item Name|<UOM|<Qty|<Rate|<Total|<CBalQty|<AvgRate|<Avg Amount|<Remarks"
        .ColWidth(1) = 900
        .ColWidth(2) = 2300
        .ColWidth(3) = 1100
        .ColWidth(4) = 800
        .ColAlignment(4) = 7
        .ColWidth(5) = 800
        .ColAlignment(5) = 7
        .ColWidth(6) = 900
        .ColAlignment(6) = 7
        .ColWidth(7) = 800
        .ColAlignment(7) = 7
        .ColWidth(8) = 800
        .ColAlignment(8) = 7
        .ColWidth(9) = 1200
        .ColAlignment(9) = 7
        .ColWidth(10) = 2500
        .Redraw = True
    End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
'    PR_ICTransfer.Close
  '  PR_IcItem.Close
    PR_Branch.Close
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then GrdGRN.SetFocus

End Sub
Function checkvalidate() As Boolean
If Trim(txtitemcode) = "" Then
    Call MsgBox("Enter Item Code !!!", vbCritical)
    txtitemcode.SetFocus
    checkvalidate = False
ElseIf Val(txtqty) = 0 Then
    Call MsgBox("Enter Quantity !!!", vbCritical)
    txtqty.SetFocus
    checkvalidate = False
ElseIf Val(txtunitprice) = 0 Then
    Call MsgBox("Enter unit price !!!", vbCritical)
    txtunitprice.SetFocus
    checkvalidate = False
Else
    checkvalidate = True
End If
End Function


Private Sub txtTransNo_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And Len(txttransno.Text) > 0 Then
         If PR_ICTransfer.State = 1 Then PR_ICTransfer.Close
         txttransno.Text = DoPad(UCase(txttransno.Text), 10)
         PR_ICTransfer.Open "select * from IC_TransferMaster where compcode = '" & Gs_compcode & "' and Transcode = '" & txttransno & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
       Select Case Mode
            Case "A"
                If Not PR_ICTransfer.EOF Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   If txttransno.Enabled Then txttransno.SetFocus
                Else
                   txtvaluedate.SetFocus
                End If
            Case Else
                If PR_ICTransfer.EOF Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   txttransno.SetFocus
                Else
                   Call SetVal
                   LoadGRNTrans
                   If Mode <> "D" Then
                      txttransno.SetFocus
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
       InitializeGrid
       
    Else
       txttransno.SetFocus
       cmdLookup.Enabled = True
    End If
    If Button.Index = 7 Then
    InitializeGrid
    End If
    
    If PB_BlnkGRN And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_ICTransfer, Me, txttransno, txttransno, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
    End If
    If Mode = "A" Then
     '  txtVchrType = "JVS"
     '  Call txtVchrType_KeyDown(vbKeyReturn, vbKeyShift)
       txttransno = maxtranscode
       txtDepartmentCode.SetFocus
    End If
End Sub


Public Sub SaveValues()
'On Error GoTo RollBack
Dim ln_cnt As Integer
Dim ls_sql As String



gc_dbcon.BeginTrans

     Select Case Mode
           Case "D"
              gc_dbcon.Execute "DELETE FROM IC_TransMaster WHERE CompCode = '" & Gs_compcode & "' AND transfercode = '" & Trim(txttransno) & "'"
              gc_dbcon.Execute "DELETE FROM IC_Trans WHERE CompCode = '" & Gs_compcode & "' AND transfercode = '" & Trim(txttransno) & "'"
              gc_dbcon.Execute "DELETE FROM IC_TransferMaster WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
              gc_dbcon.Execute "DELETE FROM IC_Transfer WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
              
              
           Case Else
                If Mode = "E" Then
                    gc_dbcon.Execute "DELETE FROM IC_TransMaster WHERE CompCode = '" & Gs_compcode & "' AND transfercode = '" & Trim(txttransno) & "'"
                    gc_dbcon.Execute "DELETE FROM IC_Trans WHERE CompCode = '" & Gs_compcode & "' AND transfercode = '" & Trim(txttransno) & "'"
                    gc_dbcon.Execute "DELETE FROM IC_TransferMaster WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
                    gc_dbcon.Execute "DELETE FROM IC_Transfer WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
            
                End If

                If Mode = "A" Then
                    txttransno = maxtranscode
                End If
                      
                      ls_sql = "INSERT into IC_TransferMaster(Compcode,BranchCode, TransCode, TransDate, AccountCode, SiteID, BinID, SiteID1, BinID1, Remarks, TotalAmount)"
                      ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & txtbranchcode & "','" & Trim(txttransno) & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & txtDepartmentCode & "','" & TxtSiteID & "','" & txtbinID & "','" & txtSiteID1 & "','" & TxtBinID1 & "','" & RepApp(TxtRemarks) & "'," & txttotalamount & " )"
                      gc_dbcon.Execute ls_sql
                
                      With GrdGRN
                                For ln_cnt = 1 To .Rows - 1
                                    If .TextMatrix(ln_cnt, 1) <> "" Then
                                         ls_sql = "INSERT into IC_Transfer(Compcode,BranchCode, TransCode , ItemCode, Quantity, ItemRate, Amount, TaxAmount,AvgRate,Remarks,BalQty)"
                                         ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & txtbranchcode & "','" & Trim(txttransno) & "','" & Trim(.TextMatrix(ln_cnt, 1)) & "'," & Val(0 & .TextMatrix(ln_cnt, 4)) & "," & Val(0 & .TextMatrix(ln_cnt, 5)) & "," & Val(0 & .TextMatrix(ln_cnt, 6)) & ",0," & Val(0 & .TextMatrix(ln_cnt, 8)) & " , '" & Trim(.TextMatrix(ln_cnt, 10)) & "' ," & Val(0 & .TextMatrix(ln_cnt, 7)) & ")"
                                        gc_dbcon.Execute ls_sql
                                    End If
                                Next
                      End With
                 
                 
                      ls_transtype = "S"
                      ls_transcodeinv = maxtranscode1
                     
                      ls_sql = "INSERT into IC_TransMaster(Compcode,branchcode, TransCode, TransType, TransDate, AccountCode, SiteID, BinID, Remarks,SubTotal, TaxAmount, TotalAmount,JobNo,TransferCode)"
                      ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & txtbranchcode & "','" & Trim(ls_transcodeinv) & "','" & ls_transtype & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & txtDepartmentCode & "','" & TxtSiteID & "','" & txtbinID & "','" & RepApp(TxtRemarks) & "'," & Val(txttotalamount) & ",0," & Val(txttotalamount) & ",'ADJ','" & txttransno & "' )"
                      gc_dbcon.Execute ls_sql
                
                        With GrdGRN
                            For ln_cnt = 1 To .Rows - 1
                                If .TextMatrix(ln_cnt, 1) <> "" Then
                                 ls_sql = "INSERT into IC_Trans(Compcode,BranchCode, TransCode ,Transtype, ItemCode, Quantity, ItemRate, Amount, TaxAmount,AvgRate,Remarks,BalQty,TransferCode)"
                                 ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & txtbranchcode & "','" & Trim(ls_transcodeinv) & "','" & Trim(ls_transtype) & "','" & Trim(.TextMatrix(ln_cnt, 1)) & "'," & Val(0 & .TextMatrix(ln_cnt, 4)) & "," & Val(0 & .TextMatrix(ln_cnt, 5)) & "," & Val(0 & .TextMatrix(ln_cnt, 6)) & ",0," & Val(0 & .TextMatrix(ln_cnt, 8)) & " , '" & Trim(.TextMatrix(ln_cnt, 10)) & "' ," & Val(0 & .TextMatrix(ln_cnt, 7)) & " ,'" & txttransno & "')"
                                 gc_dbcon.Execute ls_sql
                                End If
                            Next
                        End With
                 
                 
                      ls_transtype = "P"
                      ls_transcodeinv = maxtranscode1

                      ls_sql = "INSERT into IC_TransMaster( Compcode,branchcode, TransCode, TransType, TransDate, AccountCode, SiteID, BinID, Remarks,SubTotal, TaxAmount, TotalAmount,JobNo,TransferCode)"
                      ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & txtbranchcode & "','" & Trim(ls_transcodeinv) & "','" & ls_transtype & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & txtDepartmentCode & "','" & TxtSiteID & "','" & txtbinID & "','" & RepApp(TxtRemarks) & "'," & Val(txttotalamount) & ",0," & Val(txttotalamount) & ",'ADJ','" & txttransno & "' )"
                      gc_dbcon.Execute ls_sql
                
                    With GrdGRN
                        For ln_cnt = 1 To .Rows - 1
                            If .TextMatrix(ln_cnt, 1) <> "" Then
                                ls_sql = "INSERT into IC_Trans(Compcode,BranchCode, TransCode ,Transtype, ItemCode, Quantity, ItemRate, Amount, TaxAmount,AvgRate,Remarks,BalQty,TransferCode)"
                                ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & txtbranchcode & "','" & Trim(ls_transcodeinv) & "','" & Trim(ls_transtype) & "','" & Trim(.TextMatrix(ln_cnt, 1)) & "'," & Val(0 & .TextMatrix(ln_cnt, 4)) & "," & Val(0 & .TextMatrix(ln_cnt, 5)) & "," & Val(0 & .TextMatrix(ln_cnt, 6)) & ",0," & Val(0 & .TextMatrix(ln_cnt, 8)) & " , '" & Trim(.TextMatrix(ln_cnt, 10)) & "' ," & Val(0 & .TextMatrix(ln_cnt, 7)) & ",'" & txttransno & "')"
                                gc_dbcon.Execute ls_sql
                            End If
                       Next
                 End With
                 
     End Select
gc_dbcon.CommitTrans

If Mode = "A" Then
    txttransno = maxtranscode
End If
InitializeGrid
Exit Sub
RollBack:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Private Sub SetVoucher()
Dim ls_sql As String
Dim ln_TmpVchNo
Dim ln_OrgVchNo
        ln_OrgVchNo = txtvchrno
        ' Save reference of Voucher
        ls_sql = "INSERT into Gl_Ref(compcode,BranchCode,Value_Date,Trans_Date, Voucher_No, VchrType, Vchr_Remarks,InstrumentNo,CrncyCode,ExchgRate,userid,adddate,addtime) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & Format(Date, "YYYY/MM/DD") & "','" & txtvchrno & "','" & txtvchrType & "','" & TxtRemarks & "','" & Txtinstrument & "','PKR'," & Val(0) & ",'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "')"
        gc_dbcon.Execute ls_sql
                 
        ' Save Details of Voucher
                
        'debit the party account
        If Val(txttotalamount) + Val(txttotaltaxamount) > 0 Then
            ls_sql = "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, ExchgRate,userid,adddate,addtime,Acct_Nirration,AcctName) "
            ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtpartyaccount & "'," & 1 & ",'" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & txtvchrno & "','" & txtvchrType & "'," & Val(txttotalamount) + Val(txttotaltaxamount) & "," & 0 & "," & Val(0) & ",'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & Trim(TxtRemarks) & "','" & Trim(TxtRemarks) & "')"
            gc_dbcon.Execute ls_sql
        End If
              
        'credit the sale account
        If Val(txttotalamount) > 0 Then
            ls_sql = "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, ExchgRate,userid,adddate,addtime,Acct_Nirration,AcctName) "
            ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtsaleaccount & "'," & 2 & ",'" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & txtvchrno & "','" & txtvchrType & "',0," & Val(txttotalamount) & "," & Val(0) & ",'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & Trim(TxtRemarks) & "','" & Trim(TxtRemarks) & "')"
            gc_dbcon.Execute ls_sql
        End If
        'credit the tax account
        If Val(txttotaltaxamount) > 0 Then
            ls_sql = "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, ExchgRate,userid,adddate,addtime,Acct_Nirration,AcctName) "
            ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txttaxaccount & "'," & 3 & ",'" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & txtvchrno & "','" & txtvchrType & "',0," & Val(txttotaltaxamount) & "," & Val(0) & ",'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & Trim(TxtRemarks) & "','" & Trim(TxtRemarks) & "')"
            gc_dbcon.Execute ls_sql
        End If
        
        
        'update voucher reference
        If Mode = "A" Then
                    PR_VchCntr.Requery
                    PR_VchCntr.Filter = "Branchcode = '" & Gs_BranchCode & "' And  VchrType = '" & txtvchrType & "'"
  
                    If PR_VchType.Fields("VchrFrequency") = "1" Then
                       ln_TmpVchNo = Val(0 & PR_VchCntr.Fields("VchrMonth" & Trim(str(Month(txtvaluedate.Value))))) + 1
                    Else
                       ln_TmpVchNo = Val(0 & PR_VchCntr.Fields("VchrCount")) + 1
                    End If
 
                    If ln_TmpVchNo > ln_OrgVchNo And ln_OrgVchNo = Val(txtvchrno) Then
                        'lb_Vstat = True
                        txtvchrno = DoPad(Trim(str(ln_TmpVchNo)), 10)
                        
                        If PR_VchType.Fields("VchrFrequency") = "1" Then
                          PR_VchCntr.Fields("VchrMonth" & Trim(str(Month(txtvaluedate.Value)))) = ln_TmpVchNo
                        Else
                          PR_VchCntr.Fields("VchrCount") = ln_TmpVchNo
                        End If
                        PR_VchCntr.Update
                    Else
                        If PR_VchType.Fields("VchrFrequency") = "1" Then
                          PR_VchCntr.Fields("VchrMonth" & Trim(str(Month(txtvaluedate.Value)))) = PR_VchCntr.Fields("VchrMonth" & Trim(str(Month(txtvaluedate.Value)))) + 1
                        Else
                          PR_VchCntr.Fields("VchrCount") = PR_VchCntr.Fields("VchrCount") + 1
                        End If
                        PR_VchCntr.Update
                    End If
                End If
                
                
End Sub

Public Sub ClearVal()
End Sub
Private Sub setprint()
End Sub
Private Sub Printinvoice()
On Error GoTo LocalErr

   With CrystalReport1
        .WindowTitle = Me.Caption
        '.Destination = crptToPrinter
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "purchaseInvoice.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Invoice'"
        .SelectionFormula = "{Ic_Trans.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & "  and {Ic_Trans.Transtype}= 'P'"
        .SelectionFormula = .SelectionFormula & "  and {Ic_Trans.transcode} = '" & Trim(ls_Invoiceno) & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
   End With
Exit Sub
LocalErr:
Call SetErr("Printer Not Ready", vbCritical)
End Sub

Private Sub SetVal()
     txtvaluedate = PR_ICTransfer("Transdate")
     txtbranchcode = PR_ICTransfer("BranchCode") & ""
     Call txtBranchCode_KeyDown(vbKeyReturn, vbKeyShift)
     txtDepartmentCode = PR_ICTransfer("AccountCode") & ""
     Call txtDepartmentCode_KeyDown(vbKeyReturn, vbKeyShift)
     TxtSiteID = PR_ICTransfer("SiteID") & ""
     Call txtSiteID_KeyDown(vbKeyReturn, vbKeyShift)
     txtbinID = PR_ICTransfer("BinID") & ""
     Call txtBinID_KeyDown(vbKeyReturn, vbKeyShift)
     txtSiteID1 = PR_ICTransfer("SiteID1") & ""
     Call txtSiteID1_KeyDown(vbKeyReturn, vbKeyShift)
     TxtBinID1 = PR_ICTransfer("BinID1") & ""
     Call txtBinID1_KeyDown(vbKeyReturn, vbKeyShift)
     TxtRemarks = PR_ICTransfer("Remarks")
End Sub
Public Function ChkInputs() As Boolean
    If Len(txttransno.Text) = txttransno.MaxLength And Len(txtDepartmentCode) = txtDepartmentCode.MaxLength Then
      If Mode = "A" Or Mode = "E" Then
            If CheckPOQTY Then
             ChkInputs = True
            Else
            lb_found = MySeek(Ls_ItemName, "itemcode", PR_IcItem)
            If lb_found Then
                Ls_ItemName = Trim(PR_IcItem("Description") & "")
            End If
            'Call MsgBox("Stock of """ & Ls_ItemName & """ = " & str(ln_qty) & Chr(13) & "Sale Of  """ & Ls_ItemName & """ = " & str(LN_EnterQty) & Chr(13) & "Difference = " & str(LN_EnterQty - ln_qty) & Chr(13) & "Stock not available for Transfer !!!", vbCritical)
            ChkInputs = True
            End If
            Else
            ChkInputs = True
       End If
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function
Private Function CheckPOQTY() As Boolean
Dim ls_sql As String
Dim ls_ItemCode As String
Dim ln_POQTY As Double
Dim ln_INQTY As Double
Dim ln_TotalQTY As Double

Dim Pr_dumyPOQty As New Recordset
    
    With GrdGRN
        For ln_cnt = 1 To .Rows - 1
            ls_ItemCode = .TextMatrix(ln_cnt, 1)
            
            'check po qty
            ls_sql = "SELECT sum(IC_Trans.Quantity) AS QTY"
            ls_sql = ls_sql & " FROM IC_TransMaster INNER JOIN IC_Trans ON IC_TransMaster.Compcode = IC_Trans.Compcode AND IC_TransMaster.TransCode = IC_Trans.TransCode"
            ls_sql = ls_sql & " where IC_Trans.ItemCode = '" & ls_ItemCode & "' and  IC_TransMaster.Transtype in('P','I','R') and IC_TransMaster.compcode = '" & Gs_compcode & "' "
            Pr_dumyPOQty.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly
        
            If Not Pr_dumyPOQty.EOF Then
            ln_POQTY = Val(0 & Pr_dumyPOQty("QTY"))
            End If
            Pr_dumyPOQty.Close
        
            'check invoice qty
            ls_sql = "SELECT sum(IC_Trans.Quantity) AS QTY"
            ls_sql = ls_sql & " FROM IC_TransMaster INNER JOIN IC_Trans ON IC_TransMaster.Compcode = IC_Trans.Compcode AND IC_TransMaster.TransCode = IC_Trans.TransCode"
            ls_sql = ls_sql & " where IC_Trans.ItemCode = '" & ls_ItemCode & "' and  IC_TransMaster.Transtype in('S','O','D','B') and IC_TransMaster.compcode = '" & Gs_compcode & "' "
            Pr_dumyPOQty.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly
        
            If Not Pr_dumyPOQty.EOF Then
                ln_INQTY = Val(0 & Pr_dumyPOQty("QTY"))
            End If
            Pr_dumyPOQty.Close
            
            ln_TotalQTY = ln_POQTY - (ln_INQTY + Val(0 & .TextMatrix(ln_cnt, 2)))
            ln_qty = ln_POQTY - ln_INQTY
            LN_EnterQty = Val(0 & .TextMatrix(ln_cnt, 2))
            If ln_TotalQTY < 0 Then
                CheckPOQTY = False
                Ls_ItemName = ls_ItemCode
                Exit Function
            End If
        Next
     
     CheckPOQTY = True
    
    End With

End Function

Public Sub FrmRefresh()
    Pr_ICParty.Requery
    PR_ICTransfer.Requery
    PR_IcItem.Requery
    PR_Branch.Requery
    PR_VchCntr.Requery
    PR_VchType.Requery
End Sub

Private Sub AddToGrid()
Dim ln_cnt As Integer
            If (Val(txtqty) > 0 And Val(txtunitprice) > 0) Then
                    If PS_RowClicked = "" Then
                        If PI_SrNo = 0 Then
                            PI_SrNo = 1
                        Else
                            PI_SrNo = PI_SrNo + 1
                         End If
                     End If
        
                    If txtitemcode.Text <> "" Then
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
                               If MySeek(txtitemcode.Text, "ItemFind", PR_IcItem) Then
                                                .TextMatrix(.Row, 1) = Trim(txtitemcode)
                                                .TextMatrix(.Row, 2) = Val(txtqty)
                                                .TextMatrix(.Row, 3) = Val(txtunitprice)
                                                .TextMatrix(.Row, 4) = Val(txtamount)
                                                .TextMatrix(.Row, 5) = Val(txttaxamount)
                                                .TextMatrix(.Row, 6) = Trim(txtserialNo)
                            Else
                                Call SetErr("Item Code Not Found.", vbCritical)
                                txtitemcode.SetFocus
                            End If
                                txtitemcode.Text = ""
                                txtqty = ""
                                txtunitprice = ""
                                txttaxamount = ""
                                txtamount = ""
                                txtitemdesc = ""
                                txtmtype = ""
                                txtserialNo = ""
                                PS_RowClicked = ""
                        End With
                    End If
                        TotalAmount
                        txtitemcode.SetFocus
                   
        Else
            Call SetErr("Enter Qty./Unit Price !!!", vbCritical)
            txtqty.SetFocus
       End If
      

End Sub
Private Sub TotalAmount()
    Dim ln_cnt As Integer
      txttotalamount = ""
    With GrdGRN
        For ln_cnt = 1 To .Rows - 1
            txttotalamount = Val(txttotalamount) + Val(.TextMatrix(ln_cnt, 6))
        Next
    End With
End Sub


Private Sub txtvaluedate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtDepartmentCode.SetFocus
    End If
End Sub

Public Sub SetFrmEnv(ls_mode As String)
    txtLocCode.Enabled = IIf(ls_mode <> "D", True, False)
    txtpartycode.Enabled = IIf(ls_mode <> "D", True, False)
    TxtRemarks.Enabled = IIf(ls_mode <> "D", True, False)
    Frame2.Enabled = IIf(ls_mode <> "D", True, False)
End Sub
Private Sub LoadGRNTrans()
'On Error GoTo LocalErr
Dim Pr_LoadTrans As New Recordset
InitializeGrid
Dim ls_sql As String

ls_sql = " SELECT IC_Transfer.ItemCode, IC_Item.Description, IC_Transfer.Quantity, IC_Transfer.ItemRate, IC_Transfer.Amount, IC_ItemUM.Description AS UOM,IC_Transfer.AvgRate,IC_Transfer.Remarks,IC_Transfer.BalQty"
ls_sql = ls_sql & " FROM IC_Transfer INNER JOIN   IC_Item ON IC_Transfer.Compcode = IC_Item.Compcode AND IC_Transfer.ItemCode = IC_Item.ItemCode INNER JOIN  IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode"
ls_sql = ls_sql & "  where IC_Transfer.Compcode = '" & Gs_compcode & "' and IC_Transfer.Transcode = '" & txttransno & "'"

Pr_LoadTrans.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not Pr_LoadTrans.EOF Then
        With GrdGRN
            Do While Not Pr_LoadTrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = Trim(Pr_LoadTrans("ItemCode") & "")
                .TextMatrix(.Row, 2) = Trim(Pr_LoadTrans("Description") & "")
                .TextMatrix(.Row, 3) = Trim(Pr_LoadTrans("UOM") & "")
                .TextMatrix(.Row, 4) = Pr_LoadTrans("Quantity")
                .TextMatrix(.Row, 5) = Val(0 & Pr_LoadTrans("Itemrate"))
                .TextMatrix(.Row, 6) = Val(0 & Pr_LoadTrans("amount"))
                .TextMatrix(.Row, 7) = Val(0 & Pr_LoadTrans("BalQty"))
                .TextMatrix(.Row, 8) = Val(0 & Pr_LoadTrans("AvgRate"))
                .TextMatrix(.Row, 9) = .TextMatrix(.Row, 4) * .TextMatrix(.Row, 8)
                .TextMatrix(.Row, 10) = Trim(Pr_LoadTrans("Remarks") & "")
                .Rows = .Rows + 1
                Pr_LoadTrans.MoveNext
                If Pr_LoadTrans.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
        TotalAmount
    Else
        Call SetErr("Transaction not found.!!!", vbCritical)
        
    End If
    Pr_LoadTrans.Close
Exit Sub
LocalErr:
Call MsgBox(Err.Description)

End Sub
Private Sub GrdGRN_EnterCell()
GrdGRN.CellBackColor = vbHighlight
End Sub

Private Sub GrdGRN_LeaveCell()
With GrdGRN
    .CellBackColor = vbWindowBackground
End With
End Sub

Private Sub GrdGRN_DblClick()
    GrdGRN.SelectionMode = flexSelectionFree
End Sub

Private Sub GrdGRN_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 And GrdGRN.Col = 1 Then ' F1 key pressed
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = Text1
    Set PO_DESC = Text2
    Gs_SQL = "Select ItemCode,   Description from IC_Item "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Items"
    MyLookupOLDB.Show 1
    GrdGRN.TextMatrix(GrdGRN.Row, 1) = Text1
    Call GrdGRN_KeyPress(13)

 ElseIf KeyCode = 46 Then 'Delete Key Pressed
   GrdGRN_KeyPress (KeyCode)
 End If

 
End Sub

Private Sub GrdGRN_KeyPress(KeyAscii As Integer)
'On Error GoTo ErrHandler
 Call GetKeysAdd(GrdGRN, KeyAscii)
Exit Sub
'ErrHandler:
'MsgBox ("An Error has Occured In The MSFlexgrid1_KeyPress() Procedure") & vbCr & "Report This Error To Latifjat@hotmail.com" & vbCr & "Error Details :-" & vbCr & "Error Number : " & Err.Number & vbCr & "Error Description : " & Err.Description, vbCritical, "FlexGrid Example"
End Sub
Public Sub GetKeysAdd(argFlexGrid As MSHFlexGrid, KeyAscii As Integer)
'This Procedure is used to display the pressed key into FlexGrid in Addition Mode
'so that when you press Enter Key in the last row then one row will be added.
'When you press the BackSpace Key in an empty Row then a Row will be Removed.
'On Error GoTo ErrHandler

If KeyAscii = 13 Then 'if Enter Key then...
  Opt = ""
  With argFlexGrid
        .SelectionMode = flexSelectionByRow
        Row = .RowSel
    If .Col = 1 Then
           If .TextMatrix(Row, 1) = "" Then
             Call MsgBox("Enter/Select Item Code!!!", vbCritical)
             Exit Sub
           End If
          .TextMatrix(Row, 1) = DoPad(.TextMatrix(Row, 1), 6)
          If SearchInGrid(GrdGRN, .TextMatrix(Row, 1)) Then
             Call MsgBox("Record Already Exist in Grid", vbCritical)
            .TextMatrix(Row, 1) = ""
             Exit Sub
          End If

          If PR_IcItem.State = 1 Then PR_IcItem.Close
          PR_IcItem.Open " Select * From Ic_Item Where compcode = '" & Gs_compcode & "' and  ItemCode='" & Trim(.TextMatrix(Row, 1)) & " '", gc_dbcon, adOpenStatic, adLockReadOnly
          
          If PR_IcItem.RecordCount <= 0 Then
              Call MsgBox(Gs_RecNFMsg, vbCritical)
             .TextMatrix(Row, 1) = ""
          Else
             .TextMatrix(Row, 0) = Row
             .TextMatrix(Row, 2) = Trim(PR_IcItem("Description") & "")
             .TextMatrix(Row, 7) = CheckBalQTY(.TextMatrix(Row, 1))
             .TextMatrix(Row, 8) = PR_IcItem("AvgRate")
             .TextMatrix(Row, 9) = .TextMatrix(Row, 8) * PR_IcItem("AvgRate")
              txtitemname = .TextMatrix(Row, 2)

             .Col = .Col + 3
              PR_UOM.Open "Select * From IC_ItemUM Where MCode='" & Trim(PR_IcItem("Mcode") & "") & " '", gc_dbcon, adOpenStatic, adLockReadOnly
              If PR_UOM.RecordCount > 0 Then
                .TextMatrix(Row, 3) = Trim(PR_UOM("Description") & "")
              End If
              PR_UOM.Close
          End If
         PR_IcItem.Close
       ElseIf .Col = 2 Then
       ElseIf .Col = 3 Then
       ElseIf .Col = 4 Then
           If .TextMatrix(Row, 4) = "" Then
             Call MsgBox("Enter Quantity!!!", vbCritical)
             Exit Sub
           End If

          If PR_IcItem.State = 1 Then PR_IcItem.Close
          PR_IcItem.Open " Select * From Ic_Item Where compcode = '" & Gs_compcode & "' and  ItemCode='" & Trim(.TextMatrix(Row, 1)) & " '", gc_dbcon, adOpenStatic, adLockReadOnly
          If Not PR_IcItem.EOF Then
           .TextMatrix(Row, 5) = PR_IcItem("PurchaseCost")
           .TextMatrix(Row, 6) = Val(.TextMatrix(Row, 4)) * .TextMatrix(Row, 5)
          End If
             
          PR_IcItem.Close
          Call TotalAmount
          If Row = .Rows - 1 Then
           .Rows = .Rows + 1
          End If
          .Col = 1
          .Row = .Row + 1
          .SetFocus
          Row = Row + 1
          
        If .RowSel > 9 Then
              .TopRow = .Rows - 1 'To Move the Scrollbar
        End If
            
   End If
   End With
 Exit Sub
End If
      
If KeyAscii = 8 Then  'If BackSpace Key then...
With argFlexGrid
   If Len(Trim(.Text)) <> 0 Then  'If current cell is not empty then...
      .Text = Left(.Text, (Len(.Text) - 1)) 'Removing a character from the right side of the FlexGrid cell's text
      Opt = ""
   End If
End With
End If

If KeyAscii = 46 Then  'If Delete Key then...
   With argFlexGrid
     If .Rows > 2 Then
       If (.Rows > .FixedRows + 1) Then
           .RemoveItem .Row
       Else
           .Rows = .FixedRows
        End If
      End If
          
   End With
   
 Exit Sub
 End If
  
  If KeyAscii <> 27 And KeyAscii <> 8 Then
    With GrdGRN
      If .Col = 1 Or .Col = 4 Or .Col = 10 Then
        .Text = .Text & Chr(KeyAscii) 'Reset Value in Cell and Append the pressed character to the right.
      End If
    End With
  End If
End Sub



VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReports 
   Caption         =   "Customized Reports"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5055
   Icon            =   "frmReports.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   0
      TabIndex        =   2
      Top             =   3450
      Width           =   5055
      Begin VB.CheckBox Check1 
         Height          =   285
         Left            =   3750
         TabIndex        =   15
         Top             =   480
         Width           =   225
      End
      Begin VB.TextBox txtGroupDesc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1410
         MaxLength       =   100
         TabIndex        =   14
         Top             =   420
         Width           =   2235
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   4650
         Picture         =   "frmReports.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   420
         Width           =   315
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   1455
         Left            =   60
         TabIndex        =   18
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2566
         _Version        =   393216
         Cols            =   5
         AllowBigSelection=   0   'False
         HighLight       =   0
         SelectionMode   =   1
      End
      Begin MSMask.MaskEdBox txtGroupCode 
         Height          =   315
         Left            =   90
         TabIndex        =   13
         Top             =   420
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTypeGroup 
         Height          =   315
         Left            =   4110
         TabIndex        =   16
         Top             =   420
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Sum"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3660
         TabIndex        =   30
         Top             =   180
         Width           =   315
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Group Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   330
         TabIndex        =   28
         Top             =   180
         Width           =   870
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Group Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1890
         TabIndex        =   27
         Top             =   180
         Width           =   1305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Type Group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4080
         TabIndex        =   26
         Top             =   180
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   3390
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      TabIndex        =   1
      Top             =   570
      Width           =   5055
      Begin VB.OptionButton Option1 
         Caption         =   "Credit - Debit"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2970
         TabIndex        =   10
         Top             =   1530
         Width           =   1395
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Debit - Credit"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   1590
         TabIndex        =   9
         Top             =   1530
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1605
         TabIndex        =   23
         Top             =   1080
         Width           =   3375
         Begin VB.OptionButton optPL 
            Caption         =   "Profit n Loss"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   6
            Top             =   120
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optOthers 
            Caption         =   "Others"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2610
            TabIndex        =   8
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton optBal 
            Caption         =   "Balance Sheet"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1200
            TabIndex        =   7
            Top             =   120
            Width           =   1395
         End
      End
      Begin VB.TextBox txtRptSub 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   12
         Top             =   2310
         Width           =   3375
      End
      Begin VB.TextBox txtRptHead 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1605
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1890
         Width           =   3375
      End
      Begin VB.CommandButton cmdLookup0 
         Height          =   315
         Left            =   2190
         Picture         =   "frmReports.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   300
         Width           =   315
      End
      Begin MSMask.MaskEdBox txtRptDesc 
         Height          =   315
         Left            =   1605
         TabIndex        =   5
         Top             =   720
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   35
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRptCode 
         Height          =   315
         Left            =   1590
         TabIndex        =   3
         Tag             =   "SKIP"
         Top             =   300
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Calculation Base :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   29
         Top             =   1500
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Report Type :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   555
         TabIndex        =   24
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Report Sub Header :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   22
         Top             =   2370
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Report Header :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   390
         TabIndex        =   21
         Top             =   1950
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Report Description :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   105
         TabIndex        =   20
         Top             =   720
         Width           =   1425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Report Code :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   540
         TabIndex        =   19
         Top             =   300
         Width           =   990
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1005
      ButtonWidth     =   1217
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
            Caption         =   "&Listing"
            Description     =   "Print Listing."
            Object.ToolTipText     =   "Print listing."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
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
               Picture         =   "frmReports.frx":05EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReports.frx":0A42
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReports.frx":0E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReports.frx":12EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReports.frx":173E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReports.frx":1B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReports.frx":22E6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pb_BlnkVchr As Boolean
Dim Mode As String
Dim ls_RowClicked As String
Dim ln_rowno As Integer
Dim Ln_First As Integer
Dim Ls_CBase As String

Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_GlRptRef As Recordset
Dim PR_GlRptDetl As Recordset
Dim PR_GlTypeGroup As Recordset

Private Sub cmdLookup0_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtRptCode
    Set PO_DESC = txtRptDesc
    
    GoTop PR_GlRptRef
    MyLookup.Caption = "Customized Reports."
    MyLookup.FillGrid PR_GlRptRef, "ReportCode", "RptDescrip", 5
    MyLookup.Show 1
    
    If Len(txtRptCode.Text) > 0 Then txtRptCode_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Command1_Click()
    Text1.MaxLength = 35
    Set PO_CODE = Nothing
    Set PO_DESC = Nothing
    Set PO_AnyForm = Nothing
    
    Set PO_AnyForm = Me
    Set PO_CODE = txtTypeGroup
    Set PO_DESC = Text1
    
    GoTop PR_GlTypeGroup
    MyLookup.Caption = "Format Specifiers."
    MyLookup.FillGrid PR_GlTypeGroup, "Accumtype", "Description", 5
    MyLookup.Show 1
    
    If Len(txtTypeGroup.Text) > 0 Then txtTypeGroup_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Form_Load()
  Ls_CBase = ""
  
  SetToolBar(1) = chkRights("GLRPTS0001")
  SetToolBar(2) = chkRights("GLRPTS0002")
  SetToolBar(3) = chkRights("GLRPTS0003")
  SetToolBar(4) = chkRights("GLRPTS0004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
    
'  Frame4.Enabled = False
  
  Set PR_GlRptRef = New Recordset
  Set PR_GlRptDetl = New Recordset
  Set PR_GlTypeGroup = New Recordset

  PR_GlRptRef.Open "Select * from GlRpts_Ref where CompCode ='" & Gs_compcode & "'", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_GlRptDetl.Open "Select * from GlGroupDetl where CompCode ='" & Gs_compcode & "'", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_GlTypeGroup.Open "Select * from Accumtypes", gc_dbcon, adOpenDynamic, adLockOptimistic, 1

  
  Pb_BlnkVchr = IIf(PR_GlRptRef.EOF, True, False)
  Ln_First = 0

Grid1.ColWidth(0) = 300
Grid1.ColWidth(1) = 1050
Grid1.ColWidth(2) = 2200
Grid1.ColWidth(3) = 550
Grid1.ColWidth(4) = 500
Grid1.Row = 0
Grid1.Col = 1
Grid1.Text = "Groups"
Grid1.Col = 2
Grid1.Text = "         Group Description"
Grid1.Col = 3
Grid1.Text = " Type"
Grid1.Col = 4
Grid1.Text = " Sum"

Grid1.Rows = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_GlRptRef.Close
End Sub

Private Sub grid1_DblClick()
'        Frame4.Enabled = True
        Grid1.Row = Grid1.RowSel
        Grid1.Col = 1
        txtGroupCode.Text = Grid1.Text
        Grid1.Col = 2
        txtGroupDesc.Text = Grid1.Text
        Grid1.Col = 3
        txtTypeGroup.Text = Grid1.Text
        Grid1.Col = 4
        Check1.Value = Grid1.Text
        
        ls_RowClicked = "Y"
        ln_rowno = Grid1.RowSel
End Sub
'    With Grid1
'        If .Row > 0 Then
'            PI_CurRow = .Row
'        End If
'
'        TxtAccountNo = .TextMatrix(.Row, 1)
'        txtAcctNarration = .TextMatrix(.Row, 2)
'        txtdrAmount = Val(.TextMatrix(.Row, 3))
'        txtCrAmount = Val(.TextMatrix(.Row, 4))
'
'        PS_RowClicked = "Y"
'    End With
'End Sub

Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
    With Grid1
        If KeyCode = vbKeyDelete Then
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            .RemoveItem .Row
        End If
    End With
End Sub

Private Sub optBal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
txtRptHead.SetFocus
End If

End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
      Ls_CBase = "D"
    ElseIf Index = 1 Then
      Ls_CBase = "C"
    End If
End Sub

Private Sub optOthers_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       Option1.Item(0).Enabled = True
       Option1.Item(1).Enabled = True
    End If
End Sub

Private Sub optPL_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
txtRptHead.SetFocus
End If
End Sub

Private Sub txtgroupcode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
If txtGroupCode.Text <> "" Then
txtGroupCode.Text = DoPad(txtGroupCode.Text, 10)
txtGroupDesc.SetFocus
End If
End If

End Sub

Private Sub txtGroupDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
txtTypeGroup.SetFocus
End If

End Sub

Private Sub txtRptCode_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lb_found As Boolean

If KeyCode = vbKeyReturn Then

         PR_GlRptRef.Requery
         txtRptCode.Text = DoPad(txtRptCode, 2)
         lb_found = MySeek(txtRptCode.Text, "reportcode", PR_GlRptRef)
         PR_GlRptDetl.Close
         PR_GlRptDetl.Open "Select * from GlGroupDetl where CompCode ='" & Gs_compcode & "'and reportcode ='" & txtRptCode.Text & "'order by groupcode " '", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
       
       txtRptCode.Text = UCase(txtRptCode.Text)
       Select Case Mode
            Case "A"
                If lb_found Then
                   MsgBox "Report Code Already already exist.", vbCritical, "E-Counts 2.0"
                   Cancel = True
                   txtRptCode.SetFocus
                Else
                   txtRptDesc.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   MsgBox "Record does not exist.", vbCritical, "E-Counts 2.0"
                   Cancel = True
'                   txtRptCode.SetFocus
                Else
                   Call SetVal
                   If Mode <> "D" Then
                      txtRptDesc.SetFocus
                   End If
                End If
            End Select
   ElseIf KeyCode = vbKeyF12 Then
    cmdLookup0_Click
   End If


  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
       cmdLookup0.Enabled = False
    Else
     cmdLookup0.Enabled = True
    End If
    
    If Pb_BlnkVchr And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_GlRptRef, frmReports, txtRptCode, txtRptDesc, "x", "CompCount", 3, "ReportCode", "ReportDescrip", 1, False, Toolbar1)
    End If
        
'        Frame4.Enabled = False

End Sub

Public Sub SaveValues()
Dim cntsql As New ADODB.Command
PB_BlnkComp = False

            Dim ls_Ttype As String
            If optPL.Value = True Then
            ls_Ttype = "P"
            ElseIf optBal.Value = True Then
            ls_Ttype = "B"
            Else
            ls_Ttype = "O"
            End If

cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText

     Select Case Mode
           Case "A"
              cntsql.CommandText = "INSERT into GlRpts_Ref(compcode,reportcode,reportbase,rptdescrip,rptheader,rptsubheader,CalcBase) VALUES ('" & Gs_compcode & "','" & txtRptCode.Text & "','" & ls_Ttype & "','" & txtRptDesc.Text & "','" & txtRptHead.Text & "','" & txtRptSub.Text & "','" & Ls_CBase & "')"
              cntsql.Execute
              Grid1.Row = 1
              For ln_cnt = 1 To Grid1.Rows - 2
                Grid1.Col = 1
                If Grid1.Text <> "" Then
                    txtGroupCode.Text = Grid1.Text
                    Grid1.Col = 2
                    txtGroupDesc.Text = Grid1.Text
                    Grid1.Col = 3
                    txtTypeGroup.Text = Grid1.Text
                    Grid1.Col = 4
                    Check1.Value = Grid1.Text
                    
                    cntsql.CommandText = "INSERT into GlGroupDetl(compcode,reportcode,groupcode,groupdesc,typegroup,SumGroup) VALUES ('" & Gs_compcode & "','" & txtRptCode.Text & "','" & txtGroupCode.Text & "','" & txtGroupDesc.Text & "','" & txtTypeGroup.Text & "'," & Check1.Value & ")"
                    cntsql.Execute
                End If
                Grid1.Row = Grid1.Row + 1
              Next
           
           Case "E"
            cntsql.CommandText = "DELETE FROM GlRpts_Ref WHERE compcode = '" & Gs_compcode & "'and reportcode='" & txtRptCode.Text & "'"
            cntsql.Execute
            cntsql.CommandText = "DELETE FROM Glgroupdetl WHERE compcode = '" & Gs_compcode & "'and reportcode='" & txtRptCode.Text & "'"
            cntsql.Execute
              
              cntsql.CommandText = "INSERT into GlRpts_Ref(compcode,reportcode,reportbase,rptdescrip,rptheader,rptsubheader,CalcBase) VALUES ('" & Gs_compcode & "','" & txtRptCode.Text & "','" & ls_Ttype & "','" & txtRptDesc.Text & "','" & txtRptHead.Text & "','" & txtRptSub.Text & "','" & Ls_CBase & "')"
              cntsql.Execute
              Grid1.Row = 1
              For ln_cnt = 1 To Grid1.Rows - 2
                Grid1.Col = 1
                If Grid1.Text <> "" Then
                    txtGroupCode.Text = Grid1.Text
                    Grid1.Col = 2
                    txtGroupDesc.Text = Grid1.Text
                    Grid1.Col = 3
                    txtTypeGroup.Text = Grid1.Text
                    Grid1.Col = 4
                    Check1.Value = Grid1.Text
                    
                    cntsql.CommandText = "INSERT into GlGroupDetl(compcode,reportcode,groupcode,groupdesc,typegroup,sumGroup) VALUES ('" & Gs_compcode & "','" & txtRptCode.Text & "','" & txtGroupCode.Text & "','" & txtGroupDesc.Text & "','" & txtTypeGroup.Text & "'," & Check1.Value & ")"
                    cntsql.Execute
                End If
                Grid1.Row = Grid1.Row + 1
              Next
              
           Case "D"
            cntsql.CommandText = "DELETE FROM GlRpts_Ref WHERE compcode = '" & Gs_compcode & "'and reportcode='" & txtRptCode.Text & "'"
            cntsql.Execute
            cntsql.CommandText = "DELETE FROM Glgroupdetl WHERE compcode = '" & Gs_compcode & "'and reportcode='" & txtRptCode.Text & "'"
            cntsql.Execute

     End Select
PR_GlRptRef.Requery
PR_GlRptDetl.Requery


Option1.Item(0).Enabled = False
Option1.Item(1).Enabled = False
Check1.Value = 0
End Sub
Public Sub ClearVal()
          
     txtRptCode.Text = ""
     txtRptDesc.Text = ""
     txtRptHead.Text = ""
     txtRptSub.Text = ""
     txtGroupCode = ""
     txtGroupDesc = ""
     txtTypeGroup = ""
     Check1.Value = 0
     Option1.Item(0).Enabled = False
     Option1.Item(1).Enabled = False
     Option1.Item(0).Value = False
     Option1.Item(1).Value = False
     

Ln_First = 0
Grid1.Clear
Grid1.Row = 0
Grid1.Col = 1
Grid1.Text = "Groups"
Grid1.Col = 2
Grid1.Text = "         Group Description"
Grid1.Col = 3
Grid1.Text = " Type"
Grid1.Col = 4
Grid1.Text = " Sum"

Grid1.Rows = 2
ln_rowno = 0
ls_RowClicked = ""
     
End Sub

Private Sub SetVal()
If Not ((PR_GlRptRef.EOF = True) And (PR_GlRptRef.BOF = True)) Then
     txtRptCode = PR_GlRptRef("reportcode")
     txtRptDesc = PR_GlRptRef("rptDescrip")
     
     If PR_GlRptRef("reportbase") = "P" Then
       optPL.Value = True
     ElseIf PR_GlRptRef("reportbase") = "B" Then
       optBal.Value = True
     Else
       optOthers.Value = True
     End If

     txtRptHead = PR_GlRptRef("rptheader")
     txtRptSub = PR_GlRptRef("rptsubheader")
     
     If optOthers.Value = True Then
        Option1.Item(0).Enabled = True
        Option1.Item(1).Enabled = True
        Option1.Item(0).Value = IIf(PR_GlRptRef("CalcBase") = "D", True, False)
        Option1.Item(1).Value = IIf(PR_GlRptRef("CalcBase") = "C", True, False)
     End If
     
     Grid1.Clear
     Grid1.Rows = 2
     Grid1.Row = 1
     
     Dim lb_founddetl As Boolean
     
     lb_founddetl = MySeek(txtRptCode.Text, "reportcode", PR_GlRptDetl)

     If lb_founddetl Then
     
     PR_GlRptDetl.Requery
     PR_GlRptDetl.MoveFirst
     While Not PR_GlRptDetl.EOF
        Grid1.Col = 1
        Grid1.Text = PR_GlRptDetl("groupcode")
        Grid1.Col = 2
        Grid1.Text = PR_GlRptDetl("groupdesc") & ""
        Grid1.Col = 3
        Grid1.Text = PR_GlRptDetl("typegroup") & ""
        Grid1.Col = 4
        Grid1.Text = IIf(IsNull(PR_GlRptDetl("Sumgroup")), 0, PR_GlRptDetl("Sumgroup"))
        
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Row = Grid1.Row + 1
        
        PR_GlRptDetl.MoveNext
     Wend

    End If

End If
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtRptCode.Text) = txtRptCode.MaxLength And Len(txtRptHead) > 0 Then
       ChkInputs = True
    Else
       Call SetErr("Incomplete Data found", vbCritical)
       ChkInputs = False
    End If
End Function

Private Sub txtcompname_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
    txtcompaddr1.SetFocus
 End If
End Sub


Private Sub txtRptDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
optPL.SetFocus
End If
End Sub

Private Sub txtRptHead_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
txtRptSub.SetFocus
End If

End Sub

Private Sub txtRptSub_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
txtGroupCode.SetFocus
End If

End Sub

Private Sub txtTypeGroup_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
Dim lb_foundtype As Boolean

If KeyCode = vbKeyReturn Then

PR_GlTypeGroup.Close
PR_GlTypeGroup.Open "Select * from Accumtypes", gc_dbcon, adOpenStatic, adLockReadOnly, 1

lb_foundtype = MySeek(txtTypeGroup.Text, "AccumType", PR_GlTypeGroup)

If lb_foundtype Then

txtTypeGroup.Text = UCase(txtTypeGroup.Text)

If txtTypeGroup.Text <> "" And (Mode <> "") Then

If ls_RowClicked = "Y" Then
    Grid1.Row = ln_rowno
    Grid1.Col = 1
    Grid1.Text = txtGroupCode.Text
    Grid1.Col = 2
    Grid1.Text = txtGroupDesc.Text
    Grid1.Col = 3
    Grid1.Text = txtTypeGroup.Text
    Grid1.Col = 4
    Grid1.Text = Check1.Value
    
    txtGroupCode = ""
    txtGroupDesc = ""
    txtTypeGroup = ""
    Check1.Value = 0
    ls_RowClicked = ""
    ln_rowno = 0
Else
rowselcount = Grid1.RowSel
Grid1.Row = 1
        For ln_cnt = 1 To Grid1.Rows - 2
        Grid1.Col = 1
        If Grid1.Text = txtGroupCode.Text Then
        lb_found = True
        Exit For
        End If
        Grid1.Row = Grid1.Row + 1
        Next
        
                If (Not lb_found) And (rowselcount < 1 Or rowselcount > (Grid1.Rows - 2)) Then
                    Grid1.Row = Grid1.Rows - 1
                    Grid1.Col = 1
                    Grid1.Text = txtGroupCode.Text
                    Grid1.Col = 2
                    Grid1.Text = txtGroupDesc.Text
                    Grid1.Col = 3
                    Grid1.Text = txtTypeGroup.Text
                    Grid1.Col = 4
                    Grid1.Text = Check1.Value
                    
                    txtGroupCode = ""
                    txtGroupDesc = ""
                    txtTypeGroup = ""
                    Check1.Value = 0
                    Grid1.Rows = Grid1.Rows + 1
                    txtGroupCode.SetFocus
                    Ln_First = 1
                Else
                    If Mode <> "E" Then
                    Call SetErr("Group Code already exist.", vbCritical)
                    Cancel = True
                    txtGroupCode.SetFocus
                    Else
                    Grid1.Row = rowselcount
                    Grid1.Col = 1
                    Grid1.Text = txtGroupCode.Text
                    Grid1.Col = 2
                    Grid1.Text = txtGroupDesc.Text
                    Grid1.Col = 3
                    Grid1.Text = txtTypeGroup.Text
                    Grid1.Col = 4
                    Grid1.Text = Check1.Value

                    txtGroupCode = ""
                    txtGroupDesc = ""
                    txtTypeGroup = ""
                    Check1.Value = 0
                    End If
                End If

End If
Else
Call SetErr("Type Group does not exist.", vbCritical)
txtTypeGroup.Text = ""
End If
End If
ElseIf KeyCode = vbKeyF12 Then
      Command1_Click
End If

rowselcount = 0
Grid1.RowSel = Grid1.Rows - 1
End Sub


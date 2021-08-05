VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmUserRight 
   Caption         =   "User Maintenance"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5115
   Icon            =   "frmUserRight.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5070
      MaxLength       =   50
      TabIndex        =   12
      Tag             =   "SKIP"
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
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
      Height          =   4710
      Left            =   0
      TabIndex        =   5
      Top             =   570
      Width           =   5115
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2985
         MaxLength       =   64
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   540
         Width           =   1980
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2985
         MaxLength       =   64
         TabIndex        =   23
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   180
         Width           =   1980
      End
      Begin VB.TextBox txtPassward 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1245
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1260
         Width           =   3720
      End
      Begin VB.TextBox txtConfirm 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1245
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1620
         Width           =   3735
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   2640
         Picture         =   "frmUserRight.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   180
         Width           =   315
      End
      Begin VB.CommandButton cmdLookup0 
         Height          =   315
         Left            =   2640
         Picture         =   "frmUserRight.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   540
         Width           =   315
      End
      Begin MSMask.MaskEdBox txtUserId 
         Height          =   315
         Left            =   1245
         TabIndex        =   2
         Tag             =   "SKIPN"
         Top             =   540
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
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
      Begin MSMask.MaskEdBox txtCompCode 
         Height          =   315
         Left            =   1245
         TabIndex        =   0
         Tag             =   "SKIPN"
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         MaxLength       =   3
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
      Begin MSMask.MaskEdBox txtusername 
         Height          =   315
         Left            =   1245
         TabIndex        =   16
         Tag             =   "SKIP"
         Top             =   900
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   50
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
      Begin VB.Frame Frame1 
         Height          =   2730
         Left            =   60
         TabIndex        =   13
         Top             =   1935
         Width           =   5025
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   2265
            Picture         =   "frmUserRight.frx":05EE
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   525
            Width           =   315
         End
         Begin VB.CommandButton Command3 
            Height          =   315
            Left            =   2040
            Picture         =   "frmUserRight.frx":0760
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   180
            Width           =   315
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   2370
            MaxLength       =   64
            TabIndex        =   18
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   180
            Width           =   2535
         End
         Begin VB.TextBox txtDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   2595
            MaxLength       =   64
            TabIndex        =   14
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   525
            Width           =   2310
         End
         Begin MSFlexGridLib.MSFlexGrid Grid1 
            Height          =   1770
            Left            =   60
            TabIndex        =   15
            Top             =   870
            Width           =   4890
            _ExtentX        =   8625
            _ExtentY        =   3122
            _Version        =   393216
            Cols            =   3
            AllowBigSelection=   0   'False
            HighLight       =   0
            SelectionMode   =   1
            AllowUserResizing=   3
         End
         Begin MSMask.MaskEdBox txtmodule 
            Height          =   315
            Left            =   1170
            TabIndex        =   25
            Tag             =   "SKIP"
            Top             =   180
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            PromptInclude   =   0   'False
            MaxLength       =   3
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
         Begin MSMask.MaskEdBox txtprocess 
            Height          =   315
            Left            =   1170
            TabIndex        =   26
            Tag             =   "SKIP"
            Top             =   525
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            PromptInclude   =   0   'False
            MaxLength       =   3
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Process Id :"
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
            Left            =   225
            TabIndex        =   21
            Top             =   555
            Width           =   855
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Sys. Module :"
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
            Top             =   210
            Width           =   975
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "User Name :"
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
         Left            =   315
         TabIndex        =   17
         Top             =   930
         Width           =   885
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Confirm :"
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
         TabIndex        =   11
         Top             =   1656
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Company Code:"
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
         TabIndex        =   9
         Top             =   210
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "User Id :"
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
         Left            =   600
         TabIndex        =   8
         Top             =   570
         Width           =   600
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Password :"
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
         Left            =   360
         TabIndex        =   7
         Top             =   1302
         Width           =   840
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   1058
      ButtonWidth     =   1402
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
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
               Picture         =   "frmUserRight.frx":08D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUserRight.frx":0D26
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUserRight.frx":117A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUserRight.frx":15CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUserRight.frx":1A22
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUserRight.frx":1E76
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUserRight.frx":25CA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmUserRight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pb_BlnkVchr As Boolean
Dim Mode As String

Public PO_CODE As Object
Public PO_DESC As Object

Dim PI_CurRow    As Integer
Dim PI_SrNo      As Integer
Dim PS_RowClicked As String

Dim PR_Module As New Recordset
Dim PR_SyUser As Recordset
Dim PR_SyProc As Recordset
Dim PR_SyRights As Recordset
Dim PR_SyComp As Recordset
Dim LI_CurRow As String
Dim LS_status As String
Dim ln_cntml   As Integer

Private Sub AddToGrid()
Dim ln_cnt As Integer
        
        If PS_RowClicked = "" Then
            If PI_SrNo = 0 Then
                PI_SrNo = 1
            Else
                PI_SrNo = PI_SrNo + 1
            End If
        End If
        
            With Grid1
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
                .TextMatrix(.Row, 1) = txtprocess
                .TextMatrix(.Row, 2) = txtdesc
                .TextMatrix(.Row, 3) = "ON"
                 PS_RowClicked = ""
              End With
End Sub

Private Sub cmdLookup0_Click()
    Set PO_CODE = Nothing
    Set PO_DESC = Nothing
    Set PO_AnyForm = Nothing
    
    Set PO_AnyForm = Me
    Set PO_CODE = txtuserid
    Set PO_DESC = Text4
    
    GoTop PR_SyUser
    MyLookup.Caption = "Users"
    MyLookup.FillGrid PR_SyUser, "UserID", "UserName", 5
    MyLookup.Show 1
    
    If Len(txtuserid.Text) > 0 Then txtuserid_KeyDown vbKeyReturn, vbKeyShift
End Sub
Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtprocess
    Set PO_DESC = txtdesc
    If UCase(Trim(txtmodule)) <> "CM" Then
        GoTop PR_SyProc
        MyLookup.Caption = "Processes"
        MyLookup.FillGrid PR_SyProc, "procCode", "procdesc", 5
    Else
        GoTop PR_SyComp
        MyLookup.Caption = "Companies"
        MyLookup.FillGrid PR_SyComp, "CompCode", "CompName", 5
    End If
    MyLookup.Show 1
    txtprocess.SetFocus
    If txtprocess <> "" Then txtprocess_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command2_Click()
   
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCompCode
    Set PO_DESC = Text3
    GoTop PR_SyComp
    MyLookup.Caption = "Companies"
    MyLookup.FillGrid PR_SyComp, "Compcode", "CompName", 5
    MyLookup.Show 1
    
    If Val(txtCompCode) > 0 Then txtcompcode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command3_Click()
    Set PO_CODE = Nothing
    Set PO_DESC = Nothing
    Set PO_AnyForm = Nothing
    
    Set PO_AnyForm = Me
    Set PO_CODE = txtmodule
    Set PO_DESC = Text2
    GoTop PR_Module
    MyLookup.Caption = "Modules"
    MyLookup.FillGrid PR_Module, "IdCode", "IdDescrip", 5
    MyLookup.Show 1
    If Len(txtmodule.Text) > 0 Then txtmodule_KeyDown vbKeyReturn, vbKeyShift
End Sub



Private Sub Form_Load()

  SetToolBar(1) = chkRights("SMUSERMT01")
  SetToolBar(2) = chkRights("SMUSERMT02")
  SetToolBar(3) = chkRights("SMUSERMT03")
  SetToolBar(4) = chkRights("SMUSERMT04")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  Set PR_SyUser = New Recordset
  Set PR_SyProc = New Recordset
  Set PR_SyRights = New Recordset
  Set PR_SyComp = New Recordset
  PR_SyUser.Open "Select * from syusers", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_SyProc.Open "Select * from syproc", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_SyRights.Open "Select *,Compcode+UserId as FindField from syrights Order by compcode,userid", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_SyComp.Open "Select * from syscomp", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_Module.Open "Select * from Fcm_Ids where Recid = 'MOD'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  Pb_BlnkVchr = IIf(PR_SyUser.EOF, True, False)
  InitializeGrid
  ln_cntml = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  PR_SyProc.Close
  PR_SyRights.Close
  PR_SyComp.Close
  PR_SyUser.Close
  PR_Module.Close
End Sub

Private Sub grid1_DblClick()
    With Grid1
        If .Row > 0 Then
            PI_CurRow = .Row
        End If
         txtprocess = .TextMatrix(.Row, 1)
         txtdesc = .TextMatrix(.Row, 2)
         PS_RowClicked = "Y"
    End With
End Sub
Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
With Grid1
    If KeyCode = vbKeyDelete Then
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
                .RemoveItem .Row
                If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                    .TextMatrix(.Row, 0) = ""
                    PI_SrNo = 0
                End If
            recnt  'Re counter in grid
    ElseIf KeyCode = vbKeyReturn Then
        Call grid1_DblClick
    End If
End With
End Sub

Private Sub txtConfirm_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If txtPassward.Text = txtConfirm.Text Then
            txtmodule.SetFocus
    Else
      Call SetErr("Passward don't match", vbCritical)
      txtPassward.SetFocus
    End If
End If
End Sub

Private Sub txtmodule_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtmodule <> "" Then
      txtmodule = UCase(txtmodule)
      PR_SyProc.Filter = "proctype ='" & txtmodule & "'"
      If MySeek(LTrim(RTrim(txtmodule.Text)), "Idcode", PR_Module) Then
        Text2 = PR_Module("IdDescrip")
        If Mode = "A" Then
        If UCase(txtmodule) <> "CM" Then
             LoadModProc
        Else
             LoadCMPProc
        End If
            txtprocess.SetFocus
        Else
           txtprocess.SetFocus
        End If
      Else
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtmodule.SetFocus
      End If
ElseIf KeyCode = vbKeyPageUp Then
    txtConfirm.SetFocus
End If
End Sub

Private Sub txtPassward_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtConfirm.SetFocus
End Sub
Private Sub txtprocess_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
Dim lb_found1 As Boolean
Dim ln_cnt As Integer

If KeyCode = vbKeyReturn Then
    If txtprocess.Text <> "" Then
        If UCase(Trim(txtmodule)) <> "CM" Then
            lb_found1 = MySeek(Trim(txtprocess.Text), "procCode", PR_SyProc)
        Else
            txtprocess = DoPad(txtprocess, 3)
            lb_found1 = MySeek(Trim(txtprocess.Text), "CompCode", PR_SyComp)
        End If

        If Not lb_found1 Then
                Call SetErr(Gs_RecNFMsg, vbCritical)
        Else
                txtdesc.Text = IIf(UCase(Trim(txtmodule)) <> "CM", PR_SyProc("procdesc"), PR_SyComp("Compname"))
                txtprocess.SetFocus
                    For ln_cnt = 1 To Grid1.Rows - 1
                    If Grid1.TextMatrix(ln_cnt, 1) = txtprocess.Text Then
                    lb_found = True
                      Exit For
                    End If
                    Next

                If lb_found Then
                    Call SetErr("Process already exist.", vbCritical)
                    txtprocess.SetFocus
                Else
                  AddToGrid
                  txtprocess = ""
                  txtdesc = ""
                  txtprocess.SetFocus
                End If
        End If
    End If
End If
End Sub
Private Sub txtcompcode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

If KeyCode = vbKeyReturn Then
    If txtCompCode.Text <> "" Then
        txtCompCode.Text = DoPad(txtCompCode.Text, 3)
        lb_found = MySeek(txtCompCode, "compcode", PR_SyComp)
        
        If lb_found Then
            Text3 = PR_SyComp("CompName")
            txtuserid.SetFocus
        Else
            Call SetErr(Gs_RecNFMsg, vbCritical)
        End If
    End If
End If

End Sub

Private Sub txtuserid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

If Lastkey(KeyCode) And txtuserid <> "" Then

       lb_found = MySeek(txtuserid.Text, "userid", PR_SyUser)
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   SetClear Me
                   txtuserid.SetFocus
                Else
                      txtUsername.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   SetClear Me
                   txtuserid.SetFocus
                Else
                   Call SetVal
                    Text4 = Trim(PR_SyUser("UserName") & "")
                   LoadGRNTrans
                   If Mode <> "D" Then txtPassward.SetFocus
                End If
            End Select
    PR_SyProc.Filter = adFilterNone
End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
       cmdLookup0.Enabled = False
    ElseIf Range(Button.Index, 2, 3) Then
       cmdLookup0.Enabled = True
    End If
    If Button.Index = 1 Or Button.Index = 7 Then InitializeGrid
    If Pb_BlnkVchr And Range(Button.Index, 2, 3) Then
       Call SetErr("Data not found.", vbCritical)
       Mode = ""
    Else
       Mode = DentMode(Mode, Button.Index, PR_SyUser, frmUserRight, txtCompCode, txtPassward, "x", "CompCount", 3, "userid", "passward", 1, False, Toolbar1)
    End If
End Sub
Public Sub SaveValues()
Dim cntsql As New ADODB.Command
Dim ln_cnt As Integer
Dim Ls_Password As String
Pb_BlnkVchr = False

cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText

     Ls_Password = EnCode(txtPassward.Text, 50)
     Select Case Mode
           Case "D"
                cntsql.CommandText = "DELETE FROM syusers WHERE compcode = '" & txtCompCode & "'and userid='" & txtuserid.Text & "'"
                cntsql.Execute
                cntsql.CommandText = "DELETE FROM syrights WHERE compcode = '" & txtCompCode & "'and userid='" & txtuserid.Text & "'"
                cntsql.Execute
                
           Case Else
             If Mode = "E" Then
                cntsql.CommandText = "UPDATE syusers SET password ='" & Ls_Password & "',UserName ='" & txtUsername & "' WHERE  compcode = '" & txtCompCode & "' and userid= '" & Trim(txtuserid.Text) & "'"
                cntsql.Execute
                
                cntsql.CommandText = "DELETE FROM syrights WHERE compcode = '" & txtCompCode & "'and rtrim(userid)='" & txtuserid.Text & "'"
                cntsql.Execute
             Else
                cntsql.CommandText = "INSERT into SyUsers(compcode,userid,UserName,password) VALUES ('" & txtCompCode & "','" & txtuserid.Text & "','" & txtUsername.Text & "','" & Ls_Password & "')"
                cntsql.Execute
             End If
              
              With Grid1
              For ln_cnt = 1 To .Rows - 1
                  cntsql.CommandText = "INSERT into syrights(compcode,userid,processid,rights) VALUES ('" & txtCompCode & "','" & txtuserid.Text & "','" & .TextMatrix(ln_cnt, 1) & "','" & IIf(.TextMatrix(ln_cnt, 3) = "ON", "1", "0") & "')"
                  cntsql.Execute
              Next
              End With
     End Select
FrmRefresh
ln_cnt = 1
PI_SrNo = 0
PS_RowClicked = ""
InitializeGrid
End Sub

Private Sub SetVal()
 txtPassward = LTrim(RTrim(DeCode(PR_SyUser.Fields("Password"), 50))) & ""
 txtConfirm = txtPassward & ""
 txtUsername = Trim(PR_SyUser("UserName"))
End Sub
Public Function ChkInputs() As Boolean
  If Len(txtuserid.Text) > 0 And Len(txtPassward.Text) > 0 Then
     ChkInputs = True
  Else
     Call SetErr("Incomplete Data found", vbCritical)
     ChkInputs = False
  End If
End Function
Public Sub InitializeGrid()
    With Grid1
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Process Id.|<Description  |<Status|<Module "
        .ColWidth(1) = 1200
        .ColWidth(2) = 2500
        .ColWidth(3) = 1000
        .ColWidth(4) = 0
        .Redraw = True
    End With
End Sub

Private Sub LoadGRNTrans()
Dim lb_found As Boolean
Dim ln_cnt   As Integer
Dim temp As String
ln_cnt = 1
InitializeGrid
    
    lb_found = MySeek(txtCompCode + txtuserid, "findfield", PR_SyRights)
    
    If lb_found Then
        With Grid1
            Do While UCase(LTrim(RTrim(PR_SyRights("FindField").Value))) = txtCompCode + UCase(txtuserid)
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = ln_cnt
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = Trim(PR_SyRights("ProcessId"))
                 If Len(Trim(PR_SyRights("ProcessId"))) <> 3 Then
                    .TextMatrix(.Row, 2) = IIf(MySeek(PR_SyRights("ProcessId"), "ProcCode", PR_SyProc), PR_SyProc("ProcDesc"), "")
                 Else
                   .TextMatrix(.Row, 2) = IIf(MySeek(PR_SyRights("ProcessId"), "CompCode", PR_SyComp), PR_SyComp("CompName"), "")
                 End If
                .TextMatrix(.Row, 3) = IIf(LTrim(RTrim(PR_SyRights("Rights").Value)) = "1", "ON", "OFF")
                .Rows = .Rows + 1
                ln_cnt = ln_cnt + 1
                PR_SyRights.MoveNext
                If PR_SyRights.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" And .Rows > 2 Then .RemoveItem .Rows - 1
            txtuserid.SetFocus
        End With
    Else
        Call SetErr("Transactions not found.", vbCritical)
        txtuserid.SetFocus
    End If
End Sub
Private Sub LoadModProc()
Dim lb_found As Boolean
Dim ln_cnt As Integer
Dim temp As String
If Grid1.Row < 1 Then
InitializeGrid
End If
    With Grid1
        For ln_cnt = 1 To .Rows - 1
            If .TextMatrix(ln_cnt, 4) = txtmodule Then
               lb_found = True
               Exit For
            End If
        Next
         If lb_found Then
           Call SetErr("Process already exist.", vbCritical)
           Exit Sub
         End If
     If .TextMatrix(.Row, 1) <> "" Then .Rows = .Rows + 1
    End With
        
    PR_SyProc.Filter = "ProcType =  '" & txtmodule & "'"
    If Not PR_SyProc.EOF Then
        With Grid1
            Do While Not PR_SyProc.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = ln_cntml
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = Trim(PR_SyProc("ProcCode"))
                .TextMatrix(.Row, 2) = Trim(PR_SyProc("Procdesc"))
                .TextMatrix(.Row, 3) = "ON"
                .TextMatrix(.Row, 4) = txtmodule
                .Rows = .Rows + 1
                ln_cntml = ln_cntml + 1
                PR_SyProc.MoveNext
            Loop
            ln_cntml = ln_cntml + 1
            'If .TextMatrix(.Rows - 1, 1) = "" And .Rows > 2 Then .RemoveItem .Rows - 1
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
            txtuserid.SetFocus
            
        End With
        
    Else
        Call SetErr("Transactions not found.", vbCritical)
        txtuserid.SetFocus
    End If
    recnt
    'PR_SyProc.Filter = adFilterNone
End Sub
Private Sub LoadCMPProc()
Dim lb_found As Boolean
Dim ln_cnt As Integer
Dim temp As String
If Grid1.Row < 1 Then
InitializeGrid
End If
    With Grid1
        For ln_cnt = 1 To .Rows - 1
            If .TextMatrix(ln_cnt, 4) = txtmodule Then
               lb_found = True
               Exit For
            End If
        Next
         If lb_found Then
           Call SetErr("Process already exist.", vbCritical)
           Exit Sub
         End If
     If .TextMatrix(.Row, 1) <> "" Then .Rows = .Rows + 1
    End With
        
    PR_SyComp.MoveFirst
    If Not PR_SyComp.EOF Then
        With Grid1
            Do While Not PR_SyComp.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = ln_cntml
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = Trim(PR_SyComp("CompCode"))
                .TextMatrix(.Row, 2) = Trim(PR_SyComp("CompName"))
                .TextMatrix(.Row, 3) = "ON"
                .TextMatrix(.Row, 4) = txtmodule
                .Rows = .Rows + 1
                ln_cntml = ln_cntml + 1
                PR_SyComp.MoveNext
            Loop
            ln_cntml = ln_cntml + 1
            'If .TextMatrix(.Rows - 1, 1) = "" And .Rows > 2 Then .RemoveItem .Rows - 1
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
            txtuserid.SetFocus
            
        End With
        
    Else
        Call SetErr("Transactions not found.", vbCritical)
        txtuserid.SetFocus
    End If
    recnt
    'PR_SyProc.Filter = adFilterNone
End Sub





Private Sub recnt()
Dim ln_cnt As Integer
With Grid1
    For ln_cnt = 1 To .Rows - 1
    .TextMatrix(ln_cnt, 0) = ln_cnt
    Next
End With
End Sub


Public Sub FrmRefresh()
  PR_SyUser.Requery
  PR_SyProc.Requery
  PR_SyRights.Requery
  PR_SyComp.Requery
End Sub

Private Sub txtusername_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtUsername <> "" Then txtPassward.SetFocus
If KeyCode = vbKeyPageUp Then txtuserid.SetFocus
End Sub

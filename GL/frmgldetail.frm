VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmgldetail 
   Caption         =   "Detail A/c"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmgldetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   6630
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   30
      TabIndex        =   1
      Top             =   570
      Width           =   6540
      Begin VB.TextBox txtsub0 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   1455
         MaxLength       =   9
         TabIndex        =   21
         Tag             =   "SKIP"
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton cmdLookup 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2535
         Picture         =   "frmgldetail.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   240
         Width           =   315
      End
      Begin VB.TextBox txtSubDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   2865
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   240
         Width           =   3585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sub Ledger A/C :"
         Height          =   210
         Left            =   150
         TabIndex        =   3
         Top             =   270
         Width           =   1245
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6630
      _ExtentX        =   11695
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
            Caption         =   "&Listing"
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
               Picture         =   "frmgldetail.frx":047C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmgldetail.frx":08D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmgldetail.frx":0D24
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmgldetail.frx":1178
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmgldetail.frx":15CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmgldetail.frx":1A20
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmgldetail.frx":2174
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5040
      Left            =   30
      TabIndex        =   16
      Top             =   1215
      Width           =   6555
      Begin VB.TextBox txtcustomCode 
         Height          =   315
         Left            =   4770
         TabIndex        =   54
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Frame Frame3 
         Caption         =   "Account Base"
         ForeColor       =   &H00000080&
         Height          =   525
         Left            =   30
         TabIndex        =   47
         Top             =   1665
         Width           =   6495
         Begin VB.OptionButton Option4 
            Caption         =   "&Both"
            Height          =   255
            Left            =   3165
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   180
            Width           =   870
         End
         Begin VB.OptionButton Option3 
            Caption         =   "&Shadowed"
            Height          =   255
            Left            =   4215
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   195
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "&Balance Sheet"
            Height          =   255
            Left            =   1620
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   180
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "&Profit n Loss"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   180
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Profit n Loss Routing Note #."
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   1380
         Left            =   30
         TabIndex        =   35
         Top             =   3615
         Width           =   6495
         Begin VB.TextBox txtplndesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2475
            Locked          =   -1  'True
            MaxLength       =   64
            TabIndex        =   51
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   585
            Width           =   3975
         End
         Begin VB.CommandButton Command5 
            Height          =   315
            Left            =   2115
            Picture         =   "frmgldetail.frx":25C8
            Style           =   1  'Graphical
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   960
            Width           =   315
         End
         Begin VB.CommandButton Command4 
            Height          =   315
            Left            =   2115
            Picture         =   "frmgldetail.frx":273A
            Style           =   1  'Graphical
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   585
            Width           =   315
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   2115
            Picture         =   "frmgldetail.frx":28AC
            Style           =   1  'Graphical
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   225
            Width           =   315
         End
         Begin VB.TextBox txtplnscode 
            Height          =   315
            Left            =   1470
            MaxLength       =   4
            TabIndex        =   40
            Tag             =   "SKIP"
            Top             =   975
            Width           =   615
         End
         Begin VB.TextBox txtplncode 
            Height          =   315
            Left            =   1470
            MaxLength       =   4
            TabIndex        =   39
            Tag             =   "SKIP"
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtplcode 
            Height          =   315
            Left            =   1470
            MaxLength       =   3
            TabIndex        =   38
            Tag             =   "SKIP"
            Top             =   225
            Width           =   615
         End
         Begin VB.TextBox txtpldesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2475
            Locked          =   -1  'True
            MaxLength       =   64
            TabIndex        =   37
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   210
            Width           =   3975
         End
         Begin VB.TextBox txtplnsdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2475
            Locked          =   -1  'True
            MaxLength       =   64
            TabIndex        =   36
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   975
            Width           =   3975
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "PL Item Code :"
            Height          =   210
            Left            =   390
            TabIndex        =   46
            Top             =   1005
            Width           =   1020
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "PL Note Code :"
            Height          =   210
            Left            =   345
            TabIndex        =   45
            Top             =   630
            Width           =   1065
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "PL Code :"
            Height          =   210
            Left            =   705
            TabIndex        =   44
            Top             =   240
            Width           =   690
         End
      End
      Begin VB.TextBox txtacctdetl 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   1470
         MaxLength       =   4
         TabIndex        =   22
         Tag             =   "SKIP"
         Top             =   210
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   6510
         MaxLength       =   35
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   15
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TextAccType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   2475
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   975
         Width           =   3990
      End
      Begin VB.TextBox Textcrncy 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   2475
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   585
         Width           =   3990
      End
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2115
         Picture         =   "frmgldetail.frx":2A1E
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   975
         Width           =   315
      End
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2115
         Picture         =   "frmgldetail.frx":2B90
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   570
         Width           =   315
      End
      Begin VB.Frame Frame4 
         Caption         =   "Balace Sheet Routing Note #."
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   1380
         Left            =   30
         TabIndex        =   2
         Top             =   2205
         Width           =   6495
         Begin VB.TextBox txtbnsdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2475
            Locked          =   -1  'True
            MaxLength       =   64
            TabIndex        =   34
            TabStop         =   0   'False
            Tag             =   "SKIPN"
            Top             =   975
            Width           =   3975
         End
         Begin VB.TextBox txtbdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2475
            Locked          =   -1  'True
            MaxLength       =   64
            TabIndex        =   30
            TabStop         =   0   'False
            Tag             =   "SKIPN"
            Top             =   210
            Width           =   3975
         End
         Begin VB.TextBox txtbcode 
            Height          =   315
            Left            =   1470
            MaxLength       =   3
            TabIndex        =   29
            Tag             =   "SKIPN"
            Top             =   225
            Width           =   615
         End
         Begin VB.TextBox txtbncode 
            Height          =   315
            Left            =   1470
            MaxLength       =   4
            TabIndex        =   28
            Tag             =   "SKIPN"
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtbndesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2475
            Locked          =   -1  'True
            MaxLength       =   64
            TabIndex        =   27
            TabStop         =   0   'False
            Tag             =   "SKIPN"
            Top             =   585
            Width           =   3975
         End
         Begin VB.TextBox txtbnscode 
            Height          =   315
            Left            =   1470
            MaxLength       =   4
            TabIndex        =   26
            Tag             =   "SKIPN"
            Top             =   975
            Width           =   615
         End
         Begin VB.CommandButton cmdLookup0 
            Height          =   315
            Left            =   2115
            Picture         =   "frmgldetail.frx":2D02
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   225
            Width           =   315
         End
         Begin VB.CommandButton cmdLookup1 
            Height          =   315
            Left            =   2115
            Picture         =   "frmgldetail.frx":2E74
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   585
            Width           =   315
         End
         Begin VB.CommandButton Command3 
            Height          =   315
            Left            =   2115
            Picture         =   "frmgldetail.frx":2FE6
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   960
            Width           =   315
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "BS Code :"
            Height          =   210
            Left            =   645
            TabIndex        =   33
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "BS Note Code :"
            Height          =   210
            Left            =   285
            TabIndex        =   32
            Top             =   630
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "BS Item Code :"
            Height          =   210
            Left            =   330
            TabIndex        =   31
            Top             =   1005
            Width           =   1050
         End
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2115
         Picture         =   "frmgldetail.frx":3158
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtAcctDesc 
         Height          =   315
         Left            =   2460
         MaxLength       =   80
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   210
         Width           =   3990
      End
      Begin MSMask.MaskEdBox txtoldacct 
         Height          =   315
         Left            =   1470
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1335
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSMask.MaskEdBox txtcrncycode 
         Height          =   315
         Left            =   1470
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   630
         _ExtentX        =   1111
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
      Begin MSMask.MaskEdBox txtacctType 
         Height          =   315
         Left            =   1470
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   975
         Width           =   645
         _ExtentX        =   1138
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Custom Code :"
         Height          =   210
         Left            =   3705
         TabIndex        =   53
         Top             =   1350
         Width           =   1050
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Currency :"
         Height          =   210
         Left            =   630
         TabIndex        =   17
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Old A/c# :"
         Height          =   210
         Left            =   660
         TabIndex        =   15
         Top             =   1350
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Acct. Type :"
         Height          =   210
         Left            =   510
         TabIndex        =   14
         Top             =   990
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Detail A/c :"
         Height          =   210
         Left            =   615
         TabIndex        =   6
         Top             =   240
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmgldetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ls_actbase As String
Dim LS_status As String
Dim ls_TFields As String
Dim ls_PFields As String
Dim ls_TotalSubs As String
Dim ls_PrvAlia As String
Dim lb_BlnkMast As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim Pr_GlRptGrp As Recordset
Dim PR_GlActType As Recordset
Dim PR_SyCurr As Recordset
Dim PR_GlDetl As Recordset
Dim PR_GlSub0 As Recordset
Dim pr_dumy As New Recordset
Dim pr_dumy1 As New Recordset
Dim PR_Dumy2 As New Recordset
Dim ls_sql As String

Private Function maxtranscode() As String
pr_dumy.Open "select max(Acct_Detail) as transcode from gl_Detail where acct_sub = '" & Trim(txtsub0.Text) & "' and   compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(Val(0 & pr_dumy("transcode"))) + 1)), 4)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 4)
End If
pr_dumy.Close
End Function
Private Sub cmdLookup_Click()
Dim ln_SetLen As Integer
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtsub0
    Set PO_DESC = txtSubDesc
    
    GoTop PR_GlSub0
    If PR_GlSub0.EOF Then
       ln_SetLen = 5
    Else
       ln_SetLen = IIf(Len(PR_GlSub0("MastAcctno")) < 3, 5, Len(PR_GlSub0("MastAcctno")))
    End If
    
    Gs_SQL = "Select " & ls_PFields & "  'Account No', Acct_Desc  'Description' from " & ls_PrvAlia & ""
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
    Gs_OrderBy = "Order by Acct_Desc"
    MyLookupOLDB.Caption = "Sub Accounts."
    MyLookupOLDB.Show 1
    
    If Len(txtsub0) > 0 Then txtsub0_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub cmdLookup0_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbcode
    Set PO_DESC = txtbdesc
    Gs_SQL = "Select Bcode 'Account No', BDesc  'Description' from GL_Bsheet1"
    Gs_FindFld = "BDesc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
    Gs_OrderBy = "Order by BCode"
    
    MyLookupOLDB.Caption = "Balance Notes"
    MyLookupOLDB.Show 1
    If Len(txtbcode) > 0 Then txtbcode_KeyDown vbKeyReturn, vbKeyShift
End Sub
Private Sub cmdLookup1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbncode
    Set PO_DESC = txtbndesc
    Gs_SQL = "Select Bncode 'Account No', BnDesc  'Description' from GL_Bsheet2"
    Gs_FindFld = "BnDesc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
    Gs_OrderBy = "Order by BnCode"
    
    MyLookupOLDB.Caption = "Balance Notes"
    MyLookupOLDB.Show 1
    If Len(txtbncode) > 0 Then txtbncode_KeyDown vbKeyReturn, vbKeyShift

End Sub
Private Sub Command3_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbnscode
    Set PO_DESC = txtbnsdesc
    Gs_SQL = "Select Bnicode 'Account No', BniDesc  'Description' from GL_Bsheet3"
    Gs_FindFld = "BniDesc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
    Gs_OrderBy = "Order by BniCode"
    
    MyLookupOLDB.Caption = "Balance Notes"
    MyLookupOLDB.Show 1
    If Len(txtbnscode) > 0 Then txtbnscode_KeyDown vbKeyReturn, vbKeyShift
End Sub
Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtplcode
    Set PO_DESC = txtpldesc
    Gs_SQL = "Select plcode 'Account No', plDesc  'Description' from GL_PLsheet1"
    Gs_FindFld = "PlDesc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
    Gs_OrderBy = "Order by plCode"
    
    MyLookupOLDB.Caption = "Profit n Loss Notes"
    MyLookupOLDB.Show 1
    If Len(txtplcode) > 0 Then txtplcode_KeyDown vbKeyReturn, vbKeyShift
End Sub
Private Sub Command4_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtplncode
    Set PO_DESC = txtplndesc
    Gs_SQL = "Select PLncode 'Account No', plnDesc  'Description' from GL_plsheet2"
    Gs_FindFld = "plnDesc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
    Gs_OrderBy = "Order by plnCode"
    
    MyLookupOLDB.Caption = "Profit n Loss Notes"
    MyLookupOLDB.Show 1
    If Len(txtplncode) > 0 Then txtplncode_KeyDown vbKeyReturn, vbKeyShift

End Sub
Private Sub Command5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtplnscode
    Set PO_DESC = txtplnsdesc
    Gs_SQL = "Select plnicode 'Account No', plniDesc  'Description' from GL_plsheet3"
    Gs_FindFld = "plniDesc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
    Gs_OrderBy = "Order by plniCode"
    
    MyLookupOLDB.Caption = "Profit n Loss Notes"
    MyLookupOLDB.Show 1
    If Len(txtplnscode) > 0 Then txtplnscode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command1_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtacctdetl
    Set PO_DESC = txtAcctDesc
    Gs_SQL = "Select Acct_detail 'Account No', Acct_Desc  'Description' from gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' and acct_sub = '" & txtsub0 & "'"
    Gs_OrderBy = "Order by AccountNo"
    Gs_Subon = True
    MyLookupOLDB.Caption = "Accounts Nos."
    MyLookupOLDB.Show 1
    If Len(txtacctdetl) > 0 Then txtacctdetl_KeyDown vbKeyReturn, vbKeyShift
End Sub






Private Sub Command6_Click()
' Currency code lookup
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCrncyCode
    Set PO_DESC = Textcrncy
    
    GoTop PR_SyCurr
    MyLookup.Caption = "Currency List."
    MyLookup.FillGrid PR_SyCurr, "Crncy_Code", "Crncy_Descrip", 5
    MyLookup.Show 1
    
    If Len(txtCrncyCode) > 0 Then txtcrncycode_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Command7_Click()
' Account Type lookup
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtacctType
    Set PO_DESC = TextAccType
    
    GoTop PR_GlActType
    MyLookup.Caption = "Currency List."
    MyLookup.FillGrid PR_GlActType, "AcctType", "AcctDescrip", 3
    MyLookup.Show 1
    
    If Len(txtacctType) > 0 Then txtacctType_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF11 Then Call DentMode(Mode, 4, PR_GlDetl, frmgldetail, IIf(Mode = "A", txtsub0, txtacctdetl), txtacctdetl, "X", "CompCount", 3, "X", "X", 1, False, Toolbar1)
End Sub

Private Sub Form_Load()
  'Dim ls_PrvAlia As String
  
  Dim ls_SqlString As String
  Dim ln_cnt As Integer
  
  txtsub0.MaxLength = 50
  'Label1.Caption = "Level <" + LTrim(RTrim(str(gn_Maxlevels))) + "> :"
  
  ls_PrvAlia = "Gl_Sub" + LTrim(RTrim(str(gn_Maxlevels)))
  ls_PFields = "Acct_Sub" + LTrim(str(gn_Maxlevels - 1)) + "+Acct_Sub" + LTrim(str(gn_Maxlevels))
  
  ls_TFields = ls_PFields + "+Acct_Detail"
  ls_SqlString = "Select GlGroupDetl.*,GlRpts_Ref.ReportBase as repobase,GlGroupDetl.GroupCode+GlRpts_Ref.ReportBase As RepoKey from GlGroupDetl INNER JOIN GlRpts_Ref on GlGroupDetl.CompCode+GlGroupDetl.ReportCode = GlRpts_Ref.CompCode+GlRpts_Ref.ReportCode where GLGroupdetl.compcode ='" & Gs_compcode & "' and left(GlGroupDetl.typegroup,1) = 'P' order by 2,3"
  
  SetToolBar(1) = chkRights("GLDETLAC01")
  SetToolBar(2) = chkRights("GLDETLAC02")
  SetToolBar(3) = chkRights("GLDETLAC03")
  SetToolBar(4) = chkRights("GLDETLAC04")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)

  Set Pr_GlRptGrp = New Recordset
  Set PR_GlActType = New Recordset
  Set PR_SyCurr = New Recordset
  Set PR_GlSub0 = New Recordset
  Set PR_GlDetl = New Recordset
  Gs_SQL = "Select " + ls_PFields + " 'Account No',Acct_Desc 'Description' from " + ls_PrvAlia + " "
  
  Pr_GlRptGrp.Open ls_SqlString, gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_GlActType.Open "Select * from Gl_AccTypes order by accttype", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_SyCurr.Open "Select * from Syscurrency order by crncy_code", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_GlSub0.Open "Select *," + ls_PFields + " As MastAcctNo from " + ls_PrvAlia + " where compcode ='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_GlDetl.Open "Select * from Gl_Detail where compcode ='" & Gs_compcode & "'", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
 
  lb_BlnkMast = IIf(PR_GlDetl.EOF, True, False)
  
  cmdLookup.Enabled = Not PR_GlSub0.EOF
  Command1.Enabled = Not lb_BlnkMast
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_SyCurr.Close
    PR_GlSub0.Close
    PR_GlDetl.Close
End Sub

Private Sub Option3_Click()
     txtbcode = ""
     txtbdesc = ""
     txtbncode = ""
     txtbndesc = ""
     txtbnscode = ""
     txtbnsdesc = ""
     
     txtplcode = ""
     txtpldesc = ""
     txtplncode = ""
     txtplndesc = ""
     txtplnscode = ""
     txtplnsdesc = ""
     ls_actbase = "S"
     Frame5.Enabled = False
     Frame4.Enabled = False
End Sub

Private Sub Option2_Click()
     Frame4.Enabled = True
     Frame5.Enabled = False
     ls_actbase = "B"
     txtplcode = ""
     txtpldesc = ""
     txtplncode = ""
     txtplndesc = ""
     txtplnscode = ""
     txtplnsdesc = ""
     
     Call SearchBSNoteAccount
End Sub
Private Sub ClearPLBSHeads()
     txtbcode = ""
     txtbdesc = ""
     txtbncode = ""
     txtbndesc = ""
     txtbnscode = ""
     txtbnsdesc = ""
     
     txtplcode = ""
     txtpldesc = ""
     txtplncode = ""
     txtplndesc = ""
     txtplnscode = ""
     txtplnsdesc = ""
End Sub
Private Sub Option1_Click()
     Frame5.Enabled = True
     Frame4.Enabled = False
     ls_actbase = "P"
     txtbcode = ""
     txtbdesc = ""
     txtbncode = ""
     txtbndesc = ""
     txtbnscode = ""
     txtbnsdesc = ""
     
     Call SearchPLNoteAccount
End Sub
Private Sub Option4_Click()
     Frame4.Enabled = True
     Frame5.Enabled = True
     ls_actbase = "O"
     Call SearchBSNoteAccount
     Call SearchPLNoteAccount
End Sub
Private Sub SearchPLNoteAccount()
ls_sql = "select * from Gl_PLSheet3Detail where compcode = '" & Gs_compcode & "' and left(accountno,9) = '" & Trim(txtsub0) & "'"
pr_dumy1.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1

If Not pr_dumy1.EOF Then
  txtplcode = Trim(pr_dumy1("plcode") & "")
  txtplncode = Trim(pr_dumy1("plncode") & "")
  txtplnscode = Trim(pr_dumy1("plnicode") & "")
 
  If Len(txtplcode) > 0 Then txtplcode_KeyDown vbKeyReturn, vbKeyShift
  If Len(txtplncode) > 0 Then txtplncode_KeyDown vbKeyReturn, vbKeyShift
  If Len(txtplnscode) > 0 Then txtplnscode_KeyDown vbKeyReturn, vbKeyShift
 

Else
  txtplcode.SetFocus
End If
pr_dumy1.Close
End Sub
Private Sub SearchBSNoteAccount()
ls_sql = "select * from Gl_BSheet3Detail where compcode = '" & Gs_compcode & "' and left(accountno,9) = '" & Trim(txtsub0) & "'"
pr_dumy1.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1

If Not pr_dumy1.EOF Then
  txtbcode = Trim(pr_dumy1("bcode") & "")
  txtbncode = Trim(pr_dumy1("bncode") & "")
  txtbnscode = Trim(pr_dumy1("bnicode") & "")
  
  If Len(txtbcode) > 0 Then txtbcode_KeyDown vbKeyReturn, vbKeyShift
  If Len(txtbncode) > 0 Then txtbncode_KeyDown vbKeyReturn, vbKeyShift
  If Len(txtbnscode) > 0 Then txtbnscode_KeyDown vbKeyReturn, vbKeyShift
  
Else
  txtbcode.SetFocus
End If
pr_dumy1.Close
End Sub

Private Sub SetValPLNoteAccount()
ls_sql = "select * from Gl_PLSheet3Detail where compcode = '" & Gs_compcode & "' and accountno = '" & Trim(txtsub0) + Trim(txtacctdetl) & "'"
PR_Dumy2.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1

If Not PR_Dumy2.EOF Then
  txtplcode = Trim(PR_Dumy2("plcode") & "")
  txtplncode = Trim(PR_Dumy2("plncode") & "")
  txtplnscode = Trim(PR_Dumy2("plnicode") & "")
 
  If Len(txtplcode) > 0 Then txtplcode_KeyDown vbKeyReturn, vbKeyShift
  If Len(txtplncode) > 0 Then txtplncode_KeyDown vbKeyReturn, vbKeyShift
  If Len(txtplnscode) > 0 Then txtplnscode_KeyDown vbKeyReturn, vbKeyShift
 

Else
If Frame5.Enabled Then txtplcode.SetFocus
End If
PR_Dumy2.Close
End Sub
Private Sub SetValBSNoteAccount()
ls_sql = "select * from Gl_BSheet3Detail where compcode = '" & Gs_compcode & "' and accountno = '" & Trim(txtsub0) + Trim(txtacctdetl) & "'"
PR_Dumy2.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1

If Not PR_Dumy2.EOF Then
  txtbcode = Trim(PR_Dumy2("bcode") & "")
  txtbncode = Trim(PR_Dumy2("bncode") & "")
  txtbnscode = Trim(PR_Dumy2("bnicode") & "")
  
  If Len(txtbcode) > 0 Then txtbcode_KeyDown vbKeyReturn, vbKeyShift
  If Len(txtbncode) > 0 Then txtbncode_KeyDown vbKeyReturn, vbKeyShift
  If Len(txtbnscode) > 0 Then txtbnscode_KeyDown vbKeyReturn, vbKeyShift
  
Else
 If Frame4.Enabled Then txtbcode.SetFocus
End If
PR_Dumy2.Close
End Sub





Private Sub txtAcctDesc_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn And LastkeyPressed(KeyCode) Then
        txtAcctDesc = UCase(txtAcctDesc)
        txtCrncyCode.SetFocus
     End If
End Sub

Private Sub txtacctdetl_Change()
If txtacctdetl <> "" Then
txtcustomCode = txtsub0 + txtacctdetl
End If
End Sub

Private Sub txtacctType_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
     If KeyCode = vbKeyReturn Then
          PR_GlActType.MoveFirst
          txtacctType.Text = UCase(txtacctType.Text)
          lb_found = MySeek(txtacctType.Text, "AcctType", PR_GlActType)
          If lb_found Then

             TextAccType.Text = PR_GlActType("AcctDescrip")
            If txtacctType.Enabled Then txtacctType.SetFocus
             If txtoldacct.Enabled Then txtoldacct.SetFocus
          Else
             Call SetErr("Invalid Currency Code.", vbCritical)
             txtacctType.SetFocus
          End If
    ElseIf KeyCode = vbKeyF12 Then
        Call Command7_Click
    End If
End Sub

Private Sub txtbcode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtbcode <> "" Then
pr_dumy.Open "Select * from Gl_Bsheet1 where compcode = '" & Gs_compcode & "' and bcode = '" & txtbcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    txtbdesc = Trim(pr_dumy("bdesc") & "")
Else
    Call MsgBox("Record not found", vbCritical)
    txtbcode.SetFocus
End If
pr_dumy.Close
End If
End Sub
Private Sub txtbncode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtbncode <> "" Then
pr_dumy.Open "Select * from Gl_Bsheet2 where compcode = '" & Gs_compcode & "' and bncode = '" & txtbncode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    txtbndesc = Trim(pr_dumy("bndesc") & "")
Else
    Call MsgBox("Record not found", vbCritical)
    txtbncode.SetFocus
End If
pr_dumy.Close
End If
End Sub
Private Sub txtbnscode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtbnscode <> "" Then
pr_dumy.Open "Select * from Gl_Bsheet3 where compcode = '" & Gs_compcode & "' and bnicode = '" & txtbnscode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    txtbnsdesc = Trim(pr_dumy("bnidesc") & "")
Else
    Call MsgBox("Record not found", vbCritical)
    txtbnscode.SetFocus
End If
pr_dumy.Close

End If
End Sub
Private Sub txtplcode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtplcode <> "" Then
pr_dumy.Open "Select * from Gl_plsheet1 where compcode = '" & Gs_compcode & "' and plcode = '" & txtplcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    txtpldesc = Trim(pr_dumy("pldesc") & "")
Else
    Call MsgBox("Record not found", vbCritical)
    txtplcode.SetFocus
End If
pr_dumy.Close

End If
End Sub
Private Sub txtplncode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtplncode <> "" Then
pr_dumy.Open "Select * from Gl_plsheet2 where compcode = '" & Gs_compcode & "' and plncode = '" & txtplncode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    txtplndesc = Trim(pr_dumy("plndesc") & "")
Else
    Call MsgBox("Record not found", vbCritical)
    txtplncode.SetFocus
End If
pr_dumy.Close
End If
End Sub
Private Sub txtplnscode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtplnscode <> "" Then
pr_dumy.Open "Select * from Gl_plsheet3 where compcode = '" & Gs_compcode & "' and plnicode = '" & txtplnscode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    txtplnsdesc = Trim(pr_dumy("plnidesc") & "")
Else
    Call MsgBox("Record not found", vbCritical)
    txtplnscode.SetFocus
End If
pr_dumy.Close
End If
End Sub
Private Sub txtcrncycode_GotFocus()
   If Mode = "A" Then
     txtCrncyCode.Text = Gs_BaseCrncy
   End If
End Sub

Private Sub txtcrncycode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
     If KeyCode = vbKeyReturn Then
          'PR_SyCurr.MoveFirst
          txtCrncyCode.Text = UCase(txtCrncyCode.Text)
          lb_found = MySeek(txtCrncyCode.Text, "Crncy_Code", PR_SyCurr)
          If lb_found Then
             Textcrncy.Text = PR_SyCurr("Crncy_Descrip")
           If txtacctType.Enabled Then txtacctType.SetFocus
          Else
             Call SetErr("Invalid Currency Code.", vbCritical)
             txtCrncyCode.SetFocus
          End If
    ElseIf KeyCode = vbKeyF12 Then
    Call Command6_Click
    End If
End Sub

Private Sub txtoldacct_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Option1.SetFocus
End Sub

Private Sub txtsub0_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn And Val(txtsub0.Text) <> 0 Then
         PR_GlSub0.MoveFirst
         lb_found = MySeek(txtsub0.Text, "MastAcctNo", PR_GlSub0)
        
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtsub0.SetFocus
             txtSubDesc.Text = ""
         Else
             txtSubDesc.Text = PR_GlSub0("Acct_Desc")
             If Mode = "A" Then
                txtacctdetl.Text = maxtranscode
                txtAcctDesc.SetFocus
             Else
                txtacctdetl.SetFocus
             End If
         End If
 ElseIf KeyCode = vbKeyF12 Then
    Call cmdLookup_Click
 End If
End Sub

Private Sub txtacctdetl_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn And Val(txtacctdetl.Text) > 0 Then
         PR_GlDetl.Requery
         txtacctdetl.Text = DoPad(txtacctdetl.Text, gn_DtlLen)
         PR_GlDetl.Filter = "acct_sub = '" & txtsub0.Text & "'"
         
         lb_found = MySeek(txtacctdetl.Text, "acct_detail", PR_GlDetl)
         
         Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   'SetClear Me
                   txtacctdetl.Text = ""
                   txtacctdetl.SetFocus
                Else
                   txtAcctDesc.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
'                   SetClear Me
                   txtacctdetl.Enabled = True
                   txtacctdetl.SetFocus
                Else
                   Call SetVal
                   If Mode <> "D" Then
                      'txtacctdetl.Enabled = False
                      txtAcctDesc.SetFocus
                   End If
                End If
            End Select
    ElseIf KeyCode = vbKeyF12 Then
        Call Command1_Click
    End If
  End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If Button.Index = 1 Then
         Frame2.Enabled = True
         txtacctdetl.MaxLength = gn_DtlLen
         Command1.Enabled = False
         cmdLookup.Enabled = True
    ElseIf Button.Index <> 4 Then
         Frame2.Enabled = True
         Command1.Enabled = True
         cmdLookup.Enabled = True
         txtsub0.SetFocus
    End If

    If lb_BlnkMast And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
    Else
       Mode = DentMode(Mode, Button.Index, PR_GlDetl, frmgldetail, txtsub0, txtacctdetl, "X", "CompCount", 3, "X", "X", 1, False, Toolbar1)
       'If Button.Index <> 4 Then txtsub0.SetFocus
    End If
    
End Sub

Public Sub SaveValues()
On Error GoTo LocalErr
Dim cntsql As New ADODB.Command
Dim ls_ActNo As String
lb_BlnkMast = False
ls_ActNo = txtsub0.Text + txtacctdetl.Text
LS_status = "D"
gc_dbcon.BeginTrans
     Select Case Mode
           Case "A"
              ls_sql = "INSERT into Gl_detail(compcode,Acct_sub,Acct_Detail,AccountNo,Acct_desc,crncy_code,Acct_Base,Acct_Type,Acct_Status,userid,adddate,addtime,Customcode) VALUES ('" & Gs_compcode & "','" & txtsub0.Text & "', '" & txtacctdetl.Text & " ','" & ls_ActNo & "','" & txtAcctDesc.Text & "','" & txtCrncyCode.Text & "','" & ls_actbase & "','" & txtacctType.Text & "','" & LS_status & "','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & txtcustomCode & "')"
              gc_dbcon.Execute ls_sql
              
           Case "E"
              ls_sql = "UPDATE Gl_Detail SET Acct_Desc= '" & txtAcctDesc.Text & "',crncy_code='" & txtCrncyCode.Text & "',Acct_Base='" & ls_actbase & "',Acct_Status='" & LS_status & "',Acct_Type='" & txtacctType.Text & "',OldAccount='" & txtoldacct.Text & "',Customcode='" & txtcustomCode.Text & "' WHERE  compcode = '" & Gs_compcode & "' and AccountNo= '" & ls_ActNo & "'"
              gc_dbcon.Execute ls_sql
           Case "D"
              ls_sql = "DELETE FROM Gl_Detail WHERE AccountNo = '" & ls_ActNo & "' and compcode = '" & Gs_compcode & "'"
              gc_dbcon.Execute ls_sql
     End Select
gc_dbcon.CommitTrans
'Profit and loss and balance sheet
If ls_actbase = "P" Then
ls_sql = "delete from gl_plsheet3detail where compcode = '" & Gs_compcode & "' and accountno = '" & ls_ActNo & "'"
gc_dbcon.Execute ls_sql

ls_sql = "insert into  gl_plsheet3detail (compcode,plcode,plncode,plnicode,accountno) values ('" & Gs_compcode & "','" & txtplcode & "','" & txtplncode & "','" & txtplnscode & "','" & ls_ActNo & "')"
gc_dbcon.Execute ls_sql

End If

If ls_actbase = "B" Then
ls_sql = "delete from gl_bsheet3detail where compcode = '" & Gs_compcode & "' and accountno = '" & ls_ActNo & "'"
gc_dbcon.Execute ls_sql

ls_sql = "insert into  gl_bsheet3detail (compcode,bcode,bncode,bnicode,accountno) values ('" & Gs_compcode & "','" & txtbcode & "','" & txtbncode & "','" & txtbnscode & "','" & ls_ActNo & "')"
gc_dbcon.Execute ls_sql

End If

If ls_actbase = "O" Then
ls_sql = "delete from gl_plsheet3detail where compcode = '" & Gs_compcode & "' and accountno = '" & ls_ActNo & "'"
gc_dbcon.Execute ls_sql

ls_sql = "insert into  gl_plsheet3detail (compcode,plcode,plncode,plnicode,accountno) values ('" & Gs_compcode & "','" & txtplcode & "','" & txtplncode & "','" & txtplnscode & "','" & ls_ActNo & "')"
gc_dbcon.Execute ls_sql

ls_sql = "delete from gl_bsheet3detail where compcode = '" & Gs_compcode & "' and accountno = '" & ls_ActNo & "'"
gc_dbcon.Execute ls_sql

ls_sql = "insert into  gl_bsheet3detail (compcode,bcode,bncode,bnicode,accountno) values ('" & Gs_compcode & "','" & txtbcode & "','" & txtbncode & "','" & txtbnscode & "','" & ls_ActNo & "')"
gc_dbcon.Execute ls_sql

End If

PR_GlDetl.Requery
Frame3.Enabled = False
Frame4.Enabled = False

Exit Sub

LocalErr:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
'Command1.Enabled = Not lb_BlnkMast
End Sub

Private Sub SetVal()
     txtAcctDesc = Trim(PR_GlDetl("Acct_Desc"))
     txtoldacct = IIf(IsNull(PR_GlDetl("OldAccount")), "", PR_GlDetl("OldAccount"))
     txtacctType = Trim(PR_GlDetl("Acct_Type"))
     txtCrncyCode = Trim(PR_GlDetl("crncy_code"))
     Option1.Value = IIf(PR_GlDetl("acct_base") = "P", True, False)
     Option2.Value = IIf(PR_GlDetl("acct_base") = "B", True, False)
     Option3.Value = IIf(PR_GlDetl("acct_base") = "S", True, False)
     Option4.Value = IIf(PR_GlDetl("acct_base") = "O", True, False)
     If Len(txtCrncyCode) > 0 Then txtcrncycode_KeyDown vbKeyReturn, vbKeyShift
     If Len(txtacctType) > 0 Then txtacctType_KeyDown vbKeyReturn, vbKeyShift
     
     ClearPLBSHeads
     SetValBSNoteAccount
     SetValPLNoteAccount
     
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtsub0) > 0 And Len(txtacctdetl) >= 0 And Len(RTrim(txtAcctDesc.Text)) > 0 And Len(txtCrncyCode.Text) > 0 Then
           ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function
Public Sub FrmRefresh()
  Pr_GlRptGrp.Requery
  PR_GlActType.Requery
  PR_SyCurr.Requery
  PR_GlSub0.Requery
  PR_GlDetl.Requery
End Sub

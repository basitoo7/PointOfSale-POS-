VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAskPosting 
   Caption         =   "Posting To Lease Ledger"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3330
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
   Icon            =   "frmAskPosting.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   3330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPost 
      Caption         =   "&Post"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   495
      MaskColor       =   &H00000000&
      TabIndex        =   11
      Top             =   1410
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1770
      TabIndex        =   10
      Top             =   1410
      Width           =   1155
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1935
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
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
      Height          =   1365
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3300
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   195
         Width           =   1425
      End
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   1440
         Picture         =   "frmAskPosting.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   180
         Width           =   315
      End
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   810
         TabIndex        =   3
         Top             =   540
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
         Format          =   22675457
         CurrentDate     =   37293
      End
      Begin MSComCtl2.DTPicker dtpto 
         Height          =   315
         Left            =   810
         TabIndex        =   4
         Top             =   930
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
         Format          =   22675457
         CurrentDate     =   37293
      End
      Begin MSMask.MaskEdBox txtbranchcode 
         Height          =   315
         Left            =   810
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Default Currency"
         Top             =   180
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Branch :"
         Height          =   210
         Left            =   135
         TabIndex        =   8
         Top             =   195
         Width           =   615
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "To :"
         Height          =   210
         Left            =   480
         TabIndex        =   2
         Top             =   960
         Width           =   270
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "From :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   300
         TabIndex        =   1
         Top             =   540
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmAskPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PO_CODE As Object
Public PO_DESC As Object

Dim PR_Branch As New Recordset
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdPost_Click()
Screen.MousePointer = vbHourglass
Select Case Me.Caption
Case "Posting To Lease Ledger"
    Call Module1.LMSLedgPost(dtpfrom, DTPTo, txtbranchcode)
Case "Posting To Module"
    Call Module2.IF_Posting(dtpfrom, DTPTo, txtbranchcode)
Case "Posting To Loan & Advances Ledger"
    Call Module2.ADV_Posting(dtpfrom, DTPTo, txtbranchcode)
End Select
Screen.MousePointer = vbDefault
End Sub

Private Sub Command5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbranchcode
    Set PO_DESC = Text1
    
    GoTop PR_Branch
    MyLookup.Caption = "Company Branches"
    MyLookup.FillGrid PR_Branch, "BranchCode", "BranchDesc", txtbranchcode.MaxLength
    MyLookup.Show 1

    If Len(txtbranchcode) > 0 Then txtbranchcode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub dtpfrom_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then DTPTo.SetFocus
End Sub
Private Sub dtpto_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then cmdPost.SetFocus
 If KeyCode = vbKeyPageUp Then dtpfrom.SetFocus
End Sub
Private Sub Form_Load()
    PR_Branch.Open "Select * From SysBranch Order By BranchCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
    
    txtbranchcode = Gs_BranchCode
    dtpfrom = Date
    DTPTo = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
PR_Branch.Close
End Sub

Private Sub txtbranchcode_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) And txtbranchcode <> "" Then
     txtbranchcode = DoPad(txtbranchcode, txtbranchcode.MaxLength)
     
     If Not MySeek(txtbranchcode, "Branchcode", PR_Branch) Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtbranchcode.SetFocus
     Else
        Text1 = PR_Branch("BranchDesc")
        dtpfrom.SetFocus
     End If
  End If
End Sub



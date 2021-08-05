VERSION 5.00
Begin VB.Form Mytree 
   Caption         =   "Account Tree"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MyTree.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   6480
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Account Tree"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   30
      TabIndex        =   0
      Top             =   -15
      Width           =   6435
      Begin VB.Label lbldetail 
         AutoSize        =   -1  'True
         Caption         =   "Cash in head office"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   1935
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblsub2 
         AutoSize        =   -1  'True
         Caption         =   "Cash in lhr branch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   1530
         TabIndex        =   4
         Top             =   945
         Width           =   1500
      End
      Begin VB.Label lblsub1 
         AutoSize        =   -1  'True
         Caption         =   "Cash in hand"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   210
         Left            =   900
         TabIndex        =   3
         Top             =   615
         Width           =   1065
      End
      Begin VB.Label lblsub0 
         AutoSize        =   -1  'True
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   2
         Top             =   270
         Width           =   420
      End
   End
   Begin VB.Label lblmsg 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   45
      TabIndex        =   1
      Top             =   2895
      Width           =   45
   End
End
Attribute VB_Name = "Mytree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PR_Sub0 As New Recordset
Dim PR_Sub1 As New Recordset
Dim PR_Sub2 As New Recordset
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Dim ls_Accountno As String
    ls_Accountno = MyLookupOLDB.ls_Accountno
    PR_Sub0.Open "select gl_sub0.* from gl_sub0 where compcode = '" & Gs_compcode & "' order by  Acct_Sub0 ", gc_dbcon, adOpenStatic, adLockPessimistic, adCmdText
    If MySeek(Trim(Left(ls_Accountno, gn_sublen(0))), "Acct_sub0", PR_Sub0) Then lblsub0.Caption = PR_Sub0("Acct_desc")
    
    PR_Sub1.Open "select gl_sub1.*,gl_sub1.acct_sub0+gl_sub1.acct_sub1 as Findfld from gl_sub1 where compcode = '" & Gs_compcode & "' order by  Findfld ", gc_dbcon, adOpenStatic, adLockPessimistic, adCmdText
    If MySeek(Trim(Left(ls_Accountno, gn_sublen(0) + gn_sublen(1))), "Findfld", PR_Sub1) Then lblsub1.Caption = PR_Sub1("Acct_desc")
    
    PR_Sub2.Open "select gl_sub2.*,gl_sub2.acct_sub1+gl_sub2.acct_sub2 as Findfld from gl_sub2 where compcode = '" & Gs_compcode & "' order by  Findfld ", gc_dbcon, adOpenStatic, adLockPessimistic, adCmdText
    If MySeek(Trim(Left(ls_Accountno, gn_sublen(0) + gn_sublen(1) + gn_sublen(2) + gn_sublen(3))), "Findfld", PR_Sub2) Then lblsub2.Caption = PR_Sub2("Acct_desc")
End Sub

Private Sub Form_Unload(Cancel As Integer)
         PR_Sub0.Close
         PR_Sub1.Close
         PR_Sub2.Close
End Sub


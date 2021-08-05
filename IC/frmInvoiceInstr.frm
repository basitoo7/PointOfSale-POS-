VERSION 5.00
Begin VB.Form frminvoiceinstr 
   Caption         =   "Invoice/Delivery Challan Instruction"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3060
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvoiceInstr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   3060
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Payment Type"
      Height          =   540
      Left            =   30
      TabIndex        =   4
      Top             =   2910
      Width           =   3015
      Begin VB.OptionButton OptCreditCard 
         Caption         =   "&Credit Card"
         Height          =   210
         Left            =   1800
         TabIndex        =   7
         Top             =   255
         Width           =   1125
      End
      Begin VB.OptionButton OptCheque 
         Caption         =   "Che&que"
         Height          =   210
         Left            =   885
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   870
      End
      Begin VB.OptionButton OptCash 
         Caption         =   "Ca&sh"
         Height          =   210
         Left            =   135
         TabIndex        =   5
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   1560
      TabIndex        =   3
      Top             =   3765
      Width           =   810
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   345
      Left            =   750
      TabIndex        =   2
      Top             =   3765
      Width           =   810
   End
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
      Height          =   2955
      Left            =   30
      TabIndex        =   0
      Top             =   -45
      Width           =   3015
      Begin VB.ListBox List1 
         Height          =   2700
         Left            =   60
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   165
         Width           =   2850
      End
   End
   Begin VB.Label Label1 
      Caption         =   "You can select five instructions at a time"
      Height          =   375
      Left            =   60
      TabIndex        =   8
      Top             =   3480
      Width           =   2955
   End
End
Attribute VB_Name = "frminvoiceinstr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PB_BlnkLoca As Boolean
Dim PR_Notes As New Recordset

Private Sub Command1_Click()
Dim ln_cnt  As Integer
For ln_cnt = 0 To List1.ListCount - 1
    If List1.Selected(ln_cnt) = True Then
    PO_AnyForm.PO_DESC.AddItem List1.List(ln_cnt)
   End If
Next
If OptCash = True Then
  PO_AnyForm.PO_CODE = "S"
ElseIf OptCheque = True Then
  PO_AnyForm.PO_CODE = "Q"
Else
  PO_AnyForm.PO_CODE = "C"
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
    List1.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
List1.Clear
PR_Notes.Open "Select * from IC_notes where compcode ='" & Gs_compcode & "' ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
While Not PR_Notes.EOF
    List1.AddItem PR_Notes("Description")
    PR_Notes.MoveNext
Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
PR_Notes.Close
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call Command1_Click
End Sub



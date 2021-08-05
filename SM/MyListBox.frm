VERSION 5.00
Begin VB.Form MyListBox 
   Caption         =   "List Box :"
   ClientHeight    =   3300
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
   Icon            =   "MyListBox.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   3060
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   1560
      TabIndex        =   3
      Top             =   2940
      Width           =   810
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   345
      Left            =   750
      TabIndex        =   2
      Top             =   2940
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
End
Attribute VB_Name = "MyListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim ln_cnt  As Integer
For ln_cnt = 0 To List1.ListCount - 1
    If List1.Selected(ln_cnt) = True Then
    PO_AnyForm.PO_DESC.AddItem List1.List(ln_cnt)
   End If
Next
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
End Sub
Public Sub Filllist(anyRS As Recordset, codeFld As String, DesFld As String, Optional codeSize As Integer)
Dim ln_cnt As Integer

    GoTop anyRS
    anyRS.Sort = DesFld
    With List1
        While Not anyRS.EOF
            List1.AddItem Trim(anyRS.Fields(codeFld) & "")
            anyRS.MoveNext
        Wend
    End With
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call Command1_Click
End Sub

Private Sub list1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    With List1
        'PO_AnyForm.PO_CODE = RTrim(LTrim(.TextMatrix(.Row, 0)))
        
    End With
    Unload Me
  'ElseIf Button = 1 Then
  '     grdLookUp.RowSel = grdLookUp.Row
  End If
End Sub


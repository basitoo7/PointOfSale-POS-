VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form MyLookup 
   Caption         =   "Look up :"
   ClientHeight    =   2880
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
   Icon            =   "MyLookup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6480
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
      Height          =   2235
      Left            =   30
      TabIndex        =   3
      Top             =   600
      Width           =   6435
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdLookUp 
         Height          =   2055
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3625
         _Version        =   393216
         FixedCols       =   0
         WordWrap        =   -1  'True
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         MousePointer    =   14
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
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
      Height          =   615
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.TextBox SeekText 
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         MaxLength       =   50
         TabIndex        =   2
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Text :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "MyLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    grdLookUp.SetFocus
    grdLookUp.Row = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
    InitializeGrid
End Sub

Public Sub InitializeGrid()
Dim ln_cnt As Integer

    With grdLookUp
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Code|<Description"
        .ColWidth(1) = PO_AnyForm.PO_DESC.MaxLength * 150
        
        For ln_cnt = 2 To Gn_GridCols
          .FormatString = .FormatString & "|<" & GR_GridCols(ln_cnt - 1, 2)
          .ColWidth(ln_cnt) = GR_GridCols(ln_cnt - 1, 3) * 150
        Next
        
        .Redraw = True
        .Redraw = True
    End With
End Sub

Public Sub FillGrid(anyRS As Recordset, codeFld As String, DesFld As String, Optional codeSize As Integer)
Dim ln_cnt As Integer

    GoTop anyRS
   ' anyRS.Sort = DesFld
    With grdLookUp
        .ColWidth(0) = IIf(codeSize > 0, codeSize, PO_AnyForm.PO_CODE.MaxLength) * IIf(codeSize > 12, 100, 150)
       
        While Not anyRS.EOF
            .Row = .Rows - 1
            .TextMatrix(.Row, 0) = Trim(anyRS.Fields(codeFld) & "")
            .TextMatrix(.Row, 1) = Trim(anyRS.Fields(DesFld) & "")
            
            For ln_cnt = 2 To Gn_GridCols
            .TextMatrix(.Row, ln_cnt) = Trim(anyRS.Fields(GR_GridCols(ln_cnt - 1, 1)) & "")
            Next
            .Rows = .Rows + 1
            anyRS.MoveNext
        Wend
        If .Rows > 2 And Val(.TextMatrix(.Rows - 1, 0)) = 0 Then .RemoveItem .Rows - 1
    End With
End Sub

Private Sub grdLookUp_KeyPress(KeyAscii As Integer)
Dim ln_cnt As Integer
    If Range(KeyAscii, 48, 57) Or Range(KeyAscii, 65, 90) Or Range(KeyAscii, 97, 122) Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 45 Or KeyAscii = 95 Then
       grdLookUp.Row = IIf(SeekText.Text = "", 1, grdLookUp.Row)
       SendKeys "{Left}"
       SeekText.Text = SeekText.Text + Chr(KeyAscii)
       With grdLookUp
       For ln_cnt = .Row To .Rows - 1
           If StrComp(SeekText.Text, Left(.TextMatrix(ln_cnt, 1), Len(SeekText.Text)), 1) = 0 Then
              .Row = ln_cnt
              If ln_cnt <> grdLookUp.Row Then SendKeys "{Left}"
              Exit For
           End If
       Next
       .SetFocus
       End With
    ElseIf KeyAscii = 13 Then
         Call grdLookUp_MouseDown(2, vbKeyShift, 0, 0)
    ElseIf KeyAscii = 8 Then
         grdLookUp.Row = 1
       If Len(SeekText.Text) <= 1 Then
         SeekText = ""
       Else
         SeekText.Text = Left(SeekText.Text, (Len(SeekText.Text) - 1))
         With grdLookUp
         For ln_cnt = .Row To .Rows - 1
           If StrComp(Trim(SeekText.Text), Left(.TextMatrix(ln_cnt, 1), Len(SeekText.Text)), 1) = 0 Then
              .Row = ln_cnt
              If ln_cnt <> grdLookUp.Row Then SendKeys "{Left}"
              Exit For
           End If
          Next
           .SetFocus
          End With
       End If
       If ln_cnt <> grdLookUp.Row Then SendKeys "{Left}"
    End If
End Sub

Private Sub grdLookUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    With grdLookUp
        PO_AnyForm.PO_CODE = RTrim(LTrim(.TextMatrix(.Row, 0)))
        PO_AnyForm.PO_DESC = .TextMatrix(.Row, 1)
    End With
    Unload Me
  ElseIf Button = 1 Then
       grdLookUp.RowSel = grdLookUp.Row
  End If
End Sub


VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmARAgging 
   Caption         =   "Aging"
   ClientHeight    =   2610
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   2835
   Icon            =   "FrmARAgging.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   2835
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   30
      TabIndex        =   10
      Top             =   -60
      Width           =   2805
      Begin VB.TextBox txtAgeColFrom 
         Height          =   345
         Index           =   3
         Left            =   750
         MaxLength       =   3
         TabIndex        =   6
         Top             =   1665
         Width           =   705
      End
      Begin VB.TextBox txtAgeColFrom 
         Height          =   345
         Index           =   2
         Left            =   750
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1200
         Width           =   705
      End
      Begin VB.TextBox txtAgeColFrom 
         Height          =   345
         Index           =   1
         Left            =   750
         MaxLength       =   3
         TabIndex        =   2
         Top             =   735
         Width           =   705
      End
      Begin VB.TextBox TxtAgeColTo 
         Height          =   345
         Index           =   3
         Left            =   1950
         MaxLength       =   3
         TabIndex        =   7
         Top             =   1641
         Width           =   705
      End
      Begin VB.TextBox TxtAgeColTo 
         Height          =   345
         Index           =   2
         Left            =   1950
         MaxLength       =   3
         TabIndex        =   5
         Top             =   1184
         Width           =   705
      End
      Begin VB.TextBox TxtAgeColTo 
         Height          =   345
         Index           =   1
         Left            =   1950
         MaxLength       =   3
         TabIndex        =   3
         Top             =   727
         Width           =   705
      End
      Begin VB.TextBox TxtAgeColTo 
         Height          =   345
         Index           =   0
         Left            =   1950
         MaxLength       =   3
         TabIndex        =   1
         Top             =   270
         Width           =   705
      End
      Begin VB.TextBox txtAgeColFrom 
         Height          =   345
         Index           =   0
         Left            =   750
         MaxLength       =   3
         TabIndex        =   0
         Top             =   270
         Width           =   705
      End
      Begin VB.Label Label12 
         Caption         =   "To"
         Height          =   225
         Left            =   1590
         TabIndex        =   18
         Top             =   1740
         Width           =   195
      End
      Begin VB.Label Label11 
         Caption         =   "From"
         Height          =   225
         Left            =   180
         TabIndex        =   17
         Top             =   1740
         Width           =   465
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   225
         Left            =   1590
         TabIndex        =   16
         Top             =   1260
         Width           =   195
      End
      Begin VB.Label Label8 
         Caption         =   "From"
         Height          =   225
         Left            =   180
         TabIndex        =   15
         Top             =   1260
         Width           =   465
      End
      Begin VB.Label Label6 
         Caption         =   "To"
         Height          =   225
         Left            =   1590
         TabIndex        =   14
         Top             =   780
         Width           =   195
      End
      Begin VB.Label Label5 
         Caption         =   "From"
         Height          =   225
         Left            =   180
         TabIndex        =   13
         Top             =   780
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   225
         Left            =   1590
         TabIndex        =   12
         Top             =   330
         Width           =   195
      End
      Begin VB.Label Label2 
         Caption         =   "From"
         Height          =   225
         Left            =   180
         TabIndex        =   11
         Top             =   330
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1500
      TabIndex        =   9
      Top             =   2160
      Width           =   915
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   420
      MaskColor       =   &H00000000&
      TabIndex        =   8
      Top             =   2160
      Width           =   915
   End
   Begin Crystal.CrystalReport rptAging 
      Left            =   -90
      Top             =   2220
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
End
Attribute VB_Name = "FrmARAgging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
Dim l_Cnt As Integer
For l_Cnt = 0 To TxtAgeColTo.UBound
    If IsNumeric(txtAgeColFrom(l_Cnt)) = False Then
        SetErr "Please enter Numeric Value.", vbCritical
        txtAgeColFrom(l_Cnt).SetFocus
        Exit Sub
    ElseIf IsNumeric(TxtAgeColTo(l_Cnt)) = False Then
         SetErr "Please enter Numeric Value.", vbCritical
         TxtAgeColTo(l_Cnt).SetFocus
       Exit Sub
    End If
Next
Ar_Ageing
    With rptAging
        .ReportFileName = App.Path & Gs_ARRepoPath & "\Aging.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "1 = " & txtAgeColFrom(0).Text
        .Formulas(2) = "3 = " & txtAgeColFrom(1).Text
        .Formulas(3) = "5 =" & txtAgeColFrom(2).Text
        .Formulas(4) = "7 =" & txtAgeColFrom(3).Text
        
        .Formulas(5) = "2 = " & TxtAgeColTo(0).Text
        .Formulas(6) = "4 =" & TxtAgeColTo(1).Text
        .Formulas(7) = "6 =" & TxtAgeColTo(2).Text
        .Formulas(8) = "8 =" & TxtAgeColTo(3).Text
        .Action = 1
    End With
gc_dbcon.Execute ("DROP TABLE Tmp_ArAging;")
End Sub

Private Sub txtAgeColFrom_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Select Case Index
Case 0
If txtAgeColFrom(0).Text <> "" Then TxtAgeColTo(0).SetFocus

Case 1
If txtAgeColFrom(1).Text <> "" Then TxtAgeColTo(1).SetFocus

Case 2
If txtAgeColFrom(2).Text <> "" Then TxtAgeColTo(2).SetFocus

Case 3
If txtAgeColFrom(3).Text <> "" Then TxtAgeColTo(3).SetFocus

Case 4
If txtAgeColFrom(4).Text <> "" Then TxtAgeColTo(4).SetFocus

End Select
End If
End Sub

Private Sub TxtAgeColTo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 If Index < 3 Then
    txtAgeColFrom(Index + 1).SetFocus
 Else
     cmdGenerate.SetFocus
 End If
 End If
    
End Sub

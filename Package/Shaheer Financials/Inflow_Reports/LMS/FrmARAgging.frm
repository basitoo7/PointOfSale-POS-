VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmARAgging 
   Caption         =   "Lease Rentals Aging"
   ClientHeight    =   4200
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   2835
   Icon            =   "FrmARAgging.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   2835
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3060
      Left            =   15
      TabIndex        =   10
      Top             =   -60
      Width           =   2805
      Begin VB.TextBox TxtAgeColTo 
         Height          =   345
         Index           =   5
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   22
         Top             =   2580
         Width           =   705
      End
      Begin VB.TextBox TxtAgeColTo 
         Height          =   345
         Index           =   4
         Left            =   1950
         MaxLength       =   3
         TabIndex        =   21
         Top             =   2100
         Width           =   705
      End
      Begin VB.TextBox txtAgeColFrom 
         Height          =   345
         Index           =   5
         Left            =   750
         MaxLength       =   3
         TabIndex        =   20
         Top             =   2580
         Width           =   705
      End
      Begin VB.TextBox txtAgeColFrom 
         Height          =   345
         Index           =   4
         Left            =   750
         MaxLength       =   3
         TabIndex        =   19
         Top             =   2130
         Width           =   705
      End
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
      Begin VB.Label Label10 
         Caption         =   "To"
         Height          =   225
         Left            =   1680
         TabIndex        =   26
         Top             =   2655
         Width           =   195
      End
      Begin VB.Label Label7 
         Caption         =   "To"
         Height          =   225
         Left            =   1680
         TabIndex        =   25
         Top             =   2175
         Width           =   195
      End
      Begin VB.Label Label4 
         Caption         =   "From"
         Height          =   225
         Left            =   210
         TabIndex        =   24
         Top             =   2640
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   225
         Left            =   210
         TabIndex        =   23
         Top             =   2190
         Width           =   465
      End
      Begin VB.Label Label12 
         Caption         =   "To"
         Height          =   225
         Left            =   1680
         TabIndex        =   18
         Top             =   1740
         Width           =   195
      End
      Begin VB.Label Label11 
         Caption         =   "From"
         Height          =   225
         Left            =   210
         TabIndex        =   17
         Top             =   1740
         Width           =   465
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   225
         Left            =   1680
         TabIndex        =   16
         Top             =   1260
         Width           =   195
      End
      Begin VB.Label Label8 
         Caption         =   "From"
         Height          =   225
         Left            =   210
         TabIndex        =   15
         Top             =   1260
         Width           =   465
      End
      Begin VB.Label Label6 
         Caption         =   "To"
         Height          =   225
         Left            =   1680
         TabIndex        =   14
         Top             =   780
         Width           =   195
      End
      Begin VB.Label Label5 
         Caption         =   "From"
         Height          =   225
         Left            =   210
         TabIndex        =   13
         Top             =   780
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   225
         Left            =   1680
         TabIndex        =   12
         Top             =   330
         Width           =   195
      End
      Begin VB.Label Label2 
         Caption         =   "From"
         Height          =   225
         Left            =   210
         TabIndex        =   11
         Top             =   330
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   1005
      Left            =   1440
      TabIndex        =   9
      Top             =   3105
      Width           =   1275
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
      Height          =   1005
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   8
      Top             =   3105
      Width           =   1275
   End
   Begin Crystal.CrystalReport rptAging 
      Left            =   2445
      Top             =   4125
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

Module1.ChkTempTables "Tmp_Accruals", False
Screen.MousePointer = vbHourglass
'Module2.Ar_Ageing (DATE)
Module2.Ar_Ageing2 (Date)
Screen.MousePointer = vbDefault
    With rptAging
        .ReportFileName = App.Path & Gs_ARRepoPath & "\Aging.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(14) = "UpToPeriod = '" & "As On : " & Date & "'"
        .Formulas(1) = "1 = " & txtAgeColFrom(0).Text
        .Formulas(2) = "3 = " & txtAgeColFrom(1).Text
        .Formulas(3) = "5 =" & txtAgeColFrom(2).Text
        .Formulas(4) = "7 =" & txtAgeColFrom(3).Text
        .Formulas(5) = "9 =" & txtAgeColFrom(4).Text
        .Formulas(6) = "11 =" & txtAgeColFrom(5).Text
        
        .Formulas(7) = "2 = " & TxtAgeColTo(0).Text
        .Formulas(8) = "4 =" & TxtAgeColTo(1).Text
        .Formulas(9) = "6 =" & TxtAgeColTo(2).Text
        .Formulas(10) = "8 =" & TxtAgeColTo(3).Text
        .Formulas(11) = "10 =" & TxtAgeColTo(4).Text
        .Formulas(12) = "12 =" & TxtAgeColTo(5).Text
        .Formulas(13) = "13 =" & TxtAgeColTo(5).Text
        .SelectionFormula = "({Tmp_ArAging.AgeCol1}+{Tmp_ArAging.AgeCol2}+{Tmp_ArAging.AgeCol3}+{Tmp_ArAging.AgeCol4}+{Tmp_ArAging.AgeCol5}+{Tmp_ArAging.AgeCol6}+{Tmp_ArAging.AgeCol7}) >100"
        .Action = 1
    End With
gc_dbcon.Execute ("DROP TABLE Tmp_ARAging;")
End Sub

Private Sub DTPAson_KeyDown(KeyCode As Integer, Shift As Integer)
   If Lastkey(KeyCode) Then cmdGenerate.SetFocus
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

Case 5
If txtAgeColFrom(5).Text <> "" Then TxtAgeColTo(5).SetFocus

End Select
End If
End Sub

Private Sub TxtAgeColTo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 If Index < 5 Then
    txtAgeColFrom(Index + 1).SetFocus
 Else
     
 End If
 End If
    
End Sub

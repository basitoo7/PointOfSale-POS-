VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmanalysisReport 
   Caption         =   "Analysis Report"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AnalysisReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4830
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3855
      TabIndex        =   8
      Top             =   1800
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
      Left            =   2880
      MaskColor       =   &H00000000&
      TabIndex        =   7
      Top             =   1800
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Height          =   1770
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   4740
      Begin VB.ComboBox txtlevel 
         Height          =   330
         ItemData        =   "AnalysisReport.frx":030A
         Left            =   1365
         List            =   "AnalysisReport.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1335
         Width           =   3135
      End
      Begin VB.TextBox txtbranchcode 
         Height          =   315
         Left            =   1365
         MaxLength       =   3
         TabIndex        =   10
         Top             =   240
         Width           =   390
      End
      Begin VB.TextBox txtbranchname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2130
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   9
         Tag             =   "SKIP"
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   1785
         Picture         =   "AnalysisReport.frx":0371
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   315
      End
      Begin Crystal.CrystalReport rptTrial 
         Left            =   3435
         Top             =   1590
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
      Begin MSComCtl2.DTPicker dtpicker1 
         Height          =   315
         Left            =   1365
         TabIndex        =   2
         Top             =   600
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
         CustomFormat    =   "MM/yyyy"
         Format          =   54788099
         CurrentDate     =   37293
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   1365
         TabIndex        =   3
         Top             =   975
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
         CustomFormat    =   "MM/yyyy"
         Format          =   54788099
         CurrentDate     =   37293
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Base On:"
         Height          =   210
         Left            =   660
         TabIndex        =   12
         Top             =   1380
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "From Month :"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   615
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Branch Code :"
         Height          =   210
         Left            =   315
         TabIndex        =   5
         Top             =   255
         Width           =   1035
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "To Month :"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1005
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmanalysisReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PR_Branch As New Recordset
Public PO_CODE As Object
Public PO_DESC As Object
Public ps_Head1 As String
Public ps_Head2 As String
Public ps_Head3 As String

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdGenerate_Click()
         MDIForm1.StatusBar1.Panels(7).Text = "Processing Data  Please Wait..."
         Module1.ChkTempTables "Tmp_GlAnalysis", True
         Module1.ChkTempTables "Tmp_Coa", True
         Module1.ChkTempTables "Tmp_months", True
         
         Module2.gl_analysis txtbranchcode, DTPicker1.Value, DTPicker2.Value
         With rptTrial
              .WindowTitle = Me.Caption
              .ReportFileName = App.Path & Gs_GlRepoPath & "\AnalysisTrialBalance.RPT"
              .SelectionFormula = ""
              .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
              .Formulas(1) = "ReportName = '" & " Analysis Report " + "(" + txtlevel + ")" & "'"
              .Formulas(9) = "Trialason = '" & "From:  " & MonthName(Month(DTPicker1)) & "," & Year(DTPicker1) & "   To:  " & MonthName(Month(DTPicker2)) & "," & Year(DTPicker2) & "'"
                If txtlevel = "Control Accounts" Then
                    .Formulas(2) = "Showsub0 = 0"
                    .Formulas(3) = "Showsub1 = 0"
                    .Formulas(4) = "Showsub2 = 0"
                    .Formulas(5) = "ShowDetail = 0"
                    .Formulas(6) = "Showsub00 = 1"
                    .Formulas(7) = "Showsub11 = 0"
                    .Formulas(8) = "Showsub22 = 0"
                ElseIf txtlevel = "Detail Accounts" Then
                    .Formulas(2) = "Showsub0 = 1"
                    .Formulas(3) = "Showsub1 = 0"
                    .Formulas(4) = "Showsub2 = 0"
                    .Formulas(5) = "ShowDetail = 0"
                    .Formulas(6) = "Showsub00 = 0"
                    .Formulas(7) = "Showsub11 = 1"
                    .Formulas(8) = "Showsub22 = 0"
                ElseIf txtlevel = "Sub Ledger Accounts" Then
                    .Formulas(2) = "Showsub0 = 1"
                    .Formulas(3) = "Showsub1 = 1"
                    .Formulas(4) = "Showsub2 = 0"
                    .Formulas(5) = "ShowDetail = 0"
                    .Formulas(6) = "Showsub00 = 0"
                    .Formulas(7) = "Showsub11 = 0"
                    .Formulas(8) = "Showsub22 = 1"
                ElseIf txtlevel = "Subsidiary Sub Ledger A/c" Then
                    .Formulas(2) = "Showsub0 = 1"
                    .Formulas(3) = "Showsub1 = 1"
                    .Formulas(4) = "Showsub2 = 1"
                    .Formulas(5) = "ShowDetail = 1"
                    .Formulas(6) = "Showsub00 = 1"
                    .Formulas(7) = "Showsub11 = 1"
                    .Formulas(8) = "Showsub22 = 1"
                End If
              .Connect = "DNS=Censoft;UID=Sa"
              .Action = 1
         End With
         gc_dbcon.Execute ("DROP TABLE Tmp_Coa;")
         gc_dbcon.Execute ("DROP TABLE Tmp_GlAnalysis;")
         gc_dbcon.Execute ("DROP TABLE Tmp_months;")
         MDIForm1.StatusBar1.Panels(7).Text = "Processing Data  Please Wait..."
       
End Sub

Private Sub Command5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbranchcode
    Set PO_DESC = txtbranchname
    
    GoTop PR_Branch
    MyLookup.Caption = "Company Branches"
    MyLookup.FillGrid PR_Branch, "BranchCode", "BranchDesc", txtbranchcode.MaxLength
    MyLookup.Show 1

    If Len(txtbranchcode) > 0 Then txtBranchCode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Form_Activate()
   DTPicker1 = Date
   DTPicker2 = Date
   txtlevel = "Control Accounts"
End Sub

Private Sub Form_Load()
    PR_Branch.Open "Select * From SysBranch Where compcode = '" & Gs_compcode & "' Order By BranchCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
    txtbranchcode = Gs_BranchCode
End Sub

Private Sub Form_Unload(Cancel As Integer)
  PR_Branch.Close
End Sub

Private Sub txtBranchCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) And txtbranchcode <> "" Then
     txtbranchcode = DoPad(txtbranchcode, txtbranchcode.MaxLength)
     
     If Not MySeek(txtbranchcode, "Branchcode", PR_Branch) Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtbranchcode.SetFocus
     Else
        txtbranchname = PR_Branch("BranchDesc")
        DTPicker1.SetFocus
     End If
  ElseIf KeyCode = vbKeyF12 Then
     Command5_Click
  End If
End Sub


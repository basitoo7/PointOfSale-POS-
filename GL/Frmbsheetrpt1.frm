VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frmbsheetrpt1 
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4605
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frmbsheetrpt1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4605
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkwithdiff 
      Caption         =   "With Two Dates Difference"
      Height          =   240
      Left            =   60
      TabIndex        =   11
      Top             =   1440
      Width           =   2310
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
      Height          =   405
      Left            =   2430
      MaskColor       =   &H00000000&
      TabIndex        =   8
      Top             =   1380
      Width           =   1037
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   3525
      TabIndex        =   7
      Top             =   1380
      Width           =   1037
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00000080&
      Height          =   1380
      Left            =   45
      TabIndex        =   3
      Top             =   -30
      Width           =   4530
      Begin VB.TextBox txtbranchcode 
         Height          =   315
         Left            =   1275
         MaxLength       =   3
         TabIndex        =   0
         Top             =   225
         Width           =   405
      End
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   1695
         Picture         =   "Frmbsheetrpt1.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtbranchname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2025
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   1
         Tag             =   "SKIP"
         Top             =   210
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker dtpfromdate 
         Height          =   315
         Left            =   1275
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
         Format          =   64094209
         CurrentDate     =   37293
      End
      Begin MSComCtl2.DTPicker Dtptodate 
         Height          =   315
         Left            =   1275
         TabIndex        =   9
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
         Format          =   64094209
         CurrentDate     =   37293
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "To Date :"
         Height          =   255
         Left            =   255
         TabIndex        =   10
         Top             =   1005
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Branch Code :"
         Height          =   210
         Left            =   195
         TabIndex        =   6
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "From Date :"
         Height          =   255
         Left            =   255
         TabIndex        =   5
         Top             =   615
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport rptTrial 
      Left            =   5040
      Top             =   285
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4605
      Top             =   240
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
Attribute VB_Name = "Frmbsheetrpt1"
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
On Error GoTo LocalErr
MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."

If txtbranchname = "" Then
    Call txtBranchCode_KeyDown(vbKeyReturn, vbKeyShift)
End If
       If chkwithdiff.Value = 1 Then
         Module1.ChkTempTables "Tmp_GlActivity", True
         Module2.Balancesheetperiodic Dtptodate.Value, 1, , dtpfromdate.Value, txtbranchcode
         With rptTrial
         
              .Formulas(0) = ""
              .Formulas(1) = ""
              .Formulas(2) = ""
              .Formulas(3) = ""
              .Formulas(4) = ""
              .Formulas(5) = ""
              
              .WindowTitle = Me.Caption
              If Me.Caption = "Balance Sheet (Periodic)" Then
                .ReportFileName = App.Path & Gs_GlRepoPath & "\BPLSheet4.RPT"
              ElseIf Me.Caption = "Profit and Loss (Periodic)" Then
                .ReportFileName = App.Path & Gs_GlRepoPath & "\PNLSheet4.RPT"
              End If
              .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
              .Formulas(1) = "ReportName = '" & Me.Caption & "'"
              .Formulas(2) = "Trialason = '" & "As on " & str(Dtptodate) & "'"
              .Formulas(3) = "BRANCHNAME = '" & txtbranchcode + " - " + txtbranchname & "'"
              .Formulas(4) = "fromdate = '" & Format(dtpfromdate, "dd-mmm-yyyy") & "'"
              .Formulas(5) = "todate = '" & Format(Dtptodate, "dd-mmm-yyyy") & "'"
              .Connect = "DNS=Censoft;UID=Sa"
              .Action = 1
        End With
         gc_dbcon.Execute ("DROP TABLE Tmp_GlActivity;")
       Else
            Module1.ChkTempTables "Tmp_GlActivity", True
            Module1.ChkTempTables "Tmp_COA", True
            
           ' If Me.Caption = "Balance Sheet Notes (Periodic)" Or Me.Caption = "Balance Sheet (Periodic)" Then
            '   Module2.GlActivity DTPtodate.Value, 0, DTPfromdate.Value, txtbranchcode
            'Else
               Module2.GlActivity2 Dtptodate.Value, 2, dtpfromdate.Value, txtbranchcode
            'End If
             
             With rptTrial
              .Formulas(0) = ""
              .Formulas(1) = ""
              .Formulas(2) = ""
              .Formulas(3) = ""
              .Formulas(4) = ""
              .Formulas(5) = ""
              
              .WindowTitle = Me.Caption
              If Me.Caption = "Balance Sheet (Periodic)" Then
                .ReportFileName = App.Path & Gs_GlRepoPath & "\BPLSheet1.RPT"
              ElseIf Me.Caption = "Profit and Loss (Periodic)" Then
                .ReportFileName = App.Path & Gs_GlRepoPath & "\PNLSheet1.RPT"
              ElseIf Me.Caption = "Balance Sheet Notes (Periodic)" Then
                .ReportFileName = App.Path & Gs_GlRepoPath & "\BPLSheet2.RPT"
              ElseIf Me.Caption = "Profit and Loss Notes (Periodic)" Then
                .ReportFileName = App.Path & Gs_GlRepoPath & "\PNLSheet2.RPT"
              ElseIf Me.Caption = "Profit and Loss (Detail-Periodic)" Then
                .ReportFileName = App.Path & Gs_GlRepoPath & "\PNLSheet3.RPT"
              ElseIf Me.Caption = "Balance Sheet (Detail-Periodic)" Then
                .ReportFileName = App.Path & Gs_GlRepoPath & "\BPLSheet3.RPT"
             
              End If
              
              .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
              .Formulas(1) = "ReportName = '" & Me.Caption & "'"
              .Formulas(2) = "Trialason = '" & "From " & dtpfromdate & " to " & Dtptodate & "'"
              .Formulas(3) = "BRANCHNAME = '" & txtbranchcode + " - " + txtbranchname & "'"
               .Connect = "DNS=Censoft;UID=Sa"
              .Action = 1
        End With
         gc_dbcon.Execute ("DROP TABLE Tmp_GlActivity;")
              
       End If
        
MDIForm1.StatusBar1.Panels(7).Text = ""
Exit Sub

LocalErr:
MsgBox Err.Description
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
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
   
  dtpfromdate = Gs_Fnperiod
  Dtptodate = Date
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
       dtpfromdate.SetFocus
     End If
  ElseIf KeyCode = vbKeyF12 Then
     Command5_Click
  End If
End Sub


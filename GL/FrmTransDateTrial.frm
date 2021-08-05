VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTransDateTrial 
   Caption         =   "Trial Balance"
   ClientHeight    =   1440
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
   Icon            =   "FrmTransDateTrial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4605
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkgrouptotal 
      Caption         =   "With Group Head Total"
      Height          =   270
      Left            =   45
      TabIndex        =   7
      Top             =   1035
      Width           =   2250
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
      TabIndex        =   4
      Top             =   990
      Width           =   1037
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   3525
      TabIndex        =   3
      Top             =   990
      Width           =   1037
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00000080&
      Height          =   1005
      Left            =   45
      TabIndex        =   1
      Top             =   -30
      Width           =   4530
      Begin MSComCtl2.DTPicker dtpfromdate 
         Height          =   315
         Left            =   1275
         TabIndex        =   0
         Top             =   195
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
         Format          =   54460417
         CurrentDate     =   37293
      End
      Begin MSComCtl2.DTPicker Dtptodate 
         Height          =   315
         Left            =   1275
         TabIndex        =   5
         Top             =   570
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
         Format          =   54460417
         CurrentDate     =   37293
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "To Date :"
         Height          =   255
         Left            =   255
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "From Date :"
         Height          =   255
         Left            =   255
         TabIndex        =   2
         Top             =   210
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
Attribute VB_Name = "FrmTransDateTrial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ls_sql As String

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdGenerate_Click()
On Error GoTo LocalErr
MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."


         
         Module1.ChkTempTables "Tmp_TransTrial", True
         ls_sql = "SELECT accountno, SUM(Dr_Amount) AS Dr_Amount, SUM(Cr_Amount) AS Cr_Amount into Tmp_TransTrial From Gl_Trans"
         ls_sql = ls_sql & " WHERE Value_Date >= '" & Format(dtpfromdate, "YYYY/MM/DD") & "' AND Value_Date <= '" & Format(Dtptodate, "YYYY/MM/DD") & "' AND compcode = '" & Gs_compcode & "'"
         ls_sql = ls_sql & " GROUP BY accountno"
         gc_dbcon.Execute ls_sql
         
         With rptTrial
              .WindowTitle = Me.Caption
              If chkgrouptotal.Value = 1 Then
              .ReportFileName = App.Path & Gs_GlRepoPath & "\TransBaseTrialsum.rpt"
              Else
              .ReportFileName = App.Path & Gs_GlRepoPath & "\TransBaseTrial.rpt"
              End If
              .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
              .Formulas(1) = "ReportName = '" & Me.Caption & "'"
              .Formulas(2) = "Period = '" & "From " & dtpfromdate & " to " & Dtptodate & "'"
              .Connect = "DNS=Censoft;UID=Sa"
              .Action = 1
        End With
         
        
MDIForm1.StatusBar1.Panels(7).Text = ""
Exit Sub

LocalErr:
MsgBox Err.Description
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Private Sub Form_Activate()
  dtpfromdate = Date
  Dtptodate = Date
  End Sub


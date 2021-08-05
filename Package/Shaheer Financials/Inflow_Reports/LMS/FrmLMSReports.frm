VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmLMSReports 
   Caption         =   "Lease # :"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmLMSReports.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2745
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   4485
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   4110
         Picture         =   "FrmLMSReports.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   300
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CheckBox Chvdate 
         Alignment       =   1  'Right Justify
         Caption         =   "Active Lease Only :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2625
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1770
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   375
         Left            =   4170
         MaxLength       =   50
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1890
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2460
         MaxLength       =   50
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1965
      End
      Begin VB.CommandButton cmdLookup 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         Picture         =   "FrmLMSReports.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1020
         Width           =   315
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2460
         MaxLength       =   50
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   660
         Width           =   1965
      End
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   2100
         Picture         =   "FrmLMSReports.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   660
         Width           =   315
      End
      Begin VB.TextBox txtleaseno 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1290
         MaxLength       =   3
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1380
         Width           =   795
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2115
         Picture         =   "FrmLMSReports.frx":0760
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1380
         Width           =   315
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   855
         Left            =   2340
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1830
         Width           =   1395
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "&Generate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   855
         MaskColor       =   &H00000000&
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1830
         Width           =   1395
      End
      Begin Crystal.CrystalReport LMSReports 
         Left            =   60
         Top             =   1710
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   0   'False
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         DiscardSavedData=   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin MSComCtl2.DTPicker DTPTo 
         Height          =   315
         Left            =   3210
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   300
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
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
         Format          =   22806529
         CurrentDate     =   37293
      End
      Begin MSMask.MaskEdBox txtCustNO 
         Height          =   315
         Left            =   1290
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1020
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtpason 
         Height          =   315
         Left            =   1290
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   300
         Width           =   1155
         _ExtentX        =   2037
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
         Format          =   22806529
         CurrentDate     =   37293
      End
      Begin MSMask.MaskEdBox txtbranchcode 
         Height          =   315
         Left            =   1290
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Default Currency"
         Top             =   660
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRecoverer 
         Height          =   315
         Left            =   3510
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Default Currency"
         Top             =   300
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtSLRReturn 
         Height          =   315
         Left            =   1290
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Default Currency"
         Top             =   1035
         Visible         =   0   'False
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   11
         Format          =   "#,##0;(#,##0)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Recv.Off."
         Height          =   210
         Index           =   1
         Left            =   4725
         TabIndex        =   22
         Top             =   2340
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "To :"
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   330
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Branch Code :"
         Height          =   210
         Left            =   210
         TabIndex        =   16
         Top             =   690
         Width           =   1035
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Customer Code :"
         Height          =   315
         Left            =   45
         TabIndex        =   14
         Top             =   1065
         Width           =   1200
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Lease # :"
         Height          =   210
         Index           =   0
         Left            =   60
         TabIndex        =   13
         Top             =   1410
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "As On :"
         Height          =   255
         Left            =   690
         TabIndex        =   12
         Top             =   330
         Width           =   555
      End
   End
End
Attribute VB_Name = "FrmLMSReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PO_CODE As Object
Public PO_DESC As Object

Dim PR_Recoverer As New Recordset
Dim PR_Branch As New Recordset
Dim PR_LMSInfo As New Recordset
Dim pr_Customer As New Recordset
Dim ln_InflowStatus As Integer
Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdGenerate_Click()
Dim ls_Temp As String
MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
With LMSReports
 
        If txtleaseno = "" And txtCustNO <> "" Then
          .SelectionFormula = "{LM_Schedule.Compcode}+{LM_Schedule.BranchCode}+{LM_Schedule.CustomerNo} = '" & Gs_compcode + txtbranchcode + txtCustNO & "'"
          ls_Temp = "LM_Schedule.Compcode+LM_Schedule.BranchCode+LM_Schedule.CustomerNo = '" & Gs_compcode + txtbranchcode + txtCustNO & "'"
        ElseIf txtCustNO = "" And txtleaseno = "" Then
          .SelectionFormula = "{LM_Schedule.Compcode}+{LM_Schedule.BranchCode} = '" & Gs_compcode + txtbranchcode & "'"
          ls_Temp = "LM_Schedule.Compcode+LM_Schedule.BranchCode = '" & Gs_compcode + txtbranchcode & "'"
        ElseIf txtCustNO <> "" And txtleaseno <> "" Then
         If UCase(Left(Me.Caption, 4)) = "INFL" Or UCase(Left(Me.Caption, 4)) = "REAL" Then
           .SelectionFormula = "{LM_Schedule.Compcode}+{LM_Schedule.BranchCode}+{LM_Schedule.CustomerNo}+{LM_Schedule.LeaseNo} = '" & Gs_compcode + txtbranchcode + txtCustNO + txtleaseno & "'"
           ls_Temp = "LM_Schedule.Compcode+LM_Schedule.BranchCode+LM_Schedule.CustomerNo+LM_Schedule.FacilityAcct= '" & Gs_compcode + txtbranchcode + txtCustNO + txtleaseno & "'"
         Else
           .SelectionFormula = "{LM_Schedule.Compcode}+{LM_Schedule.BranchCode}+{LM_Schedule.CustomerNo}+{LM_Schedule.LeaseNo} = '" & Gs_compcode + txtbranchcode + txtCustNO + txtleaseno & "'"
           ls_Temp = "LM_Schedule.Compcode+LM_Schedule.BranchCode+LM_Schedule.CustomerNo+LM_Schedule.LeaseNo= '" & Gs_compcode + txtbranchcode + txtCustNO + txtleaseno & "'"
         End If
         
        End If
 
 Select Case UCase(Left(Me.Caption, 4))
     Case "CIB2" 'CIB2 Report
               Module1.ChkTempTables "Tmp_LMSData", True
               Module2.Cib2_repo (dtpason.Value)
              .ReportFileName = App.Path & Gs_ARRepoPath & "\CIB2.RPT"
              .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
              .Formulas(1) = "ReportName = '" & Me.Caption & "'"
              .SelectionFormula = ""
              .Action = 1
     Case "CLIE" 'Client Payment Schedule
              .ReportFileName = App.Path & Gs_ARRepoPath & "\ClntShdl.RPT"
              .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
              .Formulas(1) = "ReportName = '" & Me.Caption & "'"
              .Action = 1
     Case "INTE" 'Internal Payment Schedule
           If Left(Right(Me.Caption, 9), 1) = "R" Then
              .ReportFileName = App.Path & Gs_ARRepoPath & "\InterShdl.RPT"
           Else
              .ReportFileName = App.Path & Gs_ARRepoPath & "\AInterShdl.RPT"
           End If
              .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
              .Formulas(1) = "ReportName = '" & Me.Caption & "'"
              .Action = 1
     Case "BALA" ' Balance Confirmations
               If txtCustNO = "" Then
                  Call SetErr("You must give Customer No.", vbCritical)
                  txtCustNO.SetFocus
               Else
                  
                Call BalConfirm(txtCustNO.Text, txtleaseno.Text, dtpason.Value)
                .ReportFileName = App.Path & Gs_ARRepoPath & "\BalConfirm.RPT"
                .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
                .Formulas(1) = "ReportName = 'Statement of Account'"
                .Formulas(2) = "UpToPeriod = '" & "As On : " & dtpason & "'"
                .SelectionFormula = ""
                .Action = 1
                gc_dbcon.Execute ("Drop Table Tmp_BalConfirm;")
               End If
    Case "PROJ" ' Projections
            ls_Temp = ls_Temp + " And (LM_Schedule.AccrualDate Between '" & Format(dtpason.Value, "YYYY/MM/DD") & "' And '" & Format(DTPTo.Value, "YYYY/MM/DD") & "')  " & IIf(Chvdate.Value = 0, " And LM_Schedule.RentalStatus <> 'T'", "")
            Module1.ChkTempTables "Tmp_LMProject", False
            Module2.LM_Project ls_Temp
            '.LogOnServer("pdssql.dll", "server1 \ sql2000", "FCMFinancials", "sa", "") = 1
            .ReportFileName = App.Path & Gs_ARRepoPath & "\ProjectionReport.RPT"
            .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
            .Formulas(1) = "ReportName = '" & Me.Caption & "'"
            .Formulas(2) = "UptoPeriod = '" & "From : " & dtpason & " &  To " & DTPTo & "'"
            .SelectionFormula = ""
            .Action = 1
            gc_dbcon.Execute ("Drop Table Tmp_LMProject;")
    Case "RECO" ' Recovery Statement Overall
            ls_Temp = ls_Temp + IIf(txtRecoverer = "", "", " And LM_LeaseInfo.RecCode = '" & txtRecoverer & "'")
            Module1.ChkTempTables "Tmp_Process", True
            Module1.ChkTempTables "Tmp_LMRecovery", True
            Module1.ChkTempTables "Tmp_Payments", True
            Module1.ChkTempTables "Tmp_Payments1", True
            Module1.ChkTempTables "Tmp_LastRental", True
            
            Module2.LMS_Recovery ls_Temp, dtpason.Value
            .ReportFileName = App.Path & Gs_ARRepoPath & "\RecReportoverall.RPT"
            .DiscardSavedData = True
            .SelectionFormula = ""
            .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
            .Formulas(1) = "ReportName = '" & Me.Caption & "'"
            .Formulas(2) = "UpToPeriod = '" & "As On : " & dtpason & "'"
            .Action = 1
            gc_dbcon.Execute ("Drop Table Tmp_LMRecovery;")
            gc_dbcon.Execute ("Drop Table Tmp_Payments;")
            gc_dbcon.Execute ("Drop Table Tmp_LastRental;")
            gc_dbcon.Execute ("Drop Table Tmp_LastPaid;")
            gc_dbcon.Execute ("Drop Table Tmp_Process;")
    Case UCase("OVER") ' Overdue Statement
            ls_Temp = ls_Temp + IIf(txtRecoverer = "", "", " And LM_LeaseInfo.RecCode = '" & txtRecoverer & "'")
            
            Module1.ChkTempTables "Tmp_Process", True
            Module1.ChkTempTables "Tmp_LMRecovery", True
            Module1.ChkTempTables "Tmp_Payments", True
            Module1.ChkTempTables "Tmp_Payments1", True
            Module1.ChkTempTables "Tmp_LastRental", True
            
             Module2.LMS_Recovery ls_Temp, dtpason.Value
            
            .ReportFileName = App.Path & Gs_ARRepoPath & "\RecovrReport.RPT"
            .DiscardSavedData = True
            .SelectionFormula = "({Tmp_LMRecovery.Accrued} - {Tmp_LMRecovery.PaidRntls}) >2 And ({Tmp_LMRecovery.AccrAmount} - {Tmp_Payments.PaidAmount}) >=200"
            .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
            .Formulas(1) = "ReportName = '" & Me.Caption & "'"
            .Formulas(2) = "UpToPeriod = '" & "As On : " & dtpason & "'"
            .Action = 1
            gc_dbcon.Execute ("Drop Table Tmp_LMRecovery;")
            gc_dbcon.Execute ("Drop Table Tmp_Payments;")
            gc_dbcon.Execute ("Drop Table Tmp_LastRental;")
            gc_dbcon.Execute ("Drop Table Tmp_LastPaid;")
            gc_dbcon.Execute ("Drop Table Tmp_Process;")
   Case UCase("Cred") ' Overdue Statement
            ls_Temp = ls_Temp + IIf(txtRecoverer = "", "", " And LM_LeaseInfo.CreditCode = '" & txtRecoverer & "'")
            Module1.ChkTempTables "Tmp_Process", True
            Module1.ChkTempTables "Tmp_LMRecovery", True
            Module1.ChkTempTables "Tmp_Payments", True
            Module1.ChkTempTables "Tmp_Payments1", True
            
            Module2.LMS_Recovery ls_Temp, dtpason.Value
            .ReportFileName = App.Path & Gs_ARRepoPath & "\RecReportoverallCredit.RPT"
            .DiscardSavedData = True
            .SelectionFormula = ""
            .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
            .Formulas(1) = "ReportName = 'Recovery Statement (' + '" & Me.Caption & "'+')'"
            .Formulas(2) = "UpToPeriod = '" & "As On : " & dtpason & "'"
            .Action = 1
            gc_dbcon.Execute ("Drop Table Tmp_LMRecovery;")
            gc_dbcon.Execute ("Drop Table Tmp_Payments;")
            gc_dbcon.Execute ("Drop Table Tmp_LastRental;")
            gc_dbcon.Execute ("Drop Table Tmp_LastPaid;")
            gc_dbcon.Execute ("Drop Table Tmp_Process;")
    Case "INDI"
            ls_Temp = ls_Temp + IIf(txtRecoverer = "", "", " And LM_LeaseInfo.RecCode = '" & txtRecoverer & "'")
            
            Module1.ChkTempTables "Tmp_Process", True
            Module1.ChkTempTables "Tmp_LMRecovery", True
            Module1.ChkTempTables "Tmp_Payments", True
            Module1.ChkTempTables "Tmp_Payments1", True
            Module1.ChkTempTables "Tmp_LastRental", True
            
            Module2.LMS_Recovery ls_Temp, dtpason.Value
            .ReportFileName = App.Path & Gs_ARRepoPath & "\LM_IndPerformance.RPT"
            .DiscardSavedData = True
            .SelectionFormula = ""
            .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
            .Formulas(1) = "ReportName = '" & Me.Caption & "'"
            .Formulas(2) = "UpToPeriod = '" & "As On : " & dtpason & "'"
            .Action = 1
            gc_dbcon.Execute ("Drop Table Tmp_LMRecovery;")
            gc_dbcon.Execute ("Drop Table Tmp_Payments;")
            gc_dbcon.Execute ("Drop Table Tmp_LastRental;")
            gc_dbcon.Execute ("Drop Table Tmp_LastPaid;")
            gc_dbcon.Execute ("Drop Table Tmp_Process;")
    Case "RECE"
            .ReportFileName = App.Path & Gs_ARRepoPath & "\LM_Payment.RPT"
            .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
            .Formulas(1) = "ReportName = '" & Me.Caption & "'"
            .Formulas(2) = "UptoPeriod = '" & "From: " & dtpason & " To: " & DTPTo & "'"
            .SelectionFormula = Replace(.SelectionFormula, "LM_Schedule", "LM_Payments")
            .SelectionFormula = .SelectionFormula & " And {LM_Payments.RelzDate} >= Date(" & Year(dtpason) & "," & Month(dtpason) & "," & Day(dtpason) & ") And  {LM_Payments.RelzDate} <= Date(" & Year(DTPTo) & "," & Month(DTPTo) & "," & Day(DTPTo) & ")"
            .Action = 1
    Case "BILL" ' Bills of Lease Facilities
            Module1.ChkTempTables "", False
            Module2.LMS_Bills ls_Temp, dtpason.Value
       
            .ReportFileName = App.Path & Gs_ARRepoPath & "\LM_Bills.RPT"
            .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
            .Formulas(1) = "ReportName = 'Bill For Lease Facility'"
            .Formulas(2) = "UpToPeriod = '" & "Outstanding Statement On : " & dtpason & "'"
            .SelectionFormula = ""
            .Action = 1
            gc_dbcon.Execute ("Drop Table Tmp_Bill;")
    Case "CUST"
            .SelectionFormula = Replace(.SelectionFormula, "LM_Schedule", "LM_Comments")
            .ReportFileName = App.Path & Gs_ARRepoPath & "\LM_CustComm.RPT"
            .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
            .Formulas(1) = "ReportName = 'Customer History Sheet'"
            .Formulas(2) = "ReportPeriod = '" & "Upto : " & dtpason & "'"
            .Action = 1
    Case "ASSE"
            .ReportFileName = App.Path & Gs_ARRepoPath & "\LM_AssetExp.RPT"
            .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
            .Formulas(1) = "ReportName = 'Asset Wize Exposure [ Based on Gross Lease Amount ] '"
            .Formulas(2) = "ReportPeriod = '" & "Upto : " & dtpason & "'"
            .SelectionFormula = "{LM_LeaseInfo.CompCode} = '" & Gs_compcode & "'"
            .SelectionFormula = .SelectionFormula + " And {LM_LeaseInfo.Agreemntdate} <= Date(" & Year(dtpason) & "," & Month(dtpason) & "," & Day(dtpason) & ") And {LM_LeaseInfo.Activestatus} >0"
            .Action = 1
    Case "SECT"
            .ReportFileName = App.Path & Gs_ARRepoPath & "\LM_SectorExp.RPT"
            .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
            .Formulas(1) = "ReportName = 'Sector Wize Exposure [ Based on Gross Lease Amount ]'"
            .Formulas(2) = "ReportPeriod = '" & "Upto : " & dtpason & "'"
            .SelectionFormula = "{LM_LeaseInfo.CompCode} = '" & Gs_compcode & "'"
            .SelectionFormula = .SelectionFormula + " And {LM_LeaseInfo.Agreemntdate} <= Date(" & Year(dtpason) & "," & Month(dtpason) & "," & Day(dtpason) & ") And {LM_LeaseInfo.Activestatus} >0"
            .Action = 1
    Case "ENTI"
            .ReportFileName = App.Path & Gs_ARRepoPath & "\LM_EntityExp.RPT"
            .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
            .Formulas(1) = "ReportName = 'Entity Wize Exposure [ Based on Gross Lease Amount ]'"
            .Formulas(2) = "ReportPeriod = '" & "Upto : " & dtpason & "'"
            .SelectionFormula = "{LM_LeaseInfo.CompCode}= '" & Gs_compcode & "'"
            .SelectionFormula = .SelectionFormula + " And  {LM_LeaseInfo.Agreemntdate} <= Date(" & Year(dtpason) & "," & Month(dtpason) & "," & Day(dtpason) & ") And {LM_LeaseInfo.Activestatus} >0"
            .Action = 1
    Case "STAT" ' Receivable Confirmation Statement
                 Call OverDuesConfirm(ls_Temp, dtpason.Value)
                .ReportFileName = App.Path & Gs_ARRepoPath & "\LM_Recvable.RPT"
                .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
                .Formulas(1) = "ReportName = '" & Me.Caption & "'"
                .Formulas(2) = "Period = '" & "As On : " & dtpason & "'"
                .SelectionFormula = ""
                .Action = 1
                 gc_dbcon.Execute ("Drop Table Tmp_SumOvrDue;")
    Case "LEAS" ' Rease Accrual Report
             Call LM_Accrual((DateAdd("m", 1, dtpason) - Day(dtpason)), .SelectionFormula, False)
            .ReportFileName = App.Path & Gs_ARRepoPath & "\LeaseAccrual.RPT"
            .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
            .WindowTitle = Me.Caption
            .Formulas(1) = "ReportName = '" & Me.Caption & "'"
            .Formulas(2) = "UptoPeriod = '" & "For the Month Of " & MonthName(Month(dtpason)) & " ," & Year(dtpason) & "'"
            .SelectionFormula = ""
            '.SelectionFormula = "{LM_LeaseInfo.CompCode}= '" & Gs_compcode & "'"
            '.SelectionFormula = .SelectionFormula + " And  Year({LM_Schedule.AccrualDate}) = " & Year(dtpason) & " And Month({LM_Schedule.AccrualDate})= " & Month(dtpason) & "  And {LM_LeaseInfo.Activestatus} =1 And {LM_LeaseInfo.ProvisionStat} = 1"
            .Action = 1
    Case "REGI" ' Rease Accrual Report
             Call OverDuesConfirm(ls_Temp, DTPTo.Value)
            .SelectionFormula = Replace(.SelectionFormula, "LM_Schedule", "LM_LeaseInfo")
            .ReportFileName = App.Path & Gs_ARRepoPath & "\Lm_LeaseInfo.RPT"
            .SelectionFormula = .SelectionFormula + " And {LM_LeaseInfo.SchdlDate} >= Date(" & Year(dtpason) & "," & Month(dtpason) & "," & Day(dtpason) & ")"
            .SelectionFormula = .SelectionFormula + " And {LM_LeaseInfo.SchdlDate} <= Date(" & Year(DTPTo) & "," & Month(DTPTo) & "," & Day(DTPTo) & ")"
            
            .SelectionFormula = .SelectionFormula + IIf(Chvdate.Value = 1, " And {LM_LeaseInfo.ActiveStatus} =1", "")
            .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
            .WindowTitle = "Lease Assets Register"
            .Formulas(1) = "ReportName = 'Lease Assets Register'"
            .Formulas(2) = "Period = '" & "From : " & dtpason & " &  To " & DTPTo & "'"
            .Action = 1
            
            gc_dbcon.Execute ("Drop Table Tmp_SumOvrDue;")
    Case "PORT" ' Rease Accrual Report
             Call OverDuesConfirm(ls_Temp, DTPTo.Value)
            .SelectionFormula = Replace(.SelectionFormula, "LM_Schedule", "LM_LeaseInfo")
            .ReportFileName = App.Path & Gs_ARRepoPath & "\Lm_LeasePortfolia.RPT"
            .SelectionFormula = .SelectionFormula + " And {LM_LeaseInfo.SchdlDate} >= Date(" & Year(dtpason) & "," & Month(dtpason) & "," & Day(dtpason) & ")"
            .SelectionFormula = .SelectionFormula + " And {LM_LeaseInfo.SchdlDate} <= Date(" & Year(DTPTo) & "," & Month(DTPTo) & "," & Day(DTPTo) & ")"
            
            .SelectionFormula = .SelectionFormula + IIf(Chvdate.Value = 1, " And {LM_LeaseInfo.ActiveStatus} =1", "")
            .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
            .WindowTitle = "Lease Assets Register"
            .Formulas(1) = "ReportName = 'Lease Assets Register'"
            .Formulas(2) = "Period = '" & "From : " & dtpason & " &  To " & DTPTo & "'"
            .Action = 1
            
            gc_dbcon.Execute ("Drop Table Tmp_SumOvrDue;")
         
            
    Case "POST" ' Posting To General Ledger
            Call LM_Accrual((DateAdd("m", 1, dtpason) - Day(dtpason)), .SelectionFormula, True)
            'Call LM_GLPosting((DateAdd("m", 1, dtpason) - Day(dtpason)), .SelectionFormula)
            '.ReportFileName = App.Path & Gs_ARRepoPath & "\LeaseAccrual.RPT"
            '.Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
            '.WindowTitle = Me.Caption
            '.Formulas(1) = "UptoPeriod = '" & "For the Month Of " & MonthName(Month(dtpason)) & "'"
            '.SelectionFormula = "{LM_LeaseInfo.CompCode}= '" & Gs_compcode & "'"
            '.SelectionFormula = .SelectionFormula + " And  Year({LM_Schedule.AccrualDate}) = " & Year(dtpason) & " And Month({LM_Schedule.AccrualDate})= " & Month(dtpason) & "  And {LM_LeaseInfo.Activestatus} >0"
            '.Action = 1
 
 'Inflow Reports
     Case "CASH"
                 Call CashBook(ls_Temp, dtpason.Value, "S")
                .ReportFileName = App.Path & Gs_STRepoPath & "\IF_Cash.RPT"
                .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
                .Formulas(1) = "ReportName =  '" & Me.Caption & "'"
                .Formulas(2) = "ReportPeriod = '" & "As On : " & dtpason & "'"
                .SelectionFormula = ""
                .WindowTitle = Me.Caption
                .Action = 1
                 gc_dbcon.Execute ("Drop Table Tmp_IF_Cash;")
     Case "BANK"
                 Call CashBook(ls_Temp, dtpason.Value, "Q")
                .ReportFileName = App.Path & Gs_STRepoPath & "\IF_Cash.RPT"
                .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
                .Formulas(1) = "ReportName = '" & Me.Caption & "'"
                .Formulas(2) = "ReportPeriod = '" & "As On : " & dtpason & "'"
                .SelectionFormula = ""
                .WindowTitle = Me.Caption
                .Action = 1
                 gc_dbcon.Execute ("Drop Table Tmp_IF_Cash;")
     Case "RIBA"
              Call Module2.RibaIrr2(dtpason, txtbranchcode, Val(0 & txtSLRReturn))
     Case Else
            
            .SelectionFormula = Replace(.SelectionFormula, "LM_Schedule", "IF_Inflow")
            .SelectionFormula = Replace(.SelectionFormula, "LeaseNo", "FacilityAcct")
            
            .ReportFileName = App.Path & Gs_STRepoPath & "\InflowReport2.RPT"
            .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
            .Formulas(1) = "ReportName = '" & Me.Caption & "'"
            .Formulas(2) = "ReportPeriod = '" & "From : " & dtpason & " To " & DTPTo & "'"
                'Data Filter checks
            If UCase(Left(Me.Caption, 4)) = "INFL" And Chvdate.Value = vbChecked Then 'based on value date
                 .SelectionFormula = .SelectionFormula & " And {IF_Inflow.Inflowstatus}=" & ln_InflowStatus & " And {IF_Inflow.ValueDate} >= date(" & dtpason.Year & "," & dtpason.Month & "," & dtpason.Day & ") And {IF_Inflow.ValueDate} <= date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ")"
            ElseIf UCase(Left(Me.Caption, 4)) = "REAL" Or UCase(Left(Me.Caption, 4)) = "BOUN" Then  'based on realized date
                .SelectionFormula = .SelectionFormula & " And ({IF_Inflow.Inflowstatus}=" & ln_InflowStatus & " " & IIf(ln_InflowStatus = 1, " Or {IF_Inflow.InflowType} = 'S')", ")") & " And {IF_Inflow.Relzdate} >= date(" & dtpason.Year & "," & dtpason.Month & "," & dtpason.Day & ") And {IF_Inflow.Relzdate} <= date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ")"
            ElseIf UCase(Left(Me.Caption, 4)) = "UNRE" Then  'based on unrealized date
                .SelectionFormula = .SelectionFormula & " And {IF_Inflow.Inflowstatus}=" & ln_InflowStatus & ""
                .SelectionFormula = .SelectionFormula & " And {IF_Inflow.InflowType} <> 'S'"
                .SelectionFormula = .SelectionFormula & " And {IF_Inflow.Inflowstatus}=" & ln_InflowStatus & " And {IF_Inflow.ValueDate} >= date(" & dtpason.Year & "," & dtpason.Month & "," & dtpason.Day & ") And {IF_Inflow.ValueDate} <= date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ")"
            ElseIf UCase(Left(Me.Caption, 4)) = "INFL" Then 'based on inflowdate
                 .SelectionFormula = .SelectionFormula & " And (({IF_Inflow.transdate} >= date(" & dtpason.Year & "," & dtpason.Month & "," & dtpason.Day & ") And {IF_Inflow.transdate} <= date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & "))"
                 .SelectionFormula = .SelectionFormula & " OR ({IF_Inflow.Relzdate} >= date(" & dtpason.Year & "," & dtpason.Month & "," & dtpason.Day & ") And {IF_Inflow.Relzdate} <= date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ")))"
            ElseIf UCase(Left(Me.Caption, 4)) = "DISH" Then 'based on inflowdate
                  Module2.LMS_OverDue ls_Temp, dtpason.Value
                 .ReportFileName = App.Path & Gs_STRepoPath & "\Lm_Bills.RPT"
                 .SelectionFormula = .SelectionFormula & " And (({IF_Inflow.Relzdate} >= date(" & dtpason.Year & "," & dtpason.Month & "," & dtpason.Day & ") And {IF_Inflow.Relzdate} <= date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ")))"
                 .SelectionFormula = .SelectionFormula & " And {IF_Inflow.Inflowstatus}= 2"
            End If
             .Action = 1
             If UCase(Left(Me.Caption, 4)) = "DISH" Then gc_dbcon.Execute "Drop Table Tmp_LMRecovery"
             
      End Select
 MDIForm1.StatusBar1.Panels(7).Text = ""
 End With
End Sub
Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCustNO
    Set PO_DESC = Text4
    Gs_SQL = "Select Customer.Customerno 'Customer No', Customer.CustomerName  'Customer Name' from Customer Inner Join Facilities On Customer.Compcode+Customer.BranchCode+Customer.CustomerNo = Facilities.Compcode+Facilities.Branchcode+Facilities.CustomerNo"
    Gs_FindFld = "CustomerName"
    Gs_OtherPara = " Where Customer.Compcode+Customer.BranchCode = '" & Gs_compcode + Gs_BranchCode & "' And Facilities.FacilityNo = '01'"
    Gs_OrderBy = "Order by Customer.CustomerNo,Customer.CustomerName"
    MyLookupOLDB.Caption = "Customers"
    MyLookupOLDB.Show 1

'    Set PO_AnyForm = Nothing
'    Set PO_AnyForm = Me
'    Set PO_CODE = txtCustNO
'    Set PO_DESC = Text4
'    GoTop PR_Customer
'    MyLookup.Caption = "Customer"
'    MyLookup.FillGrid PR_Customer, "CustomerNo", "CustomerName", txtCustNO.MaxLength
'    MyLookup.Show 1
    If Len(txtCustNO) > 0 Then TxtCustNo_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtleaseno
    Set PO_DESC = Text2

    PR_LMSInfo.Filter = "BranchCode = '" & Gs_BranchCode & "' And CustomerNo = '" & txtCustNO & "'"
    GoTop PR_LMSInfo
    MyLookup.Caption = "Lease Agreements"
    MyLookup.FillGrid PR_LMSInfo, "LeaseNo", "LeaseAmount", txtleaseno.MaxLength
    MyLookup.Show 1
    If Len(txtleaseno) > 0 Then txtleaseno_KeyDown vbKeyReturn, vbKeyShift
    PR_LMSInfo.Filter = adFilterNone
End Sub

Private Sub Command2_Click()

    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtRecoverer
    Set PO_DESC = Text2
    
    GoTop PR_Recoverer
    If UCase(Me.Caption) = UCase("Credit Officer Wise") Then
        PR_Recoverer.Filter = "RecTag = 'CRE'"
        MyLookup.Caption = "Credit Officer"
    Else
        MyLookup.Caption = "Recovery Officer"
    End If
    MyLookup.FillGrid PR_Recoverer, "RecCode", "RecName", txtRecoverer.MaxLength
    MyLookup.Show 1
    
    If UCase(Me.Caption) = UCase("credit officer wise") Then PR_Recoverer.Filter = adFilterNone
    
    If Len(txtRecoverer) > 0 Then txtrecoverer_KeyDown vbKeyReturn, vbKeyShift
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
Private Sub DTPAson_KeyDown(KeyCode As Integer, Shift As Integer)
   If Lastkey(KeyCode) Then
    If DTPTo.Enabled And DTPTo.Visible Then DTPTo.SetFocus
   End If
End Sub
Private Sub dtpto_KeyDown(KeyCode As Integer, Shift As Integer)
If Lastkey(KeyCode) Then
    If txtbranchcode.Enabled = True Then
        txtbranchcode.SetFocus
    Else
        txtCustNO.SetFocus
    End If
End If
End Sub

Private Sub Form_Activate()
    Call txtbranchcode_KeyDown(vbKeyReturn, vbKeyShift)
    txtRecoverer.Visible = IIf(UCase(Left(Me.Caption, 4)) = "RECO" Or UCase(Left(Me.Caption, 4)) = "OVER" Or UCase(Left(Me.Caption, 4)) = "INDI" Or UCase(Left(Me.Caption, 4)) = "CRED", True, False)
    Command2.Visible = IIf(UCase(Left(Me.Caption, 4)) = "RECO" Or UCase(Left(Me.Caption, 4)) = "OVER" Or UCase(Left(Me.Caption, 4)) = "INDI" Or UCase(Left(Me.Caption, 4)) = "CRED", True, False)
    Label5.Caption = IIf(UCase(Left(Me.Caption, 4)) = "RECO" Or UCase(Left(Me.Caption, 4)) = "OVER" Or UCase(Left(Me.Caption, 4)) = "INDI" Or UCase(Left(Me.Caption, 4)) = "CRED", "Recv.Off.", "")
    Chvdate.Visible = IIf(UCase(Left(Me.Caption, 4)) = "PROJ", True, False)
    txtleaseno.Visible = IIf(UCase(Left(Me.Caption, 4)) = "CUST", False, True)
    Label2(0).Visible = IIf(UCase(Left(Me.Caption, 4)) = "CUST", False, True)
    Command1.Visible = IIf(UCase(Left(Me.Caption, 4)) = "CUST", False, True)
    Label2(0).Caption = "Lease# :"
    dtpason.Enabled = IIf(Left(Me.Caption, 4) = "CLIE" Or Left(Me.Caption, 4) = "INTE", False, True)
    If UCase(Left(Me.Caption, 4)) = "INFL" Or UCase(Left(Me.Caption, 4)) = "REAL" Or UCase(Left(Me.Caption, 4)) = "BOUN" Or UCase(Left(Me.Caption, 4)) = "PROJ" Or UCase(Left(Me.Caption, 4)) = "RECE" Or UCase(Left(Me.Caption, 4)) = "UNRE" Or UCase(Left(Me.Caption, 4)) = "DISH" Or UCase(Left(Me.Caption, 4)) = "PORT" Then
        Label1.Caption = "From :"
        Label5.Visible = True
        DTPTo.Visible = True
        dtpason.Enabled = True
        Label5.Caption = "To:"
    ElseIf UCase(Me.Caption) = UCase("Register") Then
        DTPTo.Visible = True
        Label5.Caption = "To:"
        Chvdate.Visible = True
        Label5.Visible = True
    Else
        Label1.Caption = "As On:"
        Label5.Visible = IIf(UCase(Left(Me.Caption, 4)) = "RECO", True, False)
        Label5.Visible = IIf(UCase(Left(Me.Caption, 4)) = "OVER" Or UCase(Left(Me.Caption, 4)) = "INDI" Or UCase(Left(Me.Caption, 4)) = "CRED", True, False)
        DTPTo.Visible = False
    End If
    If UCase(Left(Me.Caption, 4)) = "RIBA" Then
        Label2(0).Visible = False
        Command1.Visible = False
        txtleaseno.Visible = False
    End If
   If UCase(Left(Me.Caption, 4)) = "INFL" Then
    ln_InflowStatus = 0
    Label2(0).Caption = "Account# :"
   ElseIf UCase(Left(Me.Caption, 4)) = "REAL" Then
            ln_InflowStatus = 1
   ElseIf UCase(Left(Me.Caption, 4)) = "BOUN" Then
            ln_InflowStatus = 2
   ElseIf UCase(Left(Me.Caption, 4)) = "UNRE" Then
            ln_InflowStatus = 0
   End If
  Label2(0).Caption = "Account# :"
   
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    PR_Recoverer.Open "Select * From LM_Recoverer Order By RecCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
    PR_Branch.Open "Select * From SysBranch Order By BranchCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
    pr_Customer.Open "Select Customer.*,Customer.BranchCode+Customer.CustomerNo As FindFld from Customer Inner Join Facilities On Customer.Compcode+Customer.BranchCode+Customer.CustomerNo = Facilities.Compcode+Facilities.BranchCode+Facilities.CustomerNo Where Customer.Compcode+Customer.BranchCode = '" & Gs_compcode + Gs_BranchCode & "' And Facilities.FacilityNo = '01' Order By Customer.CustomerNo,Customer.BranchCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
    PR_LMSInfo.Open "Select *,BranchCode+CustomerNo+LeaseNo As FindFld from LM_LeaseInfo where compcode ='" & Gs_compcode & "' Order by BranchCode,CustomerNo,LeaseNo", gc_dbcon, adOpenStatic, adLockOptimistic, 1
    txtbranchcode = Gs_BranchCode
    DTPTo.Value = Date
    dtpason.Value = Date
    Screen.MousePointer = Default
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PR_Recoverer.Close
   PR_Branch.Close
   pr_Customer.Close
   PR_LMSInfo.Close
End Sub

Private Sub txtbranchcode_KeyDown(KeyCode As Integer, Shift As Integer)
   If Lastkey(KeyCode) And txtbranchcode <> "" Then
     txtbranchcode = DoPad(txtbranchcode, txtbranchcode.MaxLength)
 
     If Not MySeek(txtbranchcode, "Branchcode", PR_Branch) Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtbranchcode.SetFocus
     Else
        Text1 = PR_Branch("BranchDesc")
        If txtCustNO.Enabled And txtCustNO.Visible Then txtCustNO.SetFocus
     End If
   ElseIf KeyCode = vbKeyF12 Then
        Call Command5_Click
  End If
End Sub

Private Sub TxtCustNo_KeyDown(KeyCode As Integer, Shift As Integer)
 If Lastkey(KeyCode) And Trim(txtCustNO) <> "" Then
   txtCustNO = DoPad(txtCustNO, txtCustNO.MaxLength)
   If MySeek(txtbranchcode + txtCustNO, "FindFld", pr_Customer) Then
      Text4 = pr_Customer("CustomerName") & ""
      If txtleaseno.Visible = False Then
            txtCustNO.SetFocus
      Else
            txtleaseno.SetFocus
      End If
   Else
      Call SetErr(Gs_RecNFMsg, vbCritical)
      txtCustNO.SetFocus
   End If
 ElseIf KeyCode = vbKeyF12 Then
        Call cmdLookup_Click
 End If
End Sub

Private Sub txtleaseno_KeyDown(KeyCode As Integer, Shift As Integer)
'If LastKey(KeyCode) Then cmdGenerate.SetFocus
If Lastkey(KeyCode) And Trim(txtleaseno) <> "" Then
   txtleaseno = DoPad(txtleaseno, txtleaseno.MaxLength)
   If MySeek(txtbranchcode + txtCustNO + txtleaseno, "FindFld", PR_LMSInfo) Then
       cmdGenerate.SetFocus
   Else
      Call SetErr("Record not found", vbCritical)
      txtleaseno.SetFocus
   End If
ElseIf KeyCode = vbKeyF12 Then
        Call Command1_Click
End If
End Sub

Private Sub txtrecoverer_KeyDown(KeyCode As Integer, Shift As Integer)
 If Lastkey(KeyCode) And txtRecoverer <> "" Then
    If UCase(Me.Caption) = UCase("credit officer wise") Then PR_Recoverer.Filter = "RecTag = 'CRE'"
        txtRecoverer = DoPad(txtRecoverer, txtRecoverer.MaxLength)
        If Not MySeek(txtRecoverer, "RecCode", PR_Recoverer) Then
            Call SetErr("Record not found", vbCritical)
            txtRecoverer.SetFocus
        Else
            Text2 = PR_Recoverer("RecName") & ""
            txtCustNO.SetFocus
        End If
   If UCase(Me.Caption) = UCase("credit officer wise") Then PR_Recoverer.Filter = adFilterNone
 End If
End Sub

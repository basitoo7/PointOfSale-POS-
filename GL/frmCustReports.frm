VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCustReports 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customized Reports Format"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   Icon            =   "frmCustReports.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3750
      Left            =   30
      TabIndex        =   1
      Top             =   -90
      Width           =   6555
      Begin VB.Frame Frame2 
         Caption         =   "Modes"
         ForeColor       =   &H00000080&
         Height          =   3045
         Left            =   3660
         TabIndex        =   13
         Top             =   150
         Width           =   2835
         Begin VB.OptionButton Option4 
            BackColor       =   &H80000004&
            Caption         =   "Send &e-mail"
            Height          =   255
            Left            =   180
            TabIndex        =   17
            Top             =   2625
            Width           =   1770
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H80000004&
            Caption         =   "Print Pre&view"
            Height          =   255
            Left            =   165
            TabIndex        =   16
            Top             =   480
            Value           =   -1  'True
            Width           =   1425
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H80000004&
            Caption         =   "Print To &File"
            Height          =   255
            Left            =   180
            TabIndex        =   15
            Top             =   1890
            Width           =   1815
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H80000004&
            Caption         =   "&Print"
            Height          =   255
            Left            =   180
            TabIndex        =   14
            Top             =   1155
            Width           =   1815
         End
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   5640
         TabIndex        =   11
         Text            =   "1"
         Top             =   3315
         Width           =   405
      End
      Begin VB.Frame Frame7 
         Caption         =   "Parameters"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   3045
         Left            =   75
         TabIndex        =   5
         Top             =   150
         Width           =   3525
         Begin VB.CommandButton Command3 
            Height          =   315
            Left            =   2340
            Picture         =   "frmCustReports.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1020
            Width           =   315
         End
         Begin VB.TextBox txtRptCode 
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "XXX"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   1590
            MaxLength       =   2
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1020
            Width           =   735
         End
         Begin MSComCtl2.DTPicker dtpto 
            Height          =   315
            Left            =   1590
            TabIndex        =   8
            Top             =   1590
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
            Format          =   65536001
            CurrentDate     =   37293
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Report Code :"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   480
            TabIndex        =   10
            Top             =   1065
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Upto :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1080
            TabIndex        =   9
            Top             =   1635
            Width           =   420
         End
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   2010
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   615
         Visible         =   0   'False
         Width           =   255
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
         Left            =   645
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Top             =   3285
         Width           =   1037
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   405
         Left            =   1740
         TabIndex        =   2
         Top             =   3285
         Width           =   1037
      End
      Begin Crystal.CrystalReport rptRpts 
         Left            =   210
         Top             =   1155
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
      Begin VB.Label Label3 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Number Of Copies:-"
         Height          =   255
         Left            =   4110
         TabIndex        =   12
         Top             =   3345
         Width           =   1935
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3720
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "Description"
            TextSave        =   "Description"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCustReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Pb_BlnkVchr As Boolean
Dim Ls_CalcBase As String

Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_GrpDet As Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
On Error GoTo LocalErr
Dim lb_found As Boolean

  If Left(Me.Caption, 1) = "G" Then
  Call ChkTempTables("Tmp_CustRpt", True)
  Call GenCustRpts
  lb_found = MySeek(txtRptCode.Text, "ReportCode", PR_GrpDet)
  If lb_found Then
  With rptRpts
       .ReportFileName = App.Path & Gs_GlRepoPath & "\CustReport.RPT"
       .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
       .Formulas(1) = "ReportName = '" & UCase(PR_GrpDet.Fields("RptDescrip")) & "'"
       .Formulas(2) = "Head1 = '" & PR_GrpDet.Fields("RptHeader") & " " & dtpto.Value & "'"
       .Formulas(3) = "Head2 = '" & PR_GrpDet.Fields("RptSubHeader") & "'"
       .Action = 1
   End With
 Else
    Call SetErr("Please Select Report Type", vbCritical)
    txtRptCode.SetFocus
 End If
   gc_dbcon.Execute ("Drop Table Tmp_CustRpt;")
  Else
    With rptRpts
        .ReportFileName = App.Path & Gs_GlRepoPath & "\GroupDetal.RPT"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = '" & IIf(Val(txtRptCode) = 0, "All", txtDesc) & "'"
        .SelectionFormula = "{GlGroupDetl.CompCode} = '" & Gs_compcode & "'"
        If Len(Trim(txtRptCode)) > 0 Then .SelectionFormula = .SelectionFormula & " AND {GlGroupDetl.ReportCode} = '" & txtRptCode & "'"
        
        .Action = 1
    End With
  End If

Exit Sub

LocalErr:
'gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub Command3_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtRptCode
    Set PO_DESC = txtDesc
    If Left(Me.Caption, 1) = "G" Then
       PR_GrpDet.Filter = "ReportBase = 'O'"
    End If
    GoTop PR_GrpDet
    MyLookup.Caption = "Customized Reports"
    MyLookup.FillGrid PR_GrpDet, "ReportCode", "RptDescrip", 5
    MyLookup.Show 1
    
    If Len(txtRptCode) > 0 Then txtRptCode_KeyDown vbKeyReturn, vbKeyShift
    PR_GrpDet.Filter = adFilterNone
End Sub

Private Sub Form_Load()
  
  Set PR_GrpDet = New Recordset
  PR_GrpDet.Open "Select * from GlRpts_Ref where CompCode ='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  Pb_BlnkVchr = IIf(PR_GrpDet.EOF, True, False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_GrpDet.Close
End Sub

Private Sub txtRptCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 
 If Lastkey(KeyCode) And txtRptCode.Text <> "" Then
         txtRptCode = UCase(txtRptCode)
         lb_found = MySeek(txtRptCode.Text, "ReportCode", PR_GrpDet)
   
         If Not lb_found Then
            Call SetErr(Gs_RecNFMsg, vbCritical)
            txtRptCode.SetFocus
         Else
            StatusBar1.Panels(2) = PR_GrpDet.Fields("RptDescrip")
            Ls_CalcBase = PR_GrpDet.Fields("CalcBase") & ""
            cmdGenerate.SetFocus
         End If
  ElseIf Lastkey(KeyCode) And txtRptCode.Text = "" Then
         cmdGenerate.SetFocus
  ElseIf KeyCode = vbKeyF12 Then
     Command3_Click
  End If
End Sub

Private Sub GenCustRpts()
Dim ls_sql As String
Dim ln_cnt As Integer
Dim Ln_SumNo As Integer
Dim ln_Amount As Double
Dim ln_sum(1 To 9, 1) As Double
Dim cntsql As New ADODB.Command
Dim Tmp_GlNet As New Recordset
Dim PR_Rpts As New Recordset
Dim ln_Count As Integer
ln_Count = 0
cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText

ls_sql = "Select GlGroupDetl.*,000000000 as Amount Into Tmp_CustRpt From GlGroupDetl Where GlGroupDetl.Compcode = '" & Gs_compcode & "' and GlGroupDetl.ReportCode ='" & txtRptCode.Text & "' Order By GlGroupDetl.GroupCode "
gc_dbcon.Execute (ls_sql)

PR_Rpts.Open "Select * From Tmp_CustRpt", gc_dbcon, adOpenDynamic, adLockOptimistic

ls_sql = "Select Gl_Trans.AccountNo,sum(Gl_Trans.dr_Amount) As DrAmt,sum(Gl_Trans.cr_amount) As CrAmt,GL_RptRouting.Acct_Status,Gl_RptRouting.GroupCode,Gl_Detail.Acct_desc As AccntDesc"
ls_sql = ls_sql & " From gl_Trans INNER JOIN Gl_RptRouting ON gl_Trans.AccountNo = Gl_RptRouting.AccountNo and gl_Trans.Compcode = Gl_RptRouting.Compcode"
ls_sql = ls_sql & " LEFT OUTER JOIN Gl_Detail ON gl_Trans.AccountNo = Gl_Detail.AccountNo and gl_Trans.Compcode = Gl_Detail.Compcode"
ls_sql = ls_sql & " Where value_Date between '" & Format(DateValue(Gs_Fnperiod), "YYYY/MM/DD") & "' and '" & Format(dtpto.Value, "YYYY/MM/DD") & "' "
ls_sql = ls_sql & " and Gl_Trans.compcode = '" & Gs_compcode & "' and Gl_RptRouting.ReportCode = '" & txtRptCode.Text & "'"
ls_sql = ls_sql & " GROUP BY gl_Trans.AccountNo,GL_RptRouting.Acct_Status,Gl_RptRouting.GroupCode,Gl_Detail.Acct_desc order by 1"
Set Tmp_GlNet = gc_dbcon.Execute(ls_sql)

GoTop Tmp_GlNet
Do While Not Tmp_GlNet.EOF
ln_Amount = 0
    If Tmp_GlNet.Fields("Acct_Status") = "N" Then
       If Ls_CalcBase = "D" Then
          ln_Amount = Tmp_GlNet.Fields("DrAmt") - Tmp_GlNet.Fields("CrAmt")
       Else
          ln_Amount = Tmp_GlNet.Fields("CrAmt") - Tmp_GlNet.Fields("DrAmt")
       End If
    ElseIf Tmp_GlNet.Fields("Acct_Status") = "D" Then
        ln_Amount = (Tmp_GlNet.Fields("DrAmt").Value * IIf(Ls_CalcBase = "D", 1, -1))
    ElseIf Tmp_GlNet.Fields("Acct_Status") = "C" Then
        ln_Amount = (Tmp_GlNet.Fields("CrAmt").Value * IIf(Ls_CalcBase = "C", 1, -1))
    End If
        
    If MySeek(Tmp_GlNet.Fields("GroupCode").Value, "GroupCode", PR_Rpts) Then
      If PR_Rpts.Fields("SumGroup") = 1 Then
        cntsql.CommandText = "Update Tmp_CustRpt Set Amount = Amount + " & ln_Amount & " Where Tmp_CustRpt.Groupcode = '" & Tmp_GlNet.Fields("GroupCode") & "' And Tmp_CustRpt.TypeGroup = 'P'"
      Else
        gc_dbcon.BeginTrans
        cntsql.CommandText = "Update Tmp_CustRpt Set Amount = Amount + " & ln_Amount & " Where Tmp_CustRpt.Groupcode = '" & Tmp_GlNet.Fields("GroupCode") & "' And Tmp_CustRpt.TypeGroup = 'P'"
        cntsql.Execute
        gc_dbcon.CommitTrans
        cntsql.CommandText = "INSERT INTO Tmp_CustRpt (CompCode,ReportCode,GroupCode,GroupDesc,TypeGroup,Amount) Values ('" & Gs_compcode & "','" & txtRptCode.Text & "','" & Tmp_GlNet.Fields("GroupCode") & "','" & Tmp_GlNet.Fields("AccntDesc") & "','A', " & ln_Amount & ")"
      End If
      cntsql.Execute
    End If
    Tmp_GlNet.MoveNext
    If Tmp_GlNet.EOF Then Exit Do
Loop
Tmp_GlNet.Close
PR_Rpts.Requery
PR_Rpts.Filter = "TypeGroup <> 'A'"
GoTop PR_Rpts

Do While Not PR_Rpts.EOF
   
   If Left(PR_Rpts.Fields("TypeGroup"), 1) = "P" Then
      For ln_cnt = 1 To 9
         ln_sum(ln_cnt, 1) = ln_sum(ln_cnt, 1) + PR_Rpts.Fields("Amount")
      Next
      'Pr_GlNotes.Fields(ls_FldID) = Pr_GlNotes.Fields(ls_FldID) * IIf(Pr_GlNotes.Fields("GrpType") = "D", -1, 1)
   ElseIf Left(PR_Rpts.Fields("TypeGroup"), 1) = "Z" Then
          Ln_SumNo = Val(Right(PR_Rpts.Fields("TypeGroup"), 1))
          ln_sum(Ln_SumNo, 1) = 0
   ElseIf Left(PR_Rpts.Fields("TypeGroup"), 1) = "X" Then
          Ln_SumNo = Val(Right(PR_Rpts.Fields("TypeGroup"), 1))
          PR_Rpts.Fields("Amount") = ln_sum(Ln_SumNo, 1)
          '* IIf(ln_sum(ln_SumNo, 3) > (ln_sum(ln_SumNo, 2) * -1), 1, IIf(ln_sum(ln_SumNo, 3) <= 0, -1, 1)))
          ln_sum(Ln_SumNo, 1) = 0
          PR_Rpts.Update
   End If
   PR_Rpts.MoveNext
   If PR_Rpts.EOF Then Exit Do
Loop
PR_Rpts.Close
End Sub

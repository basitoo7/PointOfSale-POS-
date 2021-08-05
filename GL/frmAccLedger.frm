VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAccLedger 
   Caption         =   "General Ledger"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
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
   Icon            =   "frmAccLedger.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3465
      Left            =   30
      TabIndex        =   1
      Top             =   -90
      Width           =   5340
      Begin VB.CheckBox Check1 
         Caption         =   "With Under Clear Entries Option"
         Height          =   330
         Left            =   120
         TabIndex        =   22
         Top             =   2940
         Width           =   2670
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
         Height          =   405
         Left            =   3135
         MaskColor       =   &H00000000&
         TabIndex        =   19
         Top             =   2925
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   405
         Left            =   4230
         TabIndex        =   18
         Top             =   2925
         Width           =   1035
      End
      Begin VB.Frame Frame7 
         Caption         =   "Account Range"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1410
         Left            =   75
         TabIndex        =   11
         Top             =   1500
         Width           =   5220
         Begin VB.CommandButton Command1 
            Caption         =   "&Selective Accounts"
            Height          =   330
            Left            =   2835
            TabIndex        =   23
            Top             =   945
            Width           =   2295
         End
         Begin VB.TextBox txttoDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   3150
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   540
            Width           =   1995
         End
         Begin VB.TextBox txtFromDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   3150
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   180
            Width           =   1995
         End
         Begin VB.CommandButton Command4 
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
            Left            =   2820
            Picture         =   "frmAccLedger.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   540
            Width           =   315
         End
         Begin VB.TextBox txtTo 
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
            Left            =   1380
            MaxLength       =   50
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Account No"
            Top             =   555
            Width           =   1425
         End
         Begin VB.CommandButton Command3 
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
            Left            =   2820
            Picture         =   "frmAccLedger.frx":047C
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   180
            Width           =   315
         End
         Begin VB.TextBox txtFrom 
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
            Left            =   1380
            MaxLength       =   50
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Account No"
            Top             =   195
            Width           =   1410
         End
         Begin VB.Label txtselectiveaccount 
            Height          =   300
            Left            =   120
            TabIndex        =   24
            Top             =   975
            Visible         =   0   'False
            Width           =   2520
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "To Account # :"
            Height          =   210
            Left            =   255
            TabIndex        =   17
            Top             =   585
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "From Account # :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   16
            Top             =   210
            Width           =   2025
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Period"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1350
         Left            =   90
         TabIndex        =   2
         Top             =   165
         Width           =   5205
         Begin VB.CommandButton Command2 
            Caption         =   "Update Balance For Lookup"
            Height          =   345
            Left            =   2955
            TabIndex        =   25
            Top             =   945
            Width           =   2160
         End
         Begin VB.CommandButton Command5 
            Height          =   315
            Left            =   1995
            Picture         =   "frmAccLedger.frx":05EE
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
            Left            =   2325
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   210
            Width           =   2790
         End
         Begin MSComCtl2.DTPicker dtpfrom 
            Height          =   315
            Left            =   1380
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
            Format          =   131137537
            CurrentDate     =   37293
         End
         Begin MSComCtl2.DTPicker dtpto 
            Height          =   315
            Left            =   1380
            TabIndex        =   6
            Top             =   945
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
            Format          =   131137537
            CurrentDate     =   37293
         End
         Begin MSMask.MaskEdBox txtbranchcode 
            Height          =   315
            Left            =   1380
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Default Currency"
            Top             =   210
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "From Date :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   495
            TabIndex        =   10
            Top             =   600
            Width           =   825
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "To Date :"
            Height          =   210
            Left            =   675
            TabIndex        =   9
            Top             =   975
            Width           =   645
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Branch Code :"
            Height          =   210
            Left            =   285
            TabIndex        =   8
            Top             =   255
            Width           =   1035
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3360
      Width           =   5415
      _ExtentX        =   9551
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
            Object.Width           =   88194
            MinWidth        =   88194
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
   Begin Crystal.CrystalReport rptLedger 
      Left            =   30
      Top             =   2895
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
End
Attribute VB_Name = "frmAccLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pb_BlnkVchr As Boolean
Dim Mode As String
Dim PR_GlDetail As Recordset
Dim PR_Branch As New Recordset
Dim lb_found As Boolean
Public PO_DESC As Object
Public PO_CODE As Object
Dim ls_sql As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
'On Error GoTo LocalErr
Dim ls_AcctRange As String
Dim ls_vchrtype  As String
Dim ls_ActSQL    As String
Dim ls_opt    As String
Dim ls_branchdesc As String

Dim ls_Branch As String

If txtbranchname = "" Then
    Call txtBranchCode_KeyDown(vbKeyReturn, vbKeyShift)
End If

If txtbranchname <> "" Then
    ls_branchdesc = "-(" + txtbranchname + ")"
Else
    ls_branchdesc = ""
End If

        MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
        Select Case Me.Caption
        Case "Closing Balances"
            Call ChkTempTables("Tmp_GlClosings", True)
    
            If txtFrom.Text = "" Then
            txtFrom.Text = DoPad(txtFrom.Text, Gn_TotLen)
            End If
    
            If txtTo.Text = "" Then
            txtTo.Text = DoPad(txtTo.Text, Gn_TotLen, "9")
            End If

            ls_Branch = IIf(LTrim(RTrim(txtbranchcode)) = "", "", " And Gl_Trans.BranchCode = '" & txtbranchcode & "'")
            ls_ActSQL = "SELECT dbo.Gl_Trans.accountno, dbo.Gl_Detail.Acct_Desc, SUM(Case when dbo.Gl_Detail.Acct_Base='P' And gl_Trans.Value_Date < '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' Then 0 Else dbo.Gl_Trans.Dr_Amount End) AS Dr_Amount, SUM(Case when dbo.Gl_Detail.Acct_Base='P' And gl_Trans.Value_Date < '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' Then 0 Else dbo.Gl_Trans.Cr_Amount End) AS Cr_Amount,"
            ls_ActSQL = ls_ActSQL & " SUM(Case when dbo.Gl_Detail.Acct_Base='P' And gl_Trans.Value_Date < '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' Then 0 else dbo.Gl_Trans.Dr_Amount - dbo.Gl_Trans.Cr_Amount End) AS Balance INTO Tmp_GlClosings"
            ls_ActSQL = ls_ActSQL & " FROM dbo.Gl_Trans INNER JOIN"
            ls_ActSQL = ls_ActSQL & " dbo.Gl_Detail ON dbo.Gl_Trans.compcode = dbo.Gl_Detail.compcode  AND"
            ls_ActSQL = ls_ActSQL & " dbo.Gl_Trans.AccountNo = dbo.GL_Detail.AccountNo"
            ls_ActSQL = ls_ActSQL & " WHERE  Gl_Trans.CompCode = '" & Gs_compcode & "'   " & ls_Branch & " AND Gl_Trans.AccountNo >= '" & Trim(txtFrom) & "' AND Gl_Trans.AccountNo <= '" & Trim(txtTo) & "'"
            ls_ActSQL = ls_ActSQL & " AND  GL_Trans.Value_Date <= '" & Format(DTPTo, "YYYY/MM/DD") & "'"
            ls_ActSQL = ls_ActSQL & " GROUP BY dbo.Gl_Trans.accountno, dbo.Gl_Detail.Acct_Desc"
            gc_dbcon.Execute (ls_ActSQL)
    
            'GL_Trans.Value_Date >= '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' AND
    
            With rptLedger
                .ReportFileName = App.Path & Gs_GlRepoPath & "\ClBalance.RPT"
                .WindowTitle = Me.Caption
                .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
                .Formulas(1) = "ReportName = '" & Me.Caption & "'"
                .Formulas(2) = "Period = '" & "As On " + str(DTPTo) & "'"
                .Connect = "DNS=Censoft;UID=Sa"
                .Action = 1
            End With
    
                gc_dbcon.Execute ("DROP Table Tmp_GlClosings;")

        Case Else
                
                Dim PR_OpnBal As Recordset
                Dim cntsql As New ADODB.Command
                Call ChkTempTables("Tmp_GlTrans", True)
                cntsql.ActiveConnection = gc_dbcon
                cntsql.CommandType = adCmdText
                    
                ls_opt = IIf(Me.Caption = "General Ledger", " and gl_trans.pflag = 'P'", "")
                
                    
                    
                    ls_Branch = IIf(LTrim(RTrim(txtbranchcode)) = "", "", " And Gl_Trans.BranchCode = '" & txtbranchcode & "'")
                    ls_ActSQL = "SELECT Gl_Trans.compcode, Gl_Trans.BranchCode,Gl_Trans. accountno,Gl_Trans. SerialNo, Gl_Trans.AcctName, Gl_Trans.Acct_Nirration, Gl_Trans.Value_Date, Gl_Trans.Voucher_No,Gl_Trans.VchrType, Gl_Trans.ExchgRate,Gl_Trans. Dr_Amount, Gl_Trans.Cr_Amount, "
                    ls_ActSQL = ls_ActSQL & "  Gl_Trans.AddDate , Gl_Trans.Addtime, Gl_Trans.userid, Gl_Trans.pflag , Gl_Detail.Acct_Desc AS Title,"
                    ls_ActSQL = ls_ActSQL & " Gl_ref.instrumentno AS Instrumentno,Gl_Ref.Vchr_Remarks AS VchrRemarks,cast(0 As decimal(20,2)) As OB  INTO Tmp_GlTrans "
                    ls_ActSQL = ls_ActSQL & " FROM Gl_Trans LEFT OUTER JOIN"
                    ls_ActSQL = ls_ActSQL & " Gl_Ref ON Gl_Trans.compcode = Gl_Ref.CompCode AND Gl_Trans.Branchcode = Gl_Ref.BranchCode And gl_Trans.Value_Date = gl_ref.value_date and "
                    ls_ActSQL = ls_ActSQL & " Gl_Trans.Voucher_No = Gl_Ref.Voucher_no AND"
                    ls_ActSQL = ls_ActSQL & " Gl_Trans.VchrType = Gl_Ref.VchrType LEFT OUTER JOIN"
                    ls_ActSQL = ls_ActSQL & " Gl_Detail ON Gl_Trans.compcode = Gl_Detail.compcode AND"
                    ls_ActSQL = ls_ActSQL & " Gl_Trans.AccountNo = Gl_Detail.AccountNo"
                
                    If txtFrom.Text = "" Then
                        txtFrom.Text = DoPad(txtFrom.Text, Gn_TotLen)
                    End If
                    
                    If txtTo.Text = "" Then
                        txtTo.Text = DoPad(txtTo.Text, Gn_TotLen, "9")
                    End If
                    
                    If DateValue(Gs_Fnperiod) = dtpfrom Then
                        ls_vchrtype = "  AND VchrType = '" & Gs_ObVchrType & "'"
                    End If
                    If txtFrom = "Selective" Then
                        ls_AcctRange = " WHERE  Gl_Trans.CompCode = '" & Gs_compcode & "' " & ls_Branch & " AND Gl_Trans.AccountNo in (" & txtselectiveaccount & ")"
                    Else
                        ls_AcctRange = " WHERE  Gl_Trans.CompCode = '" & Gs_compcode & "' " & ls_Branch & " AND Gl_Trans.AccountNo >= '" & Trim(txtFrom) & "' AND Gl_Trans.AccountNo <= '" & Trim(txtTo) & "'"
                    End If
                    gc_dbcon.BeginTrans
                    'If Format(dtpfrom, "YYYY/MM/DD") <= Format(Gs_Fnperiod, "YYYY/MM/DD") Then
                         ls_opbsql = "SELECT Gl_Trans.CompCode,gl_Trans.BranchCode, Gl_Trans.AccountNo,'0OB','" & Format(dtpfrom, "YYYY/MM/DD") & "', SUM(Case When gl_Detail.Acct_Base = 'P' and gl_Trans.Value_Date < '" & Format(Gs_Fnperiod, "YYYY/MM/DD") & "' Then 0 Else Gl_Trans.Dr_Amount - Gl_Trans.Cr_Amount End) AS OB,0 as Dr_Amount,0 as Cr_Amount,0 As SerialNo,0 As Voucher_No  from Gl_Trans Inner Join Gl_Detail ON Gl_Trans.Compcode+Gl_Trans.AccountNo =Gl_Detail.Compcode+Gl_Detail.AccountNo  " & ls_AcctRange & "  and (gl_Trans.Value_Date < '" & Format(dtpfrom, "YYYY/MM/DD") & "' or gl_Trans.VchrType='0OB')  GROUP BY gl_Trans.CompCode,gl_Trans.BranchCode, gl_Trans.AccountNo"
                   'Else
                    '     ls_opbsql = "SELECT Gl_Trans.CompCode,gl_Trans.BranchCode, Gl_Trans.AccountNo,'0OB','" & Format(dtpfrom, "YYYY/MM/DD") & "', SUM(Gl_Trans.Dr_Amount - Gl_Trans.Cr_Amount) AS OB,0 as Dr_Amount,0 as Cr_Amount,0 As SerialNo,0 As Voucher_No from Gl_Trans " & ls_AcctRange & "  and gl_Trans.Value_Date >= '" & Format(DateValue(Gs_Fnperiod), "YYYY/MM/DD") & "' AND  gl_Trans.Value_Date < '" & Format(dtpfrom.Value, "YYYY/MM/DD") & "'" & ls_VchrType & "  GROUP BY gl_Trans.CompCode,gl_Trans.BranchCode, gl_Trans.AccountNo"
                    'End If
                    If Check1.Value = 1 Then
                        gc_dbcon.Execute (ls_ActSQL & ls_AcctRange & " AND Gl_Trans.Value_Date >= '" & Format(dtpfrom, "YYYY/MM/DD") & "' AND Gl_Trans.Value_Date <= '" & Format(Gs_FnEndPeriod, "YYYY/MM/DD") & "' And gl_Trans.VchrType <> '0OB'")
                    Else
                        gc_dbcon.Execute (ls_ActSQL & ls_AcctRange & " AND Gl_Trans.Value_Date >= '" & Format(dtpfrom, "YYYY/MM/DD") & "' AND Gl_Trans.Value_Date <= '" & Format(DTPTo.Value, "YYYY/MM/DD") & "' And gl_Trans.VchrType <> '0OB'")
                    End If
                    gc_dbcon.CommitTrans
                    
                    
                    gc_dbcon.Execute "Insert Into Tmp_GlTrans(Compcode, BranchCode,AccountNo,VchrType,Value_Date,OB,Dr_Amount,Cr_Amount,SerialNo,Voucher_No)" & ls_opbsql
                    
                    With rptLedger
                        
                        If Check1.Value = 1 Then
                          .ReportFileName = App.Path & Gs_GlRepoPath & "\LedgerOption2.RPT"
                          .Formulas(4) = "FromDate = cdate(" & Format(DTPTo, "yyyy,MM,dd") & ")"
                        Else
                          .ReportFileName = App.Path & Gs_GlRepoPath & "\Ledger.RPT"
                          .Formulas(4) = ""
                        End If
                 
                        .WindowTitle = Me.Caption
                        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
                        .Formulas(1) = "ReportName = '" & "General Ledger" + ls_branchdesc & "'"
                        .Formulas(2) = "Period = '" & "From " + str(dtpfrom) + " To " + str(DTPTo) & "' "
                        '.Formulas(3) = "PrintBy = '" & Gc_UserName & "'"
                        
                        .Connect = "DNS=Censoft;UID=Sa"
                        .Action = 1
                        .PageZoom 120
                    End With
                   ' gc_dbcon.Execute ("DROP TABLE Tmp_GlTrans;")
        End Select
        MDIForm1.StatusBar1.Panels(7).Text = ""

Exit Sub

LocalErr:
MDIForm1.StatusBar1.Panels(7).Text = ""
MsgBox Err.Description
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtselectiveaccount
    txtFrom = "Selective"
    txtTo = "Selective"
    txtFromDesc = "Selective Accounts"
    txttoDesc = "Selective Accounts"
    
    Set PO_DESC = txtFromDesc
    Gs_SQL = "Select Accountno 'Account No', Acct_Desc  'Description',BalAmount AS Amount, DRCR from gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_Subon = True
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
    Gs_OrderBy = "Order by Acct_Desc,AccountNo"
    MyLookupOLDBSelective.txtsearchbase.Clear
    MyLookupOLDBSelective.txtsearchbase.AddItem "Accountno"
    MyLookupOLDBSelective.txtsearchbase.AddItem "Acct_Desc"
    MyLookupOLDBSelective.txtsearchbase.ListIndex = 1
    MyLookupOLDBSelective.Caption = "Account Nos."
    MyLookupOLDBSelective.Show 1
    
    'If Len(txtFrom) > 0 Then txtFrom_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Command2_Click()
ls_sql = "Update Gl_Detail SET Gl_Detail.BalAmount = AccountBalance.Amount, Gl_Detail.DRCR = AccountBalance.DRCR FROM Gl_Detail INNER JOIN AccountBalance ON Gl_Detail.AccountNo = AccountBalance.accountno"
gc_dbcon.Execute ls_sql
Call MsgBox("Successfull Updated", vbInformation)
End Sub

Private Sub Command3_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtFrom
    Set PO_DESC = txtFromDesc
    Gs_SQL = "Select Accountno 'Account No', Acct_Desc  'Description',BalAmount as Amount,DRCR from gl_Detail "
    Gs_FindFld = "Acct_Desc"
    Gs_Subon = True
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
    Gs_OrderBy = "Order by Acct_Desc,AccountNo"
    MyLookupOLDB.Caption = "Account Nos."
    MyLookupOLDB.Show 1
    
    If Len(txtFrom) > 0 Then txtFrom_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command4_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtTo
    Set PO_DESC = txttoDesc
    
    Gs_SQL = "Select Accountno 'Account No', Acct_Desc  'Description',BalAmount as Amount,DRCR from gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_Subon = True
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
    Gs_OrderBy = "Order by Acct_Desc,AccountNo"
    MyLookupOLDB.Caption = "Account Nos."
    MyLookupOLDB.Show 1
    
    If Len(txtTo) > 0 Then txtto_KeyDown vbKeyReturn, vbKeyShift
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

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If Lastkey(KeyCode) Then
       DTPTo.SetFocus
    End If
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Lastkey(KeyCode) And InStr(1, Me.Caption, "Ledger") > 0 Then
       If Frame7.Enabled Then txtFrom.SetFocus
    End If
End Sub

Private Sub Form_Load()
  
  Set PR_GlDetail = New Recordset
  PR_GlDetail.Open "Select * from Gl_Detail where CompCode ='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly
  PR_Branch.Open "Select * From SysBranch Where compcode = '" & Gs_compcode & "' Order By BranchCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
  Pb_BlnkVchr = IIf(PR_GlDetail.EOF, True, False)
  dtpfrom = DateValue(Gs_Fnperiod)
  If Year(dtpfrom) <> Year(Date) Then
     DTPTo = DateValue(Gs_FnEndPeriod)
  Else
     DTPTo = Date
  End If
  txtbranchcode = Gs_BranchCode
  If MySeek(txtbranchcode, "Branchcode", PR_Branch) Then
     txtbranchname = PR_Branch("BranchDesc")
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_GlDetail.Close
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
        If dtpfrom.Enabled Then
            dtpfrom.SetFocus
        Else
          DTPTo.SetFocus
        End If
     End If
  ElseIf KeyCode = vbKeyF12 Then
     Command5_Click
  ElseIf KeyCode = vbKeyReturn And txtbranchcode = "" Then
    txtbranchname = ""
  End If
End Sub

Private Sub txtbranchcode_LostFocus()
 If txtbranchcode = "" Then
    txtbranchname = ""
 End If
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtFrom.Text <> "" Then
    If txtFrom.Text <> "" Then
        lb_found = MySeek(txtFrom.Text, "AccountNo", PR_GlDetail)
        
        If lb_found Then
           txtFromDesc = PR_GlDetail("acct_desc")
           txttoDesc = PR_GlDetail("acct_desc")
            txtTo = Trim(txtFrom)
            txtTo.SetFocus
        Else
            Call SetErr("Record not found", vbCritical)
        End If
     ElseIf txtFrom.Text = "" Then
         txtTo.SetFocus
    End If
ElseIf KeyCode = vbKeyReturn And Trim(txtFrom) = "" Then
     Command3_Click
End If
End Sub
Private Sub txtto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtTo.Text <> "" Then
    If txtTo.Text <> "" Then
        lb_found = MySeek(txtTo.Text, "AccountNo", PR_GlDetail)
        If lb_found Then
            StatusBar1.Panels(2).Text = PR_GlDetail("acct_desc")
            cmdGenerate.SetFocus
        Else
            Call SetErr("Record not found", vbCritical)
        End If
    ElseIf txtFrom.Text = "" Then
        cmdGenerate.SetFocus
    End If
ElseIf KeyCode = vbKeyReturn And txtTo = "" Then
     Command4_Click
End If
End Sub
Private Sub RunBooks(Optional ls_Branch As String)
Dim ls_BkSql As String
Dim ld_stdt As Date
Dim ld_EndDt As Date
Dim ls_BookId As String
Dim ls_VchT As String
Dim ls_InData As String

ls_Branch = IIf(LTrim(RTrim(txtbranchcode)) = "", "", " And Gl_Trans.BranchCode = '" & txtbranchcode & "'")
ld_stdt = Format(DateValue(dtpfrom.Value), "YYYY/MM/DD")
ld_EndDt = Format(DateValue(DTPTo.Value), "YYYY/MM/DD")
'ls_BookId = IIf(Left(Me.Caption, 1) = "J", "G", IIf(Left(Me.Caption, 1) = "C", "S", "B"))
ls_VchT = Left(Me.Caption, 1)
ls_InData = IIf(ls_VchT = "J", "('D','C','G')", IIf(ls_VchT = "B", "('D','C','G','S')", "('D','C','G','B')"))

ls_BkSql = "SELECT Gl_Trans.*,Gl_Ref.Vchr_Remarks , gl_Detail.Acct_Type,gl_Detail.acct_Desc INTO Tmp_Books FROM Gl_Trans left outer JOIN Gl_Ref "
ls_BkSql = ls_BkSql + " ON Gl_Trans.compcode = Gl_Ref.CompCode AND Gl_Trans.Branchcode = Gl_Ref.BranchCode And gl_Trans.value_date = gl_Ref.value_date and Gl_Trans.Voucher_No = Gl_Ref.Voucher_no AND "
ls_BkSql = ls_BkSql + " Gl_Trans.VchrType = Gl_Ref.VchrType left outer JOIN gl_Detail ON Gl_Trans.CompCode = gl_Detail.CompCode AND "
ls_BkSql = ls_BkSql + " Gl_Trans.Accountno = gl_Detail.Accountno left outer JOIN GlVchrType on Gl_Trans.CompCode = glvchrtype.CompCode and Gl_Trans.BranchCode=glvchrtype.BranchCode And   gl_Trans.vchrtype = glVchrType.VchrType "
ls_BkSql = ls_BkSql + " Where gl_Trans.value_date between '" & Format(ld_stdt, "YYYY/MM/DD") & "' and '" & Format(ld_EndDt, "YYYY/MM/DD") & "' "
ls_BkSql = ls_BkSql + " and gl_Trans.Compcode = '" & Gs_compcode & "'" & ls_Branch & " and gl_Detail.acct_Type IN " & ls_InData & " and gl_Trans.VchrType <> '0OB' and glvchrtype.VchrBase = '" & ls_VchT & "'"
ls_BkSql = ls_BkSql + " order by 5,7,6,3"
gc_dbcon.BeginTrans
gc_dbcon.Execute (ls_BkSql)
gc_dbcon.CommitTrans
End Sub


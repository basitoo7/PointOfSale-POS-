VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmsaleregisterreport 
   Caption         =   "Print Invoice"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2580
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
   Icon            =   "frmsaleregisterreport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   2580
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2565
      Left            =   30
      TabIndex        =   1
      Top             =   -90
      Width           =   2520
      Begin VB.CheckBox chkdchallan 
         Caption         =   "&Delivery Challan"
         Height          =   240
         Left            =   960
         TabIndex        =   17
         Top             =   1815
         Width           =   1515
      End
      Begin VB.CheckBox chkinvoice 
         Caption         =   "&Invoice"
         Height          =   240
         Left            =   105
         TabIndex        =   16
         Top             =   1815
         Width           =   870
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
         Left            =   165
         MaskColor       =   &H00000000&
         TabIndex        =   11
         Top             =   2100
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   1290
         TabIndex        =   10
         Top             =   2100
         Width           =   1035
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
         Height          =   1065
         Left            =   90
         TabIndex        =   5
         Top             =   135
         Width           =   2370
         Begin VB.ListBox List1 
            Height          =   270
            Left            =   45
            TabIndex        =   15
            Top             =   135
            Visible         =   0   'False
            Width           =   225
         End
         Begin MSComCtl2.DTPicker dtpto 
            Height          =   315
            Left            =   825
            TabIndex        =   6
            Top             =   645
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            Format          =   65601537
            CurrentDate     =   37309
         End
         Begin MSComCtl2.DTPicker dtpfrom 
            Height          =   315
            Left            =   825
            TabIndex        =   7
            Top             =   210
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            Format          =   65601537
            CurrentDate     =   37309
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "From :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   330
            TabIndex        =   9
            Top             =   225
            Width           =   450
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "To :"
            Height          =   210
            Left            =   510
            TabIndex        =   8
            Top             =   690
            Width           =   270
         End
      End
      Begin VB.TextBox txtAcctNarration 
         Height          =   315
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   4
         Top             =   930
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Frame Frame7 
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
         Height          =   600
         Left            =   90
         TabIndex        =   2
         Top             =   1155
         Width           =   2370
         Begin VB.TextBox txtLocCode 
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "XXX"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   840
            MaxLength       =   10
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   180
            Width           =   1080
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   1950
            Picture         =   "frmsaleregisterreport.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   150
            Width           =   315
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   -240
            MaxLength       =   50
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   435
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Invoice :"
            Height          =   210
            Left            =   180
            TabIndex        =   14
            Top             =   210
            Width           =   600
         End
      End
      Begin Crystal.CrystalReport rptLedger 
         Left            =   2430
         Top             =   1830
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   2580
      _ExtentX        =   4551
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
Attribute VB_Name = "frmsaleregisterreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PR_Item As New Recordset
Public PO_DESC As Object
Public PO_CODE As Object
Dim PR_ICTran As New Recordset
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdGenerate_Click()
Dim ln_cnt As Integer
Dim ln_cnt1 As Integer
On Error GoTo LocalErr
Set PO_AnyForm = Nothing
Set PO_AnyForm = Me
Set PO_CODE = Text1
Set PO_DESC = List1
List1.Clear
frminvoiceinstr.Show 1
   With rptLedger
        .DiscardSavedData = True
        .WindowTitle = Me.Caption
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Sale Invoice'"
        .Formulas(3) = "Sign1 = '" & Gc_UserName & "'"
        .Formulas(4) = "PaymentType = '" & Text1 & "'"
        .SelectionFormula = "{IC_Trans.Value_Date} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {IC_Trans.Value_Date} <= Date(" & dtpto.Year & "," & dtpto.Month & "," & dtpto.Day & ") "
        .SelectionFormula = .SelectionFormula & "  and {Ic_Trans.TransType} = 'I'"
        If txtLocCode <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {Ic_Trans.Transc_No} = '" & txtLocCode & "'"
        End If
        ln_cnt1 = 5
        For ln_cnt = 0 To List1.ListCount - 1
        ' ls_instr = "Instr" + Str(ln_cnt) + "  = '" & List1.List(ln_cnt) & "'"
        .Formulas(ln_cnt1) = "Instr" + Trim(Str(ln_cnt + 1)) + "  = '" & ln_cnt + 1 & " - " & List1.List(ln_cnt) & "'"
        ln_cnt1 = ln_cnt1 + 1
        If ln_cnt1 > 10 Then Exit For
        Next
        If chkinvoice.Value = 1 And chkdchallan.Value = 1 Then
            .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "Invoice.rpt"
            .Action = 1
            .Formulas(4) = ""
            .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "delivery.rpt"
            .Action = 1
        ElseIf chkinvoice.Value = 1 Then
            .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "Invoice.rpt"
            .Action = 1
        Else
            .Formulas(4) = ""
            .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "delivery.rpt"
            .Action = 1
        End If
   End With
Exit Sub
LocalErr:
Call SetErr(Err.Description, vbCritical)
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtLocCode
    Set PO_DESC = Text1
    PR_ICTran.Filter = "Value_date >= '" & Format(dtpfrom, "YYYY/MM/DD") & "' and value_date <=  '" & Format(dtpto, "YYYY/MM/DD") & "' and Transtype = 'I'"
    GoTop PR_ICTran
    MyLookup.Caption = "Invoice"
    MyLookup.FillGrid PR_ICTran, "Transc_No", "Value_Date", 10
    MyLookup.Show 1
    PR_ICTran.Filter = adFilterNone
    
    If Len(txtLocCode) > 0 Then txtLocCode_KeyDown vbKeyReturn, vbKeyShift
    
End Sub

Private Sub dtpto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtLocCode.SetFocus
End Sub

Private Sub txtLocCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 If KeyCode = vbKeyReturn And Len(txtLocCode.Text) > 0 Then
         txtLocCode.Text = DoPad(txtLocCode.Text, txtLocCode.MaxLength)
         lb_found = MySeek(txtLocCode.Text, "Transc_No", PR_ICTran)
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtLocCode.SetFocus
             'txtLocDesc.Text = ""
         Else
            Text1.Text = PR_ICTran("Value_date")
            cmdGenerate.SetFocus
         End If
 ElseIf KeyCode = vbKeyF12 Then
        Command2_Click
 End If
End Sub
Private Sub dtpfrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then dtpto.SetFocus
End Sub
Private Sub Form_Load()
dtpfrom = Date
dtpto = Date
PR_ICTran.Open "Select Transc_No,Value_date,Transtype from Ic_Trans where compcode ='" & Gs_compcode & "' group by Transc_No,Value_date,Transtype ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_ICTran.Close
End Sub



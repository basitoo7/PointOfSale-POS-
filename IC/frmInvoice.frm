VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSOInvoice 
   Caption         =   "Print Note"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4080
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
   Icon            =   "frmInvoice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4080
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2865
      Width           =   4080
      _ExtentX        =   7197
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
   Begin VB.Frame Frame1 
      Height          =   2925
      Left            =   45
      TabIndex        =   1
      Top             =   -75
      Width           =   3990
      Begin VB.TextBox txtnoofcopy 
         Height          =   315
         Left            =   60
         TabIndex        =   23
         Text            =   "1"
         Top             =   2400
         Width           =   570
      End
      Begin VB.Frame Frame3 
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
         Height          =   1335
         Left            =   0
         TabIndex        =   4
         Top             =   30
         Width           =   3990
         Begin VB.ComboBox txtnotetype 
            Height          =   330
            ItemData        =   "frmInvoice.frx":030A
            Left            =   1155
            List            =   "frmInvoice.frx":0314
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   195
            Width           =   2760
         End
         Begin MSComCtl2.DTPicker dtpto 
            Height          =   315
            Left            =   1155
            TabIndex        =   5
            Top             =   945
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   101122049
            CurrentDate     =   37309
         End
         Begin MSComCtl2.DTPicker dtpfrom 
            Height          =   315
            Left            =   1155
            TabIndex        =   6
            Top             =   585
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   101122049
            CurrentDate     =   37309
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Note Type :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   285
            TabIndex        =   22
            Top             =   210
            Width           =   825
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "From Date :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   285
            TabIndex        =   8
            Top             =   600
            Width           =   825
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "To Date :"
            Height          =   210
            Left            =   450
            TabIndex        =   7
            Top             =   975
            Width           =   645
         End
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
         Height          =   390
         Left            =   1830
         MaskColor       =   &H00000000&
         TabIndex        =   10
         Top             =   2310
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   390
         Left            =   2895
         TabIndex        =   9
         Top             =   2310
         Width           =   1035
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
         Height          =   540
         Left            =   0
         TabIndex        =   2
         Top             =   1290
         Width           =   3990
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
            Left            =   1155
            MaxLength       =   10
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   150
            Width           =   1470
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   2625
            Picture         =   "frmInvoice.frx":033B
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   135
            Width           =   315
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   2955
            MaxLength       =   50
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   135
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Note # :"
            Height          =   210
            Left            =   525
            TabIndex        =   13
            Top             =   165
            Width           =   555
         End
      End
      Begin Crystal.CrystalReport rptLedger 
         Left            =   2280
         Top             =   2775
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
      Begin VB.Frame Frame2 
         Height          =   510
         Left            =   0
         TabIndex        =   14
         Top             =   1740
         Width           =   3990
         Begin VB.CheckBox chkprinter 
            Caption         =   "To Printer"
            Height          =   345
            Left            =   60
            TabIndex        =   25
            Top             =   120
            Width           =   1545
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Commercial Invoice"
            Height          =   240
            Left            =   2925
            TabIndex        =   16
            Top             =   135
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Sale Tax Invoice"
            Height          =   240
            Left            =   2025
            TabIndex        =   15
            Top             =   180
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1710
         End
      End
      Begin VB.Frame Frame4 
         Height          =   495
         Left            =   0
         TabIndex        =   17
         Top             =   1755
         Visible         =   0   'False
         Width           =   3975
         Begin VB.CheckBox Checksummary 
            Caption         =   "Summary"
            Height          =   210
            Left            =   60
            TabIndex        =   20
            Top             =   195
            Width           =   1125
         End
         Begin VB.CheckBox Checkdetail 
            Caption         =   "Detail"
            Height          =   210
            Left            =   1425
            TabIndex        =   19
            Top             =   195
            Width           =   1125
         End
         Begin VB.CheckBox Checkcost 
            Caption         =   "Cost"
            Height          =   210
            Left            =   2790
            TabIndex        =   18
            Top             =   195
            Width           =   1125
         End
      End
      Begin VB.Label Label3 
         Caption         =   "No of Copies"
         Height          =   285
         Left            =   675
         TabIndex        =   24
         Top             =   2415
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmSOInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PR_Item As New Recordset
Public PO_DESC As Object
Public PO_CODE As Object

Dim pr_dumy As New Recordset
Dim PR_Branch As New Recordset
Dim ls_sql As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdGenerate_Click()
On Error GoTo LocalErr
If txtnoteType.ListIndex = 0 Then
With rptLedger
        .DiscardSavedData = True
        .WindowTitle = Me.Caption
        
        If ln_changeprinter = 1 Then
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "AuoraSaleInvoice.rpt"
        Else
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "SaleInvoiceDC.rpt"
        End If
        
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .SQLQuery = "SELECT SO_TransMaster.TransCode, SO_TransMaster.TransDate, SO_TransMaster.DiscAmount, SO_TransMaster.RecAmount, SO_TransMaster.BalAmount, "
        .SQLQuery = .SQLQuery & " SO_TransMaster.CompName , SO_Trans.Quantity, SO_Trans.ItemRate, SO_Trans.Amount, SyUsers.UserName, IC_Item.Description,IC_Clients.Description"
        .SQLQuery = .SQLQuery & " FROM SO_TransMaster SO_TransMaster LEFT OUTER JOIN SyUsers SyUsers ON SO_TransMaster.Compcode = SyUsers.CompCode AND SO_TransMaster.UserCode = SyUsers.UserCode LEFT OUTER JOIN"
        .SQLQuery = .SQLQuery & " SO_Trans SO_Trans ON SO_TransMaster.Compcode = SO_Trans.Compcode AND SO_TransMaster.TransCode = SO_Trans.TransCode LEFT OUTER JOIN"
        .SQLQuery = .SQLQuery & " IC_Item IC_Item ON SO_Trans.Compcode = IC_Item.Compcode AND SO_Trans.ItemCode = IC_Item.ItemCode  "
        .SQLQuery = .SQLQuery & " LEFT OUTER JOIN IC_Clients IC_Clients ON SO_TransMaster.Compcode = IC_Clients.Compcode AND SO_TransMaster.AccountCode = IC_Clients.ClientCode"
        .SQLQuery = .SQLQuery & " where SO_TransMaster.compcode = '" & Gs_compcode & "'"
       
        
        .SelectionFormula = "{SO_TransMaster.compcode} = '" & Gs_compcode & "'"
        If txtLocCode <> "" Then
            .SQLQuery = .SQLQuery & " and  SO_TransMaster.transcode = '" & Trim(txtLocCode) & "'"
        End If
            .SQLQuery = .SQLQuery & " ORDER BY SO_TransMaster.TransCode "
            .Connect = "DNS=Censoft;UID=Sa"
            .Action = 1
End With
Else
With rptLedger
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "SaleInvoiceCredit.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        '.Formulas(2) = "Reportname = 'Good Receive Note'"
        .CopiesToPrinter = Val(txtnoofcopy)
        '.RetrieveSQLQuery
         If chkprinter.Value = 1 Then
        .Destination = crptToPrinter
        Else
        .Destination = crptToWindow
        End If
        .SQLQuery = "SELECT SO_TransMaster.TransCode, SO_TransMaster.TransDate, SO_TransMaster.DiscAmount, SO_TransMaster.CompName, SO_TransMaster.MiscAmount,"
        .SQLQuery = .SQLQuery & " SO_Trans.ItemCode, SO_Trans.SRNo, SO_Trans.Quantity, SO_Trans.ItemRate, SO_Trans.Amount, IC_Clients.Description, SyUsers.UserName,"
        .SQLQuery = .SQLQuery & " IC_Item.Description AS Expr1, IC_Item.PriceDescCStatus FROM SO_CreditSaleMaster SO_TransMaster LEFT OUTER JOIN"
        .SQLQuery = .SQLQuery & " IC_Clients IC_Clients ON SO_TransMaster.Compcode = IC_Clients.Compcode AND"
        .SQLQuery = .SQLQuery & " SO_TransMaster.AccountCode = IC_Clients.ClientCode LEFT OUTER JOIN  SyUsers SyUsers ON SO_TransMaster.Compcode = SyUsers.CompCode AND SO_TransMaster.UserCode = SyUsers.UserCode LEFT OUTER JOIN"
        .SQLQuery = .SQLQuery & " SO_CreditSaleTrans SO_Trans ON SO_TransMaster.Compcode = SO_Trans.Compcode AND"
        .SQLQuery = .SQLQuery & "  SO_TransMaster.TransCode = SO_Trans.TransCode LEFT OUTER JOIN IC_Item IC_Item ON SO_Trans.Compcode = IC_Item.Compcode AND SO_Trans.ItemCode = IC_Item.ItemCode"
        .SQLQuery = .SQLQuery & " where SO_TransMaster.compcode = '" & Gs_compcode & "' "
        .SQLQuery = .SQLQuery & " and convert(varchar,SO_TransMaster.Transdate,111) >= '" & Format(dtpfrom, "YYYY/MM/DD") & "' "
        .SQLQuery = .SQLQuery & " and convert(varchar,SO_TransMaster.Transdate,111) <= '" & Format(DTPTo, "YYYY/MM/DD") & "' "
        If txtLocCode <> "" Then
            .SQLQuery = .SQLQuery & " and  SO_TransMaster.transcode = '" & Trim(txtLocCode) & "'"
        End If
        
        .SQLQuery = .SQLQuery & "  ORDER BY SO_TransMaster.TransCode"
        
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
End With
End If

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
        If txtnoteType.ListIndex = 1 Then
            Gs_SQL = "SELECT Invoices.TransCode AS InvoiceNo, Invoices.TransDate AS Invoicedate, Customer.Description AS 'Customer.Description',"
            Gs_SQL = Gs_SQL & " Invoices.NetAmount AS 'Invoices.NetAmount', SyUsers.UserName fROM SO_CreditSaleMaster Invoices INNER JOIN"
            Gs_SQL = Gs_SQL & " IC_Clients Customer ON Invoices.Compcode = Customer.Compcode AND Invoices.AccountCode = Customer.ClientCode LEFT OUTER JOIN"
            Gs_SQL = Gs_SQL & " SyUsers ON Invoices.Compcode = SyUsers.CompCode AND Invoices.UserCode = SyUsers.UserCode"
           'Gs_OtherPara = " Where  Invoices.TransDate >= '" & Format(dtpfrom, "YYYY/MM/DD") & "' and Invoices.TransDate =< '" & Format(DTPTo, "YYYY/MM/DD") & "'"
            Gs_OtherPara = " Where convert(varchar,transdate,111) >= '" & Format(dtpfrom.Value, "YYYY/MM/DD") & "' and convert(varchar,transdate,111) <= '" & Format(DTPTo.Value, "YYYY/MM/DD") & "' "
            
            
            Gs_OrderBy = "ORDER BY Invoices.TransCode Desc"
        
        
        
        
        Else
            Gs_SQL = "SELECT Invoices.TransCode AS InvoiceNo, Invoices.TransDate AS Invoicedate, Customer.Description AS 'Customer.Description',"
            Gs_SQL = Gs_SQL & " Invoices.NetAmount AS 'Invoices.NetAmount', SyUsers.UserName fROM SO_TransMaster Invoices INNER JOIN"
            Gs_SQL = Gs_SQL & " IC_Clients Customer ON Invoices.Compcode = Customer.Compcode AND Invoices.AccountCode = Customer.ClientCode LEFT OUTER JOIN"
            Gs_SQL = Gs_SQL & " SyUsers ON Invoices.Compcode = SyUsers.CompCode AND Invoices.UserCode = SyUsers.UserCode"
           'Gs_OtherPara = " Where  Invoices.TransDate >= '" & Format(dtpfrom, "YYYY/MM/DD") & "' and Invoices.TransDate =< '" & Format(DTPTo, "YYYY/MM/DD") & "'"
            Gs_OtherPara = " Where convert(varchar,transdate,111) >= '" & Format(dtpfrom.Value, "YYYY/MM/DD") & "' and convert(varchar,transdate,111) <= '" & Format(DTPTo.Value, "YYYY/MM/DD") & "' "
            Gs_OrderBy = "ORDER BY Invoices.TransCode Desc"
        End If
        
        frmSosearchRecords.Caption = "Invoices"
        frmSosearchRecords.Show 1
        
        
        If txtLocCode <> "" Then Call txtLocCode_KeyDown(vbKeyReturn, vbKeyShift)
 
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtLocCode.SetFocus
End Sub

Private Sub txtLocCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
If txtLocCode <> "" And KeyCode = vbKeyReturn Then
txtLocCode = DoPad(txtLocCode, 10)
If txtnoteType.ListIndex = 0 Then
    ls_sql = "Select TransCode, TransDate from SO_TransMaster "
    ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "'  and Transcode = '" & txtLocCode & "'"
Else
    ls_sql = "Select TransCode, TransDate from SO_CreditSaleMaster "
    ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "'  and Transcode = '" & txtLocCode & "'"
End If
    
If pr_dumy.State = 1 Then pr_dumy.Close
 pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, 1
 If pr_dumy.EOF Then
  Call MsgBox(txtnoteType.Text & " not found !!!", vbCritical)
  txtLocCode.SetFocus
 Else
  Text1.Text = pr_dumy("TransDate")
 cmdGenerate.SetFocus
 End If
 pr_dumy.Close
ElseIf txtLocCode = "" And KeyCode = vbKeyReturn Then
        Command2_Click
 End If

End Sub
Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then DTPTo.SetFocus
End Sub
Private Sub Form_Load()
dtpfrom = Date
DTPTo = Date
txtnoteType.Text = "Sale Invoice"
Me.Caption = txtnoteType.Text
End Sub
Private Sub txtnotetype_Click()
Me.Caption = txtnoteType.Text
'If txtnotetype.Text = "Proposal Document" Then
'    Frame4.Visible = True
'Else
'    Frame4.Visible = False
'End If
'
'If txtnotetype.Text = "Sale Invoice" Then
'    Frame2.Visible = True
'Else
'    Frame2.Visible = False
'End If

End Sub

Private Sub txtnoteType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then dtpfrom.SetFocus
End Sub

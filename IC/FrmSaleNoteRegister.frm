VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSOSaleNoteRegisterRpt 
   Caption         =   "Print Note"
   ClientHeight    =   2625
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
   Icon            =   "FrmSaleNoteRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4080
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2250
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
      Height          =   2325
      Left            =   45
      TabIndex        =   1
      Top             =   -75
      Width           =   3990
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
            ItemData        =   "FrmSaleNoteRegister.frx":030A
            Left            =   1155
            List            =   "FrmSaleNoteRegister.frx":0320
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   180
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
            Format          =   54460417
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
            Format          =   54460417
            CurrentDate     =   37309
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Note Type :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   285
            TabIndex        =   14
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
         Left            =   1815
         MaskColor       =   &H00000000&
         TabIndex        =   10
         Top             =   1860
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   390
         Left            =   2895
         TabIndex        =   9
         Top             =   1860
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
            Picture         =   "FrmSaleNoteRegister.frx":0383
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
         Left            =   -30
         Top             =   1785
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
End
Attribute VB_Name = "frmSOSaleNoteRegisterRpt"
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
'On Error GoTo LocalErr
With rptLedger
        .DiscardSavedData = True
        .WindowTitle = Me.Caption
   If txtnoteType.Text = "Quotation" Then
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "ProposalDetail.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Quotation'"
        .SelectionFormula = "{PO_Transmaster.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {PO_Transmaster.TransDate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {PO_Transmaster.TransDate} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ") "
        If txtLocCode <> "" Then
            .SelectionFormula = .SelectionFormula & "  and {PO_Transmaster.transcode} = '" & Trim(txtLocCode) & "'"
        End If
       .Connect = "DNS=Censoft;UID=Sa"
       .Action = 1
    
    ElseIf txtnoteType.Text = "Job Order" Then

        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "JoborderRegister.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'JOB ORDER REGISTER'"
        .Formulas(3) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .SelectionFormula = "{PO_POordernote.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {PO_POOrdernote.TransDate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {PO_POOrdernote.TransDate} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ") "
        
        If txtLocCode <> "" Then
            .SelectionFormula = .SelectionFormula & "  and {PO_POOrdernote.transcode} = '" & Trim(txtLocCode) & "'"
        End If
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
     ElseIf txtnoteType.Text = "Job Order Completion" Then

        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "JoborderCompleteRegister.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'JOB ORDER COMPLETE REGISTER'"
        .Formulas(3) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .SelectionFormula = "{PO_POordernote.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {PO_POOrdernote.TransDate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {PO_POOrdernote.TransDate} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ") "
        
        If txtLocCode <> "" Then
            .SelectionFormula = .SelectionFormula & "  and {PO_POOrdernote.transcode} = '" & Trim(txtLocCode) & "'"
        End If
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
 
    ElseIf txtnoteType.Text = "Delivery Challan" Then
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "DeliveryChallanRegister.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Delivery Challan Register'"
        .Formulas(3) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .SelectionFormula = "{PO_POordernote.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {PO_POordernote.TransDate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {PO_POordernote.TransDate} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ") "
        
        If txtLocCode <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {PO_POordernote.transcode} = '" & Trim(txtLocCode) & "'"
        End If
        .Connect = "DNS=Censoft;UID=Sa"
       .Action = 1
    ElseIf txtnoteType.Text = "Gate Pass Outward" Then
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "GatePassOutwardregister.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Gate Pass Outward Register'"
        .Formulas(3) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .SelectionFormula = "{PO_POordernote.compcode} = '" & Gs_compcode & "'"
        
        .SelectionFormula = .SelectionFormula & " and {PO_POordernote.TransDate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {PO_POordernote.TransDate} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ") "
        
        If txtLocCode <> "" Then
            .SelectionFormula = .SelectionFormula & "  and {PO_POordernote.transcode} = '" & Trim(txtLocCode) & "'"
        End If
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
    ElseIf txtnoteType.Text = "Sale Invoice" Then
         .ReportFileName = App.Path & Gs_ICRepoPath & "\Saleregister.RPT"
         .WindowTitle = Me.Caption
         .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
         .Formulas(1) = "Reportname = 'Sale Register'"
         .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
'        .Formulas(3) = "Groupon = " & txtgroupon.ListIndex + 1 & ""
         .SelectionFormula = ""
         .SelectionFormula = "{So_SaleInvoice.CompCode} = '" & Gs_compcode & "'"
         .SelectionFormula = "{So_SaleInvoice.Transdate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {So_SaleInvoice.Transdate} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ")"
        
         If txtLocCode <> "" Then
            .SelectionFormula = .SelectionFormula & "  and {So_SaleInvoice.transcode} = '" & Trim(txtLocCode) & "'"
         End If
        .Connect = "DNS=Censoft;UID=Sa"
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
    
   If txtnoteType.Text = "Proposal Document" Then
        Gs_SQL = "Select TransCode, TransDate from SO_ProposalMaster "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Job Order" Then
        Gs_SQL = "Select TransCode, TransDate from SO_JobOrderMaster "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Delivery Challan" Then
        Gs_SQL = "Select TransCode, TransDate from SO_DeliveryMaster "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Gate Pass Outward" Then
        Gs_SQL = "Select TransCode, TransDate from SO_GatePass "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Sale Invoice" Then
        Gs_SQL = "Select TransCode, TransDate from So_SaleInvoice "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    
    End If
           MyLookupOLDB.Show 1
    If txtLocCode <> "" Then Call txtLocCode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtLocCode.SetFocus
End Sub

Private Sub txtLocCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 If KeyCode = vbKeyReturn And Len(txtLocCode.Text) > 0 Then
 txtLocCode = DoPad(txtLocCode, txtLocCode.MaxLength)
 If pr_dumy.State = 1 Then pr_dumy.Close
    If txtnoteType.Text = "Proposal Document" Then
        ls_sql = "Select TransCode, TransDate from SO_ProposalMaster "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Job Order" Then
        ls_sql = "Select TransCode, TransDate from SO_JobOrderMaster "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Delivery Challan" Then
        ls_sql = "Select TransCode, TransDate from SO_DeliveryMaster "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Gate Pass Outward" Then
        ls_sql = "Select TransCode, TransDate from SO_GatePass "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Sale Invoice" Then
        ls_sql = "Select TransCode, TransDate from So_SaleInvoice "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    End If
    
pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, 1
    If pr_dumy.EOF Then
        Call MsgBox(txtnoteType.Text & " not found !!!", vbCritical)
        txtLocCode.SetFocus
    Else
        Text1.Text = pr_dumy("TransDate")
    End If

ElseIf KeyCode = vbKeyF12 Then
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
End Sub

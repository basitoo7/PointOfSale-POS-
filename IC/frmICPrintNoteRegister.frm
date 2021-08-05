VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmICNoteRegisterReport 
   Caption         =   "Print Note"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5490
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
   Icon            =   "frmICPrintNoteRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5490
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1260
      Left            =   30
      TabIndex        =   8
      Top             =   -45
      Width           =   5445
      Begin VB.ComboBox txtnoteType 
         Height          =   330
         ItemData        =   "frmICPrintNoteRegister.frx":030A
         Left            =   1155
         List            =   "frmICPrintNoteRegister.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   165
         Width           =   2610
      End
      Begin MSComCtl2.DTPicker dtpto 
         Height          =   315
         Left            =   1155
         TabIndex        =   10
         Top             =   885
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   55902209
         CurrentDate     =   37309
      End
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   1155
         TabIndex        =   11
         Top             =   525
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   55902209
         CurrentDate     =   37309
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "To Date :"
         Height          =   210
         Left            =   450
         TabIndex        =   14
         Top             =   915
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "From Date :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   285
         TabIndex        =   13
         Top             =   570
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Note Type :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   285
         TabIndex        =   12
         Top             =   195
         Width           =   825
      End
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
      Height          =   1305
      Left            =   30
      TabIndex        =   3
      Top             =   1155
      Width           =   5445
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   1800
         Picture         =   "frmICPrintNoteRegister.frx":0350
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   870
         Width           =   315
      End
      Begin VB.TextBox txtVendorDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2130
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   870
         Width           =   3225
      End
      Begin VB.TextBox txtVendorCode 
         Height          =   315
         Left            =   1155
         MaxLength       =   6
         TabIndex        =   17
         Top             =   870
         Width           =   645
      End
      Begin VB.TextBox txtjobno 
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
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   510
         Width           =   1470
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2640
         Picture         =   "frmICPrintNoteRegister.frx":04C2
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   495
         Width           =   315
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   3570
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   150
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   2625
         Picture         =   "frmICPrintNoteRegister.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   135
         Width           =   315
      End
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
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   150
         Width           =   1470
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Vendor ID :"
         Height          =   255
         Left            =   165
         TabIndex        =   21
         Top             =   885
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "JOB # :"
         Height          =   210
         Left            =   525
         TabIndex        =   20
         Top             =   510
         Width           =   525
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Note # :"
         Height          =   210
         Left            =   525
         TabIndex        =   7
         Top             =   165
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   4425
      TabIndex        =   2
      Top             =   2490
      Width           =   1035
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
      Left            =   3345
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   2490
      Width           =   1035
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2880
      Width           =   5490
      _ExtentX        =   9684
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
   Begin Crystal.CrystalReport rptLedger 
      Left            =   1155
      Top             =   2610
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
Attribute VB_Name = "frmICNoteRegisterReport"
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
With rptLedger
        .DiscardSavedData = True
        .WindowTitle = Me.Caption
   If txtnoteType.Text = "Issue Note" Then

        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "IssueNoteRegister.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Issue Note Register'"
        .Formulas(3) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .SelectionFormula = "{PO_DemandNote.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {PO_DemandNote.TransDate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {PO_DemandNote.TransDate} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ") "
        
        If txtLocCode <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {PO_DemandNote.transcode} = '" & Trim(txtLocCode) & "'"
        End If
    
        If txtjobno <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {PO_DemandNote.JobNo} = '" & Trim(txtjobno) & "'"
        End If
    
    
    ElseIf txtnoteType.Text = "Issue Return Note" Then

        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "IssueNoteReturnRegister.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Issue Return Register'"
        .Formulas(3) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .SelectionFormula = "{PO_DemandNote.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {PO_DemandNote.TransDate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {PO_DemandNote.TransDate} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ") "
        
        If txtLocCode <> "" Then
            .SelectionFormula = .SelectionFormula & "  and {PO_DemandNote.transcode} = '" & Trim(txtLocCode) & "'"
        End If
       
        If txtjobno <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {PO_DemandNote.JobNo} = '" & Trim(txtjobno) & "'"
        End If
    
        If txtVendorCode <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {PO_DemandNote.AccountCode} = '" & Trim(txtVendorCode) & "'"
        End If
     ElseIf txtnoteType.Text = "Inventory Adjustment" Then

        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "AdjustmentRegister.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Adjustment Register'"
        .Formulas(3) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .SelectionFormula = "{PO_DemandNote.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {PO_DemandNote.TransDate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {PO_DemandNote.TransDate} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ") "
        
        If txtLocCode <> "" Then
            .SelectionFormula = .SelectionFormula & "  and {PO_DemandNote.transcode} = '" & Trim(txtLocCode) & "'"
        End If
       
        If txtjobno <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {PO_DemandNote.JobNo} = '" & Trim(txtjobno) & "'"
        End If
    
        If txtVendorCode <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {PO_DemandNote.AccountCode} = '" & Trim(txtVendorCode) & "'"
        End If
    End If
       .Connect = "DNS=Censoft;UID=Sa"
      .Action = 1
End With
Exit Sub
LocalErr:
Call SetErr(Err.Description, vbCritical)
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtjobno
    Set PO_DESC = Text1
    Gs_SQL = "SELECT  SO_JobOrderMaster.Transcode, IC_Clients.Description,IC_Clients.ClientCode"
    Gs_SQL = Gs_SQL & " FROM  SO_JobOrderMaster INNER JOIN IC_Clients ON SO_JobOrderMaster.Compcode = IC_Clients.Compcode AND SO_JobOrderMaster.ClientCode = IC_Clients.ClientCode"
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where SO_JobOrderMaster.Compcode  = '" & Gs_compcode & "' "
    MyLookupOLDB.Caption = "JobNo"
    MyLookupOLDB.Show 1
    
    If txtjobno <> "" Then Call txtJobNo_KeyDown(vbKeyReturn, vbKeyShift)
End Sub



Private Sub txtJobNo_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtjobno) <> "" And KeyCode = vbKeyReturn Then
        txtjobno = DoPad(txtjobno, 10)
        If pr_dumy.State = 1 Then pr_dumy.Close
        pr_dumy.Open "Select TransCode from SO_JobOrderMaster where Transcode  = '" & txtjobno & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Job Code not found !!!", vbCritical)
            txtjobno = ""
            txtjobno.SetFocus
        Else
            Text1 = Trim(pr_dumy("TransCode") & "")
            
        End If
        pr_dumy.Close
End If
End Sub
Private Sub Command5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtVendorCode
    Set PO_DESC = txtVendorDesc
    Gs_SQL = "Select SupplierCode, Description from IC_Supplier "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Supplier"
    MyLookupOLDB.Show 1
    
    If txtVendorCode <> "" Then Call txtVendorCode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub txtVendorCode_Change()
If txtVendorCode = "" Then txtVendorDesc = ""

End Sub

Private Sub txtVendorCode_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(txtVendorCode) <> "" And KeyCode = vbKeyReturn Then
        txtVendorCode = DoPad(txtVendorCode, 6)
        If pr_dumy.State = 1 Then pr_dumy.Close
        pr_dumy.Open "Select * from IC_Supplier where Compcode  = '" & Gs_compcode & "' and Suppliercode = '" & txtVendorCode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Vendor Code not found !!!", vbCritical)
            txtVendorCode = ""
            txtVendorDesc = ""
            txtVendorCode.SetFocus
        Else
            txtVendorDesc = pr_dumy("Description")
        End If
        pr_dumy.Close

ElseIf Trim(txtVendorCode) = "" And KeyCode = vbKeyReturn Then
        txtVendorCode = ""
        txtVendorDesc = ""
End If

End Sub

Private Sub Command2_Click()
   Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtLocCode
    Set PO_DESC = Text1
    
    
   If txtnoteType.Text = "Issue Note" Then
        Gs_SQL = "Select TransCode, TransDate from IC_IssueNoteMaster "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Issue Return Note" Then
        Gs_SQL = "Select TransCode, TransDate from IC_IssueReturnNoteMaster "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Inventory Adjustment" Then
        Gs_SQL = "Select TransCode, TransDate from IC_InventoryAdjMaster "
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
    If txtnoteType.Text = "Issue Note" Then
        ls_sql = "Select TransCode, TransDate from IC_IssueNoteMaster "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Issue Return Note" Then
        ls_sql = "Select TransCode, TransDate from IC_IssueReturnNoteMaster "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
         ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Inventory Adjustment" Then
        ls_sql = "Select TransCode, TransDate from IC_InventoryAdjMaster "
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
txtnoteType.Text = "Issue Note"

End Sub

Private Sub txtnotetype_Click()
Me.Caption = txtnoteType.Text
End Sub

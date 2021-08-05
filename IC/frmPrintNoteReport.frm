VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmICNoteReport 
   Caption         =   "Print Note"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4890
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
   Icon            =   "frmPrintNoteReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2205
      Width           =   4890
      _ExtentX        =   8625
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
      Height          =   2295
      Left            =   45
      TabIndex        =   1
      Top             =   -90
      Width           =   4785
      Begin VB.CommandButton Command1 
         Caption         =   "&Post"
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
         Left            =   60
         MaskColor       =   &H00000000&
         TabIndex        =   16
         Top             =   1875
         Width           =   1035
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
         Height          =   1260
         Left            =   0
         TabIndex        =   4
         Top             =   105
         Width           =   4755
         Begin VB.ComboBox txtnoteType 
            Height          =   330
            ItemData        =   "frmPrintNoteReport.frx":030A
            Left            =   1155
            List            =   "frmPrintNoteReport.frx":0317
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   165
            Width           =   2610
         End
         Begin MSComCtl2.DTPicker dtpto 
            Height          =   315
            Left            =   1155
            TabIndex        =   5
            Top             =   885
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   132382721
            CurrentDate     =   37309
         End
         Begin MSComCtl2.DTPicker dtpfrom 
            Height          =   315
            Left            =   1155
            TabIndex        =   6
            Top             =   525
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   132382721
            CurrentDate     =   37309
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Note Type :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   285
            TabIndex        =   15
            Top             =   195
            Width           =   825
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "From Date :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   285
            TabIndex        =   8
            Top             =   570
            Width           =   825
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "To Date :"
            Height          =   210
            Left            =   450
            TabIndex        =   7
            Top             =   915
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
         Left            =   2640
         MaskColor       =   &H00000000&
         TabIndex        =   10
         Top             =   1845
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   390
         Left            =   3720
         TabIndex        =   9
         Top             =   1845
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
         Width           =   4755
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
            Picture         =   "frmPrintNoteReport.frx":034B
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   135
            Width           =   315
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   3570
            MaxLength       =   50
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   150
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
         Left            =   -15
         Top             =   2085
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
Attribute VB_Name = "frmICNoteReport"
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

        '.ReportFileName = App.Path & Gs_ICRepoPath & "\" & "IssueNote.rpt"
        
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "ISSUENOTE_New.rpt"
        
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = '" & txtnoteType.Text & "'"
        .SelectionFormula = "{PO_DemandNote.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {PO_DemandNote.TransDate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {PO_DemandNote.TransDate} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ") "
        
        If txtLocCode <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {PO_DemandNote.transcode} = '" & Trim(txtLocCode) & "'"
        End If
   ElseIf txtnoteType.Text = "Adjustment Note" Then

        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "AdjustmentNote.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = '" & txtnoteType.Text & "'"
        .SelectionFormula = "{PO_DemandNote.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {PO_DemandNote.TransDate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {PO_DemandNote.TransDate} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ") "
        
        If txtLocCode <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {PO_DemandNote.transcode} = '" & Trim(txtLocCode) & "'"
        End If
   ElseIf txtnoteType.Text = "Issue Return Note" Then

        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "IssueNoteReturn.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = '" & txtnoteType.Text & "'"
        .SelectionFormula = "{PO_DemandNote.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {PO_DemandNote.TransDate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {PO_DemandNote.TransDate} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ") "
        
        If txtLocCode <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {PO_DemandNote.transcode} = '" & Trim(txtLocCode) & "'"
        End If

    ElseIf txtnoteType.Text = "Purchase Order" Then

        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "Purchaseorder.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'PURCHASE ORDER'"
        .SelectionFormula = "{PO_POordernote.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {PO_POOrdernote.TransDate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {PO_POOrdernote.TransDate} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ") "
        
        If txtLocCode <> "" Then
            .SelectionFormula = .SelectionFormula & "  and {PO_POOrdernote.transcode} = '" & Trim(txtLocCode) & "'"
        End If

    ElseIf txtnoteType.Text = "Gate Pass Inward" Then
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "GatePassInward.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Gate Pass Inward'"
        .SelectionFormula = "{PO_GatePass.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {PO_GatePass.TransDate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {PO_GatePass.TransDate} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ") "
        
        If txtLocCode <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {PO_GatePass.transcode} = '" & Trim(txtLocCode) & "'"
        End If
    ElseIf txtnoteType.Text = "Inspection Note" Then
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "POInspection.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Inspection Note'"
        .SelectionFormula = "{PO_Inspection.compcode} = '" & Gs_compcode & "'"
        
        .SelectionFormula = .SelectionFormula & " and {PO_Inspection.TransDate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {PO_Inspection.TransDate} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ") "
        
        If txtLocCode <> "" Then
            .SelectionFormula = .SelectionFormula & "  and {PO_Inspection.transcode} = '" & Trim(txtLocCode) & "'"
        End If
      
    ElseIf txtnoteType.Text = "Good Receive Note" Then
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "POGRN.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Good Receive Note'"
        .SelectionFormula = "{PO_POOrderNote.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {PO_POOrdernote.TransDate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {PO_POOrdernote.TransDate} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ") "
        If txtLocCode <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {PO_POOrderNote.transcode} = '" & Trim(txtLocCode) & "'"
        End If
    ElseIf txtnoteType.Text = "Good Receive Return Note" Then
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "POGRNreturn.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Good Receive Return Note'"
        .SelectionFormula = "{PO_POOrderNote.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {PO_POOrdernote.TransDate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {PO_POOrdernote.TransDate} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ") "
        If txtLocCode <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {PO_POOrderNote.transcode} = '" & Trim(txtLocCode) & "'"
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
If txtnoteType.Text = "Issue Note" Then
ls_sql = "update IC_IssueNoteMaster set glstatus = 1 "
ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' "
If txtLocCode <> "" Then
    ls_sql = ls_sql & " and Transcode = '" & txtLocCode & "'"
End If
 
ElseIf txtnoteType.Text = "Adjustment Note" Then
ls_sql = "update IC_InventoryAdjMaster set glstatus = 1 "
ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' "
If txtLocCode <> "" Then
    ls_sql = ls_sql & " and Transcode = '" & txtLocCode & "'"
End If
End If
gc_dbcon.Execute ls_sql
Call MsgBox("Successfully Posted", vbInformation)
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
   ElseIf txtnoteType.Text = "Adjustment Note" Then
        Gs_SQL = "Select TransCode, TransDate from IC_InventoryAdjMaster "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' "
        MyLookupOLDB.Caption = txtnoteType.Text

    ElseIf txtnoteType.Text = "Purchase Order" Then
        Gs_SQL = "Select TransCode, TransDate from PO_POOrderNote "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Gate Pass Inward" Then
        Gs_SQL = "Select TransCode, TransDate from PO_GatePass "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Inspection Note" Then
        Gs_SQL = "Select TransCode, TransDate from PO_Inspection "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Good Receive Note" Then
        Gs_SQL = "Select TransCode, TransDate from PO_POGRN "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Good Receive Return Note" Then
        Gs_SQL = "Select TransCode, TransDate from PO_POGRNReturn "
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
    ElseIf txtnoteType.Text = "Adjustment Note" Then
        ls_sql = "Select TransCode, TransDate from IC_InventoryAdjMaster "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    
    ElseIf txtnoteType.Text = "Purchase Order" Then
        ls_sql = "Select TransCode, TransDate from PO_POOrderNote "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Gate Pass Inward" Then
        ls_sql = "Select TransCode, TransDate from PO_GatePass "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Inspection Note" Then
        ls_sql = "Select TransCode, TransDate from PO_Inspection "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Good Receive Note" Then
        ls_sql = "Select TransCode, TransDate from PO_POGRN "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Good Receive Return Note" Then
        ls_sql = "Select TransCode, TransDate from PO_POGRNReturn "
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

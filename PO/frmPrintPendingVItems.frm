VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPOPendingVitems 
   Caption         =   "Vendor Pending Items"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
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
   Icon            =   "frmPrintPendingVItems.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   5670
      TabIndex        =   7
      Top             =   1260
      Width           =   1005
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
      Height          =   360
      Left            =   4650
      TabIndex        =   6
      Top             =   1260
      Width           =   1005
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
      Left            =   30
      TabIndex        =   1
      Top             =   -45
      Width           =   6645
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   0
         MaxLength       =   50
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   0
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   1815
         Picture         =   "frmPrintPendingVItems.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   885
         Width           =   315
      End
      Begin VB.TextBox txtLocDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   885
         Width           =   4410
      End
      Begin VB.TextBox txtLocCode 
         Height          =   315
         Left            =   1155
         MaxLength       =   6
         TabIndex        =   8
         Top             =   885
         Width           =   645
      End
      Begin MSComCtl2.DTPicker dtpto 
         Height          =   315
         Left            =   1155
         TabIndex        =   2
         Top             =   510
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61472769
         CurrentDate     =   37309
      End
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   1155
         TabIndex        =   3
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61472769
         CurrentDate     =   37309
      End
      Begin Crystal.CrystalReport rptLedger 
         Left            =   0
         Top             =   0
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
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Vendor ID :"
         Height          =   255
         Left            =   195
         TabIndex        =   11
         Top             =   915
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "From Date :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   285
         TabIndex        =   5
         Top             =   195
         Width           =   825
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "To Date :"
         Height          =   210
         Left            =   450
         TabIndex        =   4
         Top             =   540
         Width           =   645
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1695
      Width           =   6720
      _ExtentX        =   11853
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
Attribute VB_Name = "frmPOPendingVitems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PR_Item As New Recordset
Public PO_DESC As Object
Public PO_CODE As Object

Dim PR_Dumy As New Recordset
Dim PR_Branch As New Recordset
Dim ls_sql As String

Private Sub cmdGenerate_Click()
On Error GoTo LocalErr
With rptLedger
        .DiscardSavedData = True
        .WindowTitle = Me.Caption
 If txtnoteType.Text = "Good Receive Note" Then
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "POGRN.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Good Receive Note'"
        .SelectionFormula = "{PO_POOrderNote.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {PO_POOrdernote.TransDate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {PO_POOrdernote.TransDate} <= Date(" & dtpto.Year & "," & dtpto.Month & "," & dtpto.Day & ") "
        If txtLocCode <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {PO_POOrderNote.transcode} = '" & Trim(txtLocCode) & "'"
        End If
        If txtjobNo <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {PO_POOrdernote.JobNo} = '" & Trim(txtjobNo) & "'"
        End If
        If txtVendorCode <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {PO_POOrdernote.AccountCode} = '" & Trim(txtVendorCode) & "'"
        End If
    ElseIf txtnoteType.Text = "Good Receive Return Note" Then
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "POGRNreturn.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Good Receive Return Note'"
        .SelectionFormula = "{PO_POOrderNote.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {PO_POOrdernote.TransDate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {PO_POOrdernote.TransDate} <= Date(" & dtpto.Year & "," & dtpto.Month & "," & dtpto.Day & ") "
        If txtLocCode <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {PO_POOrderNote.transcode} = '" & Trim(txtLocCode) & "'"
        End If

        If txtjobNo <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {PO_POOrdernote.JobNo} = '" & Trim(txtjobNo) & "'"
        End If
        If txtVendorCode <> "" Then
        .SelectionFormula = .SelectionFormula & "  and {PO_POOrdernote.AccountCode} = '" & Trim(txtVendorCode) & "'"
        End If
End If
      .Action = 1
End With
Exit Sub
LocalErr:
Call SetErr(Err.Description, vbCritical)
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub Command5_Click()
 Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtLocCode
    Set PO_DESC = Text1
    
    
   If txtnoteType.Text = "Demand Note" Then
        Gs_SQL = "Select TransCode, TransDate from PO_DemandNote "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Purchase Order" Then
        Gs_SQL = "Select TransCode, TransDate from PO_POOrderNote "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Gate Pass Inward" Then
        Gs_SQL = "Select TransCode, TransDate from PO_GatePass "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Inspection Note" Then
        Gs_SQL = "Select TransCode, TransDate from PO_Inspection "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Inspection Return Note" Then
        Gs_SQL = "Select TransCode, TransDate from PO_InspectionReturn "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Good Receive Note" Then
        Gs_SQL = "Select TransCode, TransDate from PO_POGRN "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Good Receive Return Note" Then
        Gs_SQL = "Select TransCode, TransDate from PO_POGRNReturn "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    
    End If
        MyLookupOLDB.Show 1
    
   If txtLocCode <> "" Then Call txtLocCode_KeyDown(vbKeyReturn, vbKeyShift)
   End Sub

Private Sub txtVendorCode_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(txtVendorCode) <> "" And KeyCode = vbKeyReturn Then
        txtVendorCode = DoPad(txtVendorCode, 6)
        If PR_Dumy.State = 1 Then PR_Dumy.Close
        PR_Dumy.Open "Select * from IC_Supplier where Compcode  = '" & Gs_compcode & "' and Suppliercode = '" & txtVendorCode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If PR_Dumy.EOF Then
            Call MsgBox("Vendor Code not found !!!", vbCritical)
            txtVendorCode = ""
            txtVendorDesc = ""
            txtVendorCode.SetFocus
        Else
            txtVendorDesc = PR_Dumy("Description")
            cmdGenerate.SetFocus
        End If
        PR_Dumy.Close

ElseIf Trim(txtVendorCode) = "" And KeyCode = vbKeyReturn Then
        txtVendorCode = ""
        txtVendorDesc = ""
        Command5_Click
End If

End Sub


Private Sub Command3_Click()
Unload Me
End Sub
Private Sub Command2_Click()
 Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtLocCode
    Set PO_DESC = txtLocDesc
    
    
   If txtnoteType.Text = "Demand Note" Then
        Gs_SQL = "Select TransCode, TransDate from PO_DemandNote "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Purchase Order" Then
        Gs_SQL = "Select TransCode, TransDate from PO_POOrderNote "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Gate Pass Inward" Then
        Gs_SQL = "Select TransCode, TransDate from PO_GatePass "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Inspection Note" Then
        Gs_SQL = "Select TransCode, TransDate from PO_Inspection "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Inspection Return Note" Then
        Gs_SQL = "Select TransCode, TransDate from PO_InspectionReturn "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Good Receive Note" Then
        Gs_SQL = "Select TransCode, TransDate from PO_POGRN "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = txtnoteType.Text
    ElseIf txtnoteType.Text = "Good Receive Return Note" Then
        Gs_SQL = "Select TransCode, TransDate from PO_POGRNReturn "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
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

 If PR_Dumy.State = 1 Then PR_Dumy.Close
    If txtnoteType.Text = "Demand Note" Then
        ls_sql = "Select TransCode, TransDate from PO_DemandNote "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Purchase Order" Then
        ls_sql = "Select TransCode, TransDate from PO_POOrderNote "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Gate Pass Inward" Then
        ls_sql = "Select TransCode, TransDate from PO_GatePass "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Inspection Note" Then
        ls_sql = "Select TransCode, TransDate from PO_Inspection "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Inspection Return Note" Then
        ls_sql = "Select TransCode, TransDate from PO_Inspectionreturn "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Good Receive Note" Then
        ls_sql = "Select TransCode, TransDate from PO_POGRN "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    ElseIf txtnoteType.Text = "Good Receive Return Note" Then
        ls_sql = "Select TransCode, TransDate from PO_POGRNReturn "
        ls_sql = ls_sql & "  where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "' and Transcode = '" & txtLocCode & "'"
    End If
    PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, 1
    If PR_Dumy.EOF Then
        Call MsgBox(txtnoteType.Text & " not found !!!", vbCritical)
        txtLocCode.SetFocus
    Else
        Text1.Text = PR_Dumy("TransDate")
        txtVendorCode.SetFocus
    End If

 ElseIf KeyCode = vbKeyReturn And Len(txtLocCode.Text) = 0 Then
        Command2_Click
  End If

End Sub
Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then dtpto.SetFocus
End Sub
Private Sub Form_Load()
dtpfrom = Date
dtpto = Date

End Sub


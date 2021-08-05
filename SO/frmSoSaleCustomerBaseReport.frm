VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSoSaleCustomerBaseReport 
   Caption         =   "Customer Base Item Wise Report"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
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
   Icon            =   "frmSoSaleCustomerBaseReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLocCode 
      Height          =   315
      Left            =   1375
      MaxLength       =   10
      TabIndex        =   17
      Top             =   600
      Width           =   1035
   End
   Begin VB.CheckBox chkinvdatewisesum 
      Caption         =   "Invoice && Date Wise Summary"
      Height          =   315
      Left            =   60
      TabIndex        =   14
      Top             =   1950
      Width           =   2700
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   4695
      TabIndex        =   7
      Top             =   1920
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
      Left            =   3675
      TabIndex        =   6
      Top             =   1920
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
      Height          =   1965
      Left            =   30
      TabIndex        =   1
      Top             =   -45
      Width           =   5730
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2430
         Picture         =   "frmSoSaleCustomerBaseReport.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox txtcustomerdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   270
         Width           =   2880
      End
      Begin VB.TextBox txtcustomercode 
         Height          =   315
         Left            =   1350
         TabIndex        =   10
         Top             =   270
         Width           =   1050
      End
      Begin VB.CheckBox chkcash 
         Caption         =   "Cash"
         Height          =   270
         Left            =   4680
         TabIndex        =   9
         Top             =   1560
         Width           =   1020
      End
      Begin VB.CheckBox chkcredit 
         Caption         =   "Credit"
         Height          =   270
         Left            =   3870
         TabIndex        =   8
         Top             =   1560
         Width           =   1020
      End
      Begin MSComCtl2.DTPicker dtpto 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   1485
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   102563841
         CurrentDate     =   37309
      End
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   1350
         TabIndex        =   3
         Top             =   1125
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   102563841
         CurrentDate     =   37309
      End
      Begin Crystal.CrystalReport rptLedger 
         Left            =   120
         Top             =   1440
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
      Begin VB.Label txtselectiveaccount 
         Height          =   300
         Left            =   2760
         TabIndex        =   16
         Top             =   240
         Width           =   2520
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Invoice No :"
         Height          =   210
         Left            =   480
         TabIndex        =   15
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Customer Code :"
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   285
         Width           =   1200
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "From Date :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   480
         TabIndex        =   5
         Top             =   1155
         Width           =   825
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "To Date :"
         Height          =   210
         Left            =   645
         TabIndex        =   4
         Top             =   1515
         Width           =   645
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   2415
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   635
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Customer Code :"
      Height          =   210
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   1200
   End
End
Attribute VB_Name = "frmSoSaleCustomerBaseReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PR_Item As New Recordset
Public PO_DESC As Object
Public PO_CODE As Object
Dim ls_sql As String
Dim pr_dumy As New Recordset

Private Sub chkcredit_Click()
If chkcredit.Value = 1 Then
chkcash.Value = 0
End If
End Sub
Private Sub chkcash_Click()
If chkcash.Value = 1 Then
chkcredit.Value = 0
End If
End Sub

Private Sub cmdGenerate_Click()
On Error GoTo LocalErr

If chkinvdatewisesum.Value = 1 Then

MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
 
Call ChkTempTables("Tmp_SaleReport", True)


ls_sql = "SELECT SO_TransMaster.TransCode,  SO_TransMaster.TransDate , SO_TransMaster.AccountCode,"
ls_sql = ls_sql & " Sum(SO_Trans.Amount) as Amount, sum(SO_Trans.Discamount) as Discamount into Tmp_SaleReport"
ls_sql = ls_sql & " FROM  SO_TransMastermain SO_TransMaster INNER JOIN  SO_Transmain SO_Trans ON SO_TransMaster.TransCode = SO_Trans.TransCode"
ls_sql = ls_sql & " WHERE CONVERT(varchar, SO_TransMaster.TransDate, 111) >=  '" & Format(dtpfrom, "YYYY/MM/DD") & "' and   CONVERT(varchar, SO_TransMaster.TransDate, 111) <=  '" & Format(dtpto, "YYYY/MM/DD") & "' "

If txtcustomercode <> "" Then
    If txtcustomercode = "Selective" Then
    ls_sql = ls_sql & "  and SO_TransMaster.AccountCode in (" & txtselectiveaccount & ")"
    Else
    ls_sql = ls_sql & "  and SO_TransMaster.AccountCode = '" & txtcustomercode & "'"
    End If
End If



If chkcash.Value = 1 Then
    ls_sql = ls_sql & "  and SO_TransMaster.salestatus  = 0"

ElseIf chkcredit.Value = 1 Then
ls_sql = ls_sql & "  and SO_TransMaster.salestatus  = 1"
End If

ls_sql = ls_sql & " Group by SO_TransMaster.TransCode,  SO_TransMaster.TransDate , SO_TransMaster.AccountCode"

ls_sql = ls_sql & " Union All"

ls_sql = ls_sql & " SELECT SO_TransMaster.TransCode, SO_TransMaster.TransDate, SO_TransMaster.AccountCode,"
ls_sql = ls_sql & " Sum(SO_Trans.Amount) as Amount, sum(SO_Trans.Discamount) as DiscAmount"
ls_sql = ls_sql & " FROM SO_TransMaster INNER JOIN SO_Trans ON SO_TransMaster.TransCode = SO_Trans.TransCode"
ls_sql = ls_sql & " WHERE CONVERT(varchar, SO_TransMaster.TransDate, 111) >=  '" & Format(dtpfrom, "YYYY/MM/DD") & "' and   CONVERT(varchar, SO_TransMaster.TransDate, 111) <=  '" & Format(dtpto, "YYYY/MM/DD") & "' "



If txtcustomercode <> "" Then
    If txtcustomercode = "Selective" Then
    ls_sql = ls_sql & "  and SO_TransMaster.AccountCode in (" & txtselectiveaccount & ")"
    Else
    ls_sql = ls_sql & "  and SO_TransMaster.AccountCode = '" & txtcustomercode & "'"
    End If
End If

If chkcash.Value = 1 Then
    ls_sql = ls_sql & "  and SO_TransMaster.salestatus  = 0"

ElseIf chkcredit.Value = 1 Then
ls_sql = ls_sql & "  and SO_TransMaster.salestatus  = 1"
End If

ls_sql = ls_sql & " Group by SO_TransMaster.TransCode,  SO_TransMaster.TransDate , SO_TransMaster.AccountCode"


ls_sql = ls_sql & " Union All"

ls_sql = ls_sql & " SELECT SO_CreditSaleMaster.TransCode, SO_CreditSaleMaster.TransDate, SO_CreditSaleMaster.AccountCode,"
ls_sql = ls_sql & "  Sum(SO_CreditSaleTrans.Amount) as Amount, sum(SO_CreditSaleTrans.Discamount) as Discamount"
ls_sql = ls_sql & " FROM SO_CreditSaleMaster INNER JOIN SO_CreditSaleTrans ON SO_CreditSaleMaster.TransCode = SO_CreditSaleTrans.TransCode"
ls_sql = ls_sql & " WHERE CONVERT(varchar, SO_CreditSaleMaster.TransDate, 111) >=  '" & Format(dtpfrom, "YYYY/MM/DD") & "' and   CONVERT(varchar, SO_CreditSaleMaster.TransDate, 111) <=  '" & Format(dtpto, "YYYY/MM/DD") & "' "



If txtcustomercode <> "" Then
    If txtcustomercode = "Selective" Then
    ls_sql = ls_sql & "  and SO_CreditSaleMaster.AccountCode in (" & txtselectiveaccount & ")"
    Else
    ls_sql = ls_sql & "  and SO_CreditSaleMaster.AccountCode = '" & txtcustomercode & "'"
    End If
End If

If chkcash.Value = 1 Then
    ls_sql = ls_sql & "  and SO_CreditSaleMaster.salestatus  = 0"

ElseIf chkcredit.Value = 1 Then
ls_sql = ls_sql & "  and SO_CreditSaleMaster.salestatus  = 1"
End If

ls_sql = ls_sql & " Group by SO_CreditSaleMaster.TransCode,  SO_CreditSaleMaster.TransDate , SO_CreditSaleMaster.AccountCode"

gc_dbcon.Execute ls_sql

With rptLedger
        .DiscardSavedData = True
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "SaleReportCustomerbasesum.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = 'Customer Statement'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & dtpto & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
End With
MDIForm1.StatusBar1.Panels(7).Text = ""



Exit Sub
End If


MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
 
Call ChkTempTables("Tmp_SaleReport", True)


ls_sql = "SELECT SO_TransMaster.TransCode,  SO_TransMaster.TransDate , SO_TransMaster.AccountCode,"
ls_sql = ls_sql & " SO_Trans.CustomCode, SO_Trans.ItemCode, SO_Trans.Quantity, SO_Trans.ItemRate, SO_Trans.Amount, SO_Trans.Discamount into Tmp_SaleReport"
ls_sql = ls_sql & " FROM  SO_TransMastermain SO_TransMaster INNER JOIN  SO_Transmain SO_Trans ON SO_TransMaster.TransCode = SO_Trans.TransCode"
ls_sql = ls_sql & " WHERE CONVERT(varchar, SO_TransMaster.TransDate, 111) >=  '" & Format(dtpfrom, "YYYY/MM/DD") & "' and   CONVERT(varchar, SO_TransMaster.TransDate, 111) <=  '" & Format(dtpto, "YYYY/MM/DD") & "' "

If txtcustomercode <> "" Then
    If txtcustomercode = "Selective" Then
    ls_sql = ls_sql & "  and SO_TransMaster.AccountCode in (" & txtselectiveaccount & ")"
    Else
    ls_sql = ls_sql & "  and SO_TransMaster.AccountCode = '" & txtcustomercode & "'"
    End If
End If


If txtLocCode <> "" Then
    'If txtcustomercode = "Selective" Then
    'ls_sql = ls_sql & "  and SO_TransMaster.AccountCode in (" & txtselectiveaccount & ")"
    'Else
    ls_sql = ls_sql & "  and SO_TransMaster.Transcode = '" & txtLocCode & "'"
   ' End If
End If



If chkcash.Value = 1 Then
    ls_sql = ls_sql & "  and SO_TransMaster.salestatus  = 0"

ElseIf chkcredit.Value = 1 Then
ls_sql = ls_sql & "  and SO_TransMaster.salestatus  = 1"
End If



ls_sql = ls_sql & " Union All"

ls_sql = ls_sql & " SELECT SO_TransMaster.TransCode, SO_TransMaster.TransDate, SO_TransMaster.AccountCode,"
ls_sql = ls_sql & " SO_Trans.CustomCode , SO_Trans.ItemCode, SO_Trans.Quantity, SO_Trans.ItemRate, SO_Trans.Amount, SO_Trans.Discamount"
ls_sql = ls_sql & " FROM SO_TransMaster INNER JOIN SO_Trans ON SO_TransMaster.TransCode = SO_Trans.TransCode"
ls_sql = ls_sql & " WHERE CONVERT(varchar, SO_TransMaster.TransDate, 111) >=  '" & Format(dtpfrom, "YYYY/MM/DD") & "' and   CONVERT(varchar, SO_TransMaster.TransDate, 111) <=  '" & Format(dtpto, "YYYY/MM/DD") & "' "



If txtcustomercode <> "" Then
    If txtcustomercode = "Selective" Then
    ls_sql = ls_sql & "  and SO_TransMaster.AccountCode in (" & txtselectiveaccount & ")"
    Else
    ls_sql = ls_sql & "  and SO_TransMaster.AccountCode = '" & txtcustomercode & "'"
    End If
End If

If txtLocCode <> "" Then
    'If txtcustomercode = "Selective" Then
    'ls_sql = ls_sql & "  and SO_TransMaster.AccountCode in (" & txtselectiveaccount & ")"
    'Else
    ls_sql = ls_sql & "  and SO_TransMaster.Transcode = '" & txtLocCode & "'"
   ' End If
End If



If chkcash.Value = 1 Then
    ls_sql = ls_sql & "  and SO_TransMaster.salestatus  = 0"

ElseIf chkcredit.Value = 1 Then
ls_sql = ls_sql & "  and SO_TransMaster.salestatus  = 1"
End If


ls_sql = ls_sql & " Union All"

ls_sql = ls_sql & " SELECT SO_CreditSaleMaster.TransCode, SO_CreditSaleMaster.TransDate, SO_CreditSaleMaster.AccountCode,"
ls_sql = ls_sql & " SO_CreditSaleTrans.CustomCode , SO_CreditSaleTrans.ItemCode, SO_CreditSaleTrans.Quantity, SO_CreditSaleTrans.ItemRate, SO_CreditSaleTrans.Amount, SO_CreditSaleTrans.Discamount"
ls_sql = ls_sql & " FROM SO_CreditSaleMaster INNER JOIN SO_CreditSaleTrans ON SO_CreditSaleMaster.TransCode = SO_CreditSaleTrans.TransCode"
ls_sql = ls_sql & " WHERE CONVERT(varchar, SO_CreditSaleMaster.TransDate, 111) >=  '" & Format(dtpfrom, "YYYY/MM/DD") & "' and   CONVERT(varchar, SO_CreditSaleMaster.TransDate, 111) <=  '" & Format(dtpto, "YYYY/MM/DD") & "' "



If txtcustomercode <> "" Then
    If txtcustomercode = "Selective" Then
    ls_sql = ls_sql & "  and SO_CreditSaleMaster.AccountCode in (" & txtselectiveaccount & ")"
    Else
    ls_sql = ls_sql & "  and SO_CreditSaleMaster.AccountCode = '" & txtcustomercode & "'"
    End If
End If



If chkcash.Value = 1 Then
    ls_sql = ls_sql & "  and SO_CreditSaleMaster.salestatus  = 0"

ElseIf chkcredit.Value = 1 Then
ls_sql = ls_sql & "  and SO_CreditSaleMaster.salestatus  = 1"
End If


gc_dbcon.Execute ls_sql

With rptLedger
        .DiscardSavedData = True
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "SaleReportCustomerbase.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = 'Customer Statement'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & dtpto & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
End With
MDIForm1.StatusBar1.Panels(7).Text = ""

Exit Sub
LocalErr:
Call SetErr(Err.Description, vbCritical)
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtselectiveaccount
    Set PO_DESC = txtcustomercode
    txtcustomercode = "Selective"
    txtcustomerdesc = "Selective Items"
    Gs_SQL = "SELECT ClientCode,Description FROM IC_Clients "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' "
    MyLookupOLDBSelective.txtsearchbase.Clear
    MyLookupOLDBSelective.txtsearchbase.AddItem "ClientCode"
    MyLookupOLDBSelective.txtsearchbase.AddItem "Description"
    MyLookupOLDBSelective.txtsearchbase.ListIndex = 1
    MyLookupOLDBSelective.Caption = "Customers"
    MyLookupOLDBSelective.Show 1
    
    
End Sub


Private Sub Command3_Click()
Unload Me
End Sub
Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdGenerate.SetFocus
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then dtpto.SetFocus
End Sub
Private Sub Form_Load()
dtpfrom = Date
dtpto = Date
chkcredit.Value = 1
End Sub


Private Sub txtcustomercode_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtcustomercode) <> "" And KeyCode = vbKeyReturn Then
        txtcustomercode = DoPad(txtcustomercode, 6)
        If pr_dumy.State = 1 Then pr_dumy.Close
        pr_dumy.Open "Select * from Ic_Clients where Compcode  = '" & Gs_compcode & "' and Clientcode = '" & txtcustomercode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Client Code not found !!!", vbCritical)
            txtcustomercode = ""
            txtcustomerdesc = ""
            txtcustomercode.SetFocus
            Exit Sub
        Else
            txtcustomerdesc = pr_dumy("Description")
            
        End If
        pr_dumy.Close
           

ElseIf Trim(txtcustomercode) = "" And KeyCode = vbKeyReturn Then
       txtcustomercode = ""
        txtcustomerdesc = ""
        Call Command1_Click
End If

End Sub

Private Sub txtLocCode_KeyDown(KeyCode As Integer, Shift As Integer)
If txtLocCode <> "" And KeyCode = vbKeyReturn Then
   txtLocCode = DoPad(txtLocCode, 10)
End If
End Sub

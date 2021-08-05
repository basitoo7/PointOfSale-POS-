VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPaymentReport 
   Caption         =   "Party Ledger"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4515
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
   Icon            =   "frmPaymentReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2220
      Left            =   0
      TabIndex        =   5
      Top             =   -45
      Width           =   4410
      Begin VB.ComboBox rpttype 
         Height          =   330
         ItemData        =   "frmPaymentReport.frx":030A
         Left            =   1110
         List            =   "frmPaymentReport.frx":0323
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1320
         Width           =   2370
      End
      Begin VB.TextBox txtpartydesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2115
         MaxLength       =   64
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   600
         Width           =   2190
      End
      Begin VB.ComboBox txttype 
         Height          =   330
         ItemData        =   "frmPaymentReport.frx":039A
         Left            =   1110
         List            =   "frmPaymentReport.frx":03B0
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   210
         Width           =   3195
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   405
         Left            =   2370
         TabIndex        =   9
         Top             =   1740
         Width           =   1035
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
         Left            =   1275
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Top             =   1740
         Width           =   1035
      End
      Begin VB.TextBox txtParty 
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
         Left            =   1110
         MaxLength       =   6
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   600
         Width           =   660
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   1785
         Picture         =   "frmPaymentReport.frx":0411
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   600
         Width           =   315
      End
      Begin Crystal.CrystalReport rptLedger 
         Left            =   3960
         Top             =   1725
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
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   1110
         TabIndex        =   2
         Top             =   960
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   54394881
         CurrentDate     =   37309
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Report Type :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   135
         TabIndex        =   13
         Top             =   1380
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type :"
         Height          =   210
         Index           =   8
         Left            =   630
         TabIndex        =   10
         Top             =   255
         Width           =   450
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "As On :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   60
         TabIndex        =   8
         Top             =   990
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Code :"
         Height          =   210
         Index           =   0
         Left            =   615
         TabIndex        =   7
         Top             =   615
         Width           =   465
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   2235
      Width           =   4515
      _ExtentX        =   7964
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
Attribute VB_Name = "frmPaymentReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PO_DESC As Object
Public PO_CODE As Object
Dim ls_CodeID As String
Dim Pr_ICParty As New Recordset


Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdGenerate_Click()
Dim ls_sql As String
Dim ls_Tag As String
Dim ls_Exp As String

ls_CodeID = IIf(txttype = "Distributors", "T", IIf(txttype = "Doctors /Medical Stores", "O", Left(txttype, 1)))

ls_Tag = IIf(txttype = "Supplier", "In ('G')", " In ('I','A')")
ls_Exp = " And  Issuetype = '" & ls_CodeID & "'"
ls_Exp = ls_Exp + " And  Compcode = '" & Gs_compcode & "'"
ls_Exp = ls_Exp + IIf(txtParty <> "", " And PartyCode = '" & txtParty & "'", "")

Call ChkTempTables("Tmp_PartyLedger", True)

Select Case UCase(rpttype)
Case UCase("Expired Quantity")
    
    ls_sql = "SELECT Compcode,Locationcode1,ItemClass,ItemCode,Batchno,ItemserialNo,Sum(case When Transtype = 'G' then Quantity else 0 end) as TRQty,Sum(case When Transtype = 'I' then Quantity else 0 end) as TIqty  Into Tmp_PartyLedger"
    ls_sql = ls_sql + " From IC_Trans where Expdate <= '" & Format(dtpfrom, "YYYY/MM/DD") & "' group by  Compcode,Locationcode1,ItemClass,ItemCode,BatchNo,ItemserialNo "
    gc_dbcon.Execute ls_sql
    
       With rptLedger
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "IC_ExpiredQty.Rpt"
        .SelectionFormula = "{@Onhandqty} > 0"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = 'Expired Stock Ledger'"
        .Formulas(2) = "Period = '" & "As On " & dtpfrom & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
       End With

Case UCase("Balance Quantity")
    Call ChkTempTables("Tmp_Balance", True)
    ls_sql = "SELECT Compcode,Locationcode1,ItemClass,ItemCode,Batchno,ItemserialNo,Sum(case When Transtype = 'G' then Quantity else 0 end) as TRQty,Sum(case When Transtype = 'I' then Quantity else 0 end) as TIqty ,Sum(case When Transtype = 'G' then Quantity*UnitCost else 0 end) as TVRQty,Sum(case When Transtype = 'I' then Quantity*UnitCost else 0 end) as TVIqty Into Tmp_PartyLedger"
    ls_sql = ls_sql + " From IC_Trans where Value_Date <= '" & Format(dtpfrom, "YYYY/MM/DD") & "' group by  Compcode,Locationcode1,ItemClass,ItemCode,BatchNo,ItemserialNo "
    gc_dbcon.Execute ls_sql
    
       With rptLedger
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "IC_QtyBalance.Rpt"
        .SelectionFormula = "{@Onhandqty} <> 0"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = 'Stock Ledger'"
        .Formulas(2) = "Period = '" & "As On " & dtpfrom & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
       End With

Case UCase("Balance Report")
    Call ChkTempTables("Tmp_Balance", True)
    ls_sql = "SELECT SupplierCode AS PartyCode,(-1)*(Quantity*UnitCost) as AAmount, ((-1)*(Quantity*UnitCost)* OtherExp/100) as OtherExp , ((-1)*(Quantity*UnitCost)*STaxRate/100) as STaxRate , 0 as Paidamt Into Tmp_PartyLedger"
    ls_sql = ls_sql + " From IC_Trans where TransType  " & ls_Tag & "" + Replace(ls_Exp, "PartyCode", "SupplierCode") + " And Value_Date <= '" & Format(dtpfrom, "YYYY/MM/DD") & "'"
    ls_sql = ls_sql + " Union All "
    ls_sql = ls_sql + " SELECT PartyCode,0 as AAmount, 0 as OtherExp, 0 as STaxRate , Amount as PaidAmt "
    ls_sql = ls_sql + " From IC_Payments where  TransDate <= '" & Format(dtpfrom, "YYYY/MM/DD") & "'" + Replace(ls_Exp, "Issuetype", "PartyType")
    gc_dbcon.Execute ls_sql
    
    ls_sql = "SELECT  PartyCode,sum(AAmount) as AAmount, sum(OtherExp) as OtherExp , sum(STaxRate) as STaxRate , sum(Paidamt) as PaidAmt  Into Tmp_Balance from Tmp_PartyLedger group by partycode "
    gc_dbcon.Execute ls_sql
    gc_dbcon.Execute "Drop Table Tmp_PartyLedger;"
    
       With rptLedger
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "IC_Balance.Rpt"
        .SelectionFormula = "{@Balance} <> 0"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = '" & txttype + " Ledger" & "'"
        .Formulas(2) = "Period = '" & "As On " & dtpfrom & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
       End With
Case UCase("Detail Report")
    ls_sql = "SELECT Compcode,TransType as Transid, SupplierCode AS PartyCode, IssueType AS PartyType, Transc_No, Value_Date, NULL AS PaymentMode, NULL AS BankCode, NULL AS Instrument, "
    ls_sql = ls_sql + " 0 AS Amount, Quantity, UnitCost, OtherExp, STaxRate,Batchno,ItemSerialNo,Locationcode1,ItemClass,ItemCode  Into Tmp_PartyLedger"
    ls_sql = ls_sql + " From IC_Trans where TransType   " & ls_Tag & "" + Replace(ls_Exp, "PartyCode", "SupplierCode") + " And Value_Date <= '" & Format(dtpfrom, "YYYY/MM/DD") & "'"
    ls_sql = ls_sql + " Union All "
    ls_sql = ls_sql + " SELECT Compcode, 'P' as Transid,PartyCode, PartyType, TransCode, TransDate, PaymentMode, BankCode, InstrumentNo, Amount, 0 AS Quantity, 0 AS UnitCost, 0 as OtherExp, 0 as STaxRate,Null as Batchno,Null as ItemSerialNo,Null as Locationcode1,Null as ItemClass,Null as ItemCode "
    ls_sql = ls_sql + " From IC_Payments where  TransDate <= '" & Format(dtpfrom, "YYYY/MM/DD") & "'" + Replace(ls_Exp, "Issuetype", "PartyType")
    gc_dbcon.Execute ls_sql
    
       With rptLedger
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "PartyLedger.Rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = '" & txttype + " Ledger" & "'"
        .Formulas(2) = "Period = '" & "As On " & dtpfrom & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
       End With
Case UCase("Payments Info.")
    
    ls_sql = ls_sql + " SELECT PartyCode, PartyType, TransCode as Transc_No , TransDate, PaymentMode, BankCode, InstrumentNo, Amount, 0 AS Quantity, 0 AS UnitCost"
    ls_sql = ls_sql + " Into Tmp_PartyLedger From IC_Payments where  TransDate <= '" & Format(dtpfrom, "YYYY/MM/DD") & "'" + Replace(ls_Exp, "Issuetype", "PartyType")
    gc_dbcon.Execute ls_sql
    
       With rptLedger
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "PartyCLedger.Rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = '" & txttype + " Ledger" & "'"
        .Formulas(2) = "Period = '" & "As On " & dtpfrom & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
       End With
 Case UCase("Products Info.")
    
        ls_sql = "SELECT Locationcode1,ItemClass,ItemCode,SupplierCode AS PartyCode, IssueType AS PartyType, Transc_No, Value_Date, Quantity, UnitCost, OtherExp, STaxRate,BatchNo,ItemSerialNo Into Tmp_PartyLedger"
        ls_sql = ls_sql + " From IC_Trans where TransType  " & ls_Tag & "" + Replace(ls_Exp, "PartyCode", "SupplierCode") + " And Value_Date <= '" & Format(dtpfrom, "YYYY/MM/DD") & "'"
        gc_dbcon.Execute ls_sql
        
       With rptLedger
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "PartyPLedger.Rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = '" & txttype + " Ledger" & "'"
        .Formulas(2) = "Period = '" & "As On " & dtpfrom & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
       End With
 Case UCase("Personal Info.")
       With rptLedger
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "IC_Parties.Rpt"
        .SelectionFormula = "{IC_Supplier.CodeId} = '" & ls_CodeID & "'"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = '" & txttype + " Personal Info" & "'"
        .Formulas(2) = "Period = '" & "As On " & dtpfrom & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
       End With
 
 End Select
 
End Sub

Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtParty
    Set PO_DESC = txtpartydesc
    
    ls_CodeID = IIf(txttype = "Distributors", "T", IIf(txttype = "Doctors /Medical Stores", "O", Left(txttype, 1)))
    
    Pr_ICParty.Filter = "CodeId= '" & ls_CodeID & "'"
    
    GoTop Pr_ICParty
    MyLookup.Caption = "Jobs "
    MyLookup.FillGrid Pr_ICParty, "SupplierCode", "Description", 6
    MyLookup.Show 1
    Pr_ICParty.Filter = adFilterNone
    If Len(txtParty) > 0 Then txtparty_KeyDown vbKeyReturn, vbKeyShift
    
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdGenerate.SetFocus
End Sub

Private Sub Form_Load()
dtpfrom.Value = Date
rpttype = "Detail Report"
txttype = "Customer"
Pr_ICParty.Open "Select * from IC_Supplier where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Pr_ICParty.Close
End Sub

Private Sub txtparty_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And txtParty.Text <> "" Then
        
             txtParty.Text = DoPad(txtParty.Text, txtParty.MaxLength)
             ls_CodeID = IIf(txttype = "Distributors", "T", IIf(txttype = "Doctors /Medical Stores", "O", Left(txttype, 1)))
             Pr_ICParty.Filter = "CodeId= '" & ls_CodeID & "'"
             If Not MySeek(txtParty.Text, "SupplierCode", Pr_ICParty) Then
                    Call SetErr(Gs_RecNFMsg, vbCritical)
                    txtParty.SetFocus
                    txtpartydesc.Text = ""
                Else
                    txtpartydesc.Text = Pr_ICParty("Description")
                    dtpfrom.SetFocus
              End If
                Pr_ICParty.Filter = adFilterNone
ElseIf KeyCode = vbKeyF12 Then
    Call Command2_Click
End If
End Sub

Private Sub txttype_Click()
If txttype = "Department" Then rpttype = "Products Info."
End Sub

Private Sub txttype_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtParty.SetFocus
End Sub


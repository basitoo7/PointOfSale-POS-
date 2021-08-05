VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransLedger 
   Caption         =   "Transaction Ledger"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2550
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
   Icon            =   "frmTransLedger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   2550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   30
      TabIndex        =   1
      Top             =   -75
      Width           =   2520
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
         Left            =   165
         MaskColor       =   &H00000000&
         TabIndex        =   14
         Top             =   3030
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   405
         Left            =   1320
         TabIndex        =   13
         Top             =   3030
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
         TabIndex        =   8
         Top             =   135
         Width           =   2370
         Begin MSComCtl2.DTPicker dtpto 
            Height          =   315
            Left            =   1035
            TabIndex        =   9
            Top             =   645
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            Format          =   22806529
            CurrentDate     =   37309
         End
         Begin MSComCtl2.DTPicker dtpfrom 
            Height          =   315
            Left            =   1035
            TabIndex        =   10
            Top             =   210
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            Format          =   22806529
            CurrentDate     =   37309
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "From :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   540
            TabIndex        =   12
            Top             =   225
            Width           =   450
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "To :"
            Height          =   210
            Left            =   720
            TabIndex        =   11
            Top             =   690
            Width           =   270
         End
      End
      Begin VB.TextBox txtAcctNarration 
         Height          =   315
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   7
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
         Height          =   1005
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
            Left            =   1050
            MaxLength       =   2
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   180
            Width           =   600
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   1665
            Picture         =   "frmTransLedger.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   165
            Width           =   315
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
            Left            =   1650
            Picture         =   "frmTransLedger.frx":047C
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   585
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
            Height          =   330
            Left            =   1050
            MaxLength       =   5
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   570
            Width           =   570
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
            Caption         =   "Location :"
            Height          =   210
            Left            =   300
            TabIndex        =   17
            Top             =   210
            Width           =   705
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Item Name :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   195
            TabIndex        =   6
            Top             =   615
            Width           =   825
         End
      End
      Begin Crystal.CrystalReport rptLedger 
         Left            =   2430
         Top             =   1575
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
         Height          =   885
         Left            =   90
         TabIndex        =   18
         Top             =   2115
         Width           =   2370
         Begin VB.CheckBox opbal 
            Alignment       =   1  'Right Justify
            Caption         =   "Open Balance :"
            Height          =   240
            Left            =   75
            TabIndex        =   21
            Top             =   585
            Width           =   1410
         End
         Begin VB.ComboBox txtledgertype 
            Height          =   330
            ItemData        =   "frmTransLedger.frx":05EE
            Left            =   1065
            List            =   "frmTransLedger.frx":05FB
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   210
            Width           =   1245
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ledger Type :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   60
            TabIndex        =   19
            Top             =   255
            Width           =   1005
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3465
      Width           =   2550
      _ExtentX        =   4498
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
Attribute VB_Name = "frmTransLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pb_BlnkVchr As Boolean
Dim Mode As String
Dim PR_Item As New Recordset
Dim lb_found As Boolean
Public PO_DESC As Object
Public PO_CODE As Object
Dim pi_Event As Integer
Dim PR_ICItmLoc As New Recordset
Dim ls_ItemClass As String


Private Sub cmdCancel_Click()
    pi_Event = 0
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
On Error GoTo LocalErr
Select Case Left(Me.Caption, 1)
  Case "T"
   With rptLedger
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & Me.Tag
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & dtpto & "'"
        .SelectionFormula = "{IC_Trans.Value_Date} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {IC_Trans.Value_Date} <= Date(" & dtpto.Year & "," & dtpto.Month & "," & dtpto.Day & ") "
        .SelectionFormula = .SelectionFormula & "  and {Ic_Trans.TransType} = '" & Left(Me.Tag, 1) & "'"
        If Trim(Len(txtFrom)) > 0 Then .SelectionFormula = .SelectionFormula & " AND {IC_Trans.LocationCode1}+{IC_Trans.ItemClass}+ {IC_Trans.ItemCode} ='" & txtFrom & "'"
        .Action = 1
   End With
  Case "S"
     Call ProssStkLedg
   With rptLedger
        .SelectionFormula = ""
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\BinCard.Rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & dtpto & "'"
        '.SelectionFormula = "{IC_Trans.Value_Date} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {IC_Trans.Value_Date} <= Date(" & dtpto.Year & "," & dtpto.Month & "," & dtpto.Day & ")"
        ' If Trim(Len(txtFrom)) > 0 Then .SelectionFormula = .SelectionFormula & " AND {IC_Trans.ItemClass} = '" & txtItemClass & "' AND {IC_Trans.ItemCode} ='" & TxtItemCode & "' AND {IC_Trans.LocationCode1} ='" & txtLocationCode & "'"
        .Action = 1
   End With
  gc_dbcon.Execute ("Drop Table Tmp_IcTrans;")
  
  Case "D"
   With rptLedger
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\DeptCons.Rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & dtpto & "'"
        .SelectionFormula = "{IC_Trans.Value_Date} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {IC_Trans.Value_Date} <= Date(" & dtpto.Year & "," & dtpto.Month & "," & dtpto.Day & ")"
        If Trim(Len(txtFrom)) > 0 Then .SelectionFormula = .SelectionFormula & " AND {IC_Supplier.CodeID} = 'D'"
        .Action = 1
   End With
  Case "J"
   With rptLedger
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\JobCode.Rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & dtpto & "'"
        .SelectionFormula = "{IC_Trans.Value_Date} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {IC_Trans.Value_Date} <= Date(" & dtpto.Year & "," & dtpto.Month & "," & dtpto.Day & ")"
        If Trim(Len(txtFrom)) > 0 Then .SelectionFormula = .SelectionFormula & " AND {IC_Trans.JobCode} = '" & txtFrom & "'"
        .Action = 1
   End With
End Select

Exit Sub
LocalErr:
Call SetErr(Err.Description, vbCritical)
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
    'gc_dbcon.Execute ("SELECT CompCode, SUM(Quantity * UnitCost) AS OB into Tmp_TrnsOB from IC_Trans  WHERE Value_Date >= '" & Format(DateValue(Gs_Fnperiod), "YYYY/MM/DD") & "' and Value_Date < '" & Format(dtpFrom, "YYYY/MM/DD") & "'" & ls_VchrType & "  GROUP BY CompCode")
End Sub
Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtLocCode
    Set PO_DESC = Text1
    GoTop PR_ICItmLoc
    MyLookup.Caption = "Items. "
    MyLookup.FillGrid PR_ICItmLoc, "LocationCode", "Description", 2
    MyLookup.Show 1
    
    If Len(txtLocCode) > 0 Then txtLocCode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub txtLocCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 
 If KeyCode = vbKeyReturn And Len(txtLocCode.Text) > 0 Then
         txtLocCode.Text = DoPad(txtLocCode.Text, txtLocCode.MaxLength)
         lb_found = MySeek(txtLocCode.Text, "LocationCode", PR_ICItmLoc)
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtLocCode.SetFocus
             'txtLocDesc.Text = ""
         Else
         Text1.Text = PR_ICItmLoc("Description")
             txtFrom.SetFocus
         End If
 ElseIf KeyCode = vbKeyF12 Then
        Command2_Click
 End If
End Sub


Private Sub Command3_Click()
Dim ln_len As Integer
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtFrom
    Set PO_DESC = Text1
    
    GoTop PR_Item
    
    If PR_Item.EOF Then
      ln_len = 10
    Else
      ln_len = Len(Trim(RTrim(PR_Item.Fields("ItemID"))))
    End If
    PR_Item.Filter = "LocationCode = '" & txtLocCode & "'"
    MyLookup.Caption = "Items"
    MyLookup.FillGrid PR_Item, "Itemid", "Descr", 5
    MyLookup.Show 1
    PR_Item.Filter = adFilterNone
    If Len(txtFrom) > 0 Then txtFrom_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub dtpfrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If Lastkey(KeyCode) Then
       dtpto.SetFocus
    End If
End Sub

Private Sub dtpto_KeyDown(KeyCode As Integer, Shift As Integer)
    If Lastkey(KeyCode) And InStr(1, Me.Caption, "Ledger") > 0 Then
       txtFrom.SetFocus
    End If
End Sub

Private Sub Form_Activate()
  If pi_Event = 1 Then Exit Sub
  Select Case Left(Me.Caption, 1)
      Case "T", "S"
          PR_Item.Open "Select *,LTrim(RTrim(locationcode))+LTrim(RTrim(ItemCode)) AS ItemID, Description AS Descr from IC_Item where CompCode ='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly
      Case "D"
          PR_Item.Open "Select SupplierCode As ItemCode, SupplierCode AS ItemID, Description AS Descr from IC_Supplier where CodeId = 'D' ", gc_dbcon, adOpenStatic, adLockReadOnly
      Case "J"
          PR_Item.Open "Select JobCode As ItemCode, JobCode AS ItemID, Description AS Descr from IC_Job where CompCode ='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly
  End Select
  Pb_BlnkVchr = IIf(PR_Item.EOF, True, False)
  pi_Event = 1
End Sub

Private Sub Form_Load()
dtpfrom.Value = Date
dtpto.Value = Date
PR_ICItmLoc.Open "Select * from Ic_Locations where compcode ='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_Item.Close
    PR_ICItmLoc.Close
    Set frmTransLedger = Nothing
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtFrom.Text <> "" Then
        PR_Item.Filter = "LocationCode = '" & txtLocCode & "'"
        txtFrom.Text = DoPad(txtFrom, txtFrom.MaxLength)
        lb_found = MySeek(txtFrom.Text, "ItemId", PR_Item)
        If lb_found Then
          If Left(Me.Caption, 1) = "T" Or Left(Me.Caption, 1) = "S" Then
            StatusBar1.Panels(2).Text = PR_Item("Descr")
            cmdGenerate.SetFocus
          End If
        Else
            Call SetErr("Record not found", vbCritical)
        End If
        PR_Item.Filter = adFilterNone
End If
End Sub

Private Sub ProssStkLedg()
Dim ls_sql As String
Dim ls_ledgertype As String
Call Module1.ChkTempTables("Tmp_ICTrans", True)
    
   If txtLocCode.Text <> "" And txtFrom.Text <> "" Then
        ls_Item = " and (ltrim(rtrim(Ic_Trans.LocationCode1))+(rtrim(Ic_Trans.ItemCode))) = '" & Trim(txtFrom.Text) & "'"
    ElseIf txtLocCode.Text <> "" Then
        ls_Item = " and ltrim(rtrim(Ic_Trans.LocationCode1)) = '" & Trim(txtLocCode) & "'"
    End If
        If txtledgertype = "Receipts" Then
            ls_ledgertype = "And TransType = 'G'"
        ElseIf txtledgertype = "Issue" Then
            ls_ledgertype = "And TransType = 'I'"
        Else
            ls_ledgertype = ""
        End If
        If opbal.Value = 1 Then
            ls_sql = "Select Compcode,Null as Transc_No,Null as TransType,'" & Format(dtpfrom, "YYYY/MM/DD") & "' as Value_date, locationcode1,ItemClass,Itemcode,Null as IssueType,null as ItemSerialNo,null as batchno, SUM(Quantity) as Quantity,(Sum(Quantity*UnitCost)/SUM(Quantity)) as UnitCost, 'Opening Balance' as Remarks,Null as SupplierCode ,'A' as Transid Into Tmp_ICTrans "
            ls_sql = ls_sql + "From Ic_Trans where Value_Date >= '" & Format(DateValue(Gs_Fnperiod), "YYYY/MM/DD") & "' and  Value_Date < '" & Format(dtpfrom, "YYYY/MM/DD") & "'"
            ls_sql = ls_sql + " And compcode = '" & Gs_compcode & "'" + ls_ledgertype + ls_Item + " Group By Compcode,Locationcode1,ItemClass,ItemCode "
            
            ls_sql = ls_sql + "Union All "
            ls_sql = ls_sql + "Select CompCode, Transc_No,TransType,Value_date,Locationcode1,ItemClass,Itemcode,IssueType,ItemSerialNo,batchno,Quantity,UnitCost,Remarks,SupplierCode, 'B' as TransId "
            ls_sql = ls_sql + "From Ic_Trans where Value_Date >= '" & Format(dtpfrom.Value, "YYYY/MM/DD") & "' and  Value_Date <= '" & Format(dtpto.Value, "YYYY/MM/DD") & "' "
            ls_sql = ls_sql + "and compcode = '" & Gs_compcode & "'" + ls_ledgertype + ls_Item
        Else
            ls_sql = "Select CompCode, Transc_No,TransType,Value_date,Locationcode1,ItemClass,Itemcode,IssueType,ItemSerialNo,batchno,Quantity,UnitCost,Remarks,SupplierCode, 'B' as TransId  Into Tmp_ICTrans "
            ls_sql = ls_sql + "From Ic_Trans where Value_Date >= '" & Format(dtpfrom.Value, "YYYY/MM/DD") & "' and  Value_Date <= '" & Format(dtpto.Value, "YYYY/MM/DD") & "' "
            ls_sql = ls_sql + "and compcode = '" & Gs_compcode & "'" + ls_ledgertype + ls_Item
        End If
           gc_dbcon.Execute (ls_sql)
            
            
End Sub

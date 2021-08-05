VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransLedger1 
   Caption         =   "Transaction Ledger"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2850
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
   Icon            =   "frmTransLedger1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   2850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2010
      Left            =   30
      TabIndex        =   1
      Top             =   -75
      Width           =   2775
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
         Left            =   240
         MaskColor       =   &H00000000&
         TabIndex        =   9
         Top             =   1530
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   405
         Left            =   1395
         TabIndex        =   8
         Top             =   1530
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
         Height          =   1365
         Left            =   90
         TabIndex        =   3
         Top             =   135
         Width           =   2625
         Begin VB.ComboBox Combo1 
            Height          =   330
            ItemData        =   "frmTransLedger1.frx":030A
            Left            =   1050
            List            =   "frmTransLedger1.frx":031A
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   945
            Width           =   1485
         End
         Begin MSComCtl2.DTPicker dtpto 
            Height          =   315
            Left            =   1035
            TabIndex        =   4
            Top             =   585
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            Format          =   63504385
            CurrentDate     =   37309
         End
         Begin MSComCtl2.DTPicker dtpfrom 
            Height          =   315
            Left            =   1035
            TabIndex        =   5
            Top             =   210
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            Format          =   63504385
            CurrentDate     =   37309
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "To Date :"
            Height          =   210
            Left            =   345
            TabIndex        =   11
            Top             =   960
            Width           =   645
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "From Date :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   180
            TabIndex        =   7
            Top             =   225
            Width           =   825
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "To Date :"
            Height          =   210
            Left            =   345
            TabIndex        =   6
            Top             =   585
            Width           =   645
         End
      End
      Begin VB.TextBox txtAcctNarration 
         Height          =   315
         Left            =   3090
         MaxLength       =   50
         TabIndex        =   2
         Top             =   555
         Visible         =   0   'False
         Width           =   315
      End
      Begin Crystal.CrystalReport rptLedger 
         Left            =   150
         Top             =   675
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
      Top             =   1965
      Width           =   2850
      _ExtentX        =   5027
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
Attribute VB_Name = "frmTransLedger1"
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
Unload Me
End Sub

Private Sub cmdGenerate_Click()
   With rptLedger
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\stockledger.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Period = '" & "From " & dtpFrom & " to " & dtpTo & "'"
        .Formulas(3) = "Fromdate = Date(" & dtpFrom.Year & "," & dtpFrom.Month & "," & dtpFrom.Day & ")"
        
        .SelectionFormula = "{IC_Trans.Value_Date} >= Date(" & dtpFrom.Year & "," & dtpFrom.Month & "," & dtpFrom.Day & ") AND {IC_Trans.Value_Date} <= Date(" & dtpTo.Year & "," & dtpTo.Month & "," & dtpTo.Day & ") "
         If txtLocCode <> "" Then .SelectionFormula = .SelectionFormula & " and {IC_Trans.locationcode} = '" & txtLocCode & "'"
         If txtitemcode <> "" Then .SelectionFormula = .SelectionFormula & " and  {IC_Trans.itemcode} = '" & txtitemcode & "'"
        .Action = 1
   End With

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
             
         Else
         txtlocationdesc.Text = PR_ICItmLoc("Description")
             txtitemcode.SetFocus
         End If
 ElseIf KeyCode = vbKeyF12 Then
        Command2_Click
 End If
End Sub


Private Sub Command3_Click()
Dim ln_len As Integer
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtitemcode
    Set PO_DESC = TxtItemdesc
    
    GoTop PR_Item
    PR_Item.Filter = "LocationCode = '" & txtLocCode & "'"
    MyLookup.Caption = "Items"
    MyLookup.FillGrid PR_Item, "Itemcode", "Description", 5
    MyLookup.Show 1
    PR_Item.Filter = adFilterNone
    If Len(txtitemcode) > 0 Then txtitemcode_KeyDown vbKeyReturn, vbKeyShift

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
dtpFrom.Value = Date
dtpTo.Value = Date
PR_ICItmLoc.Open "Select * from Ic_Locations where compcode ='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_Item.Close
    PR_ICItmLoc.Close
    Set frmTransLedger = Nothing
End Sub

Private Sub txtitemcode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtitemcode.Text <> "" Then
        PR_Item.Filter = "LocationCode = '" & txtLocCode & "'"
        txtitemcode.Text = DoPad(txtitemcode, txtitemcode.MaxLength)
        lb_found = MySeek(txtitemcode.Text, "Itemcode", PR_Item)
        If lb_found Then
            TxtItemdesc.Text = PR_Item("description")
            cmdGenerate.SetFocus
        Else
            Call SetErr("Record not found", vbCritical)
        End If
        PR_Item.Filter = adFilterNone
End If
End Sub

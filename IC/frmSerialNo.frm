VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSerialNo 
   Caption         =   "Search Item Serial #"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3165
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
   Icon            =   "frmSerialNo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3165
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   3165
      _ExtentX        =   5583
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
      Height          =   1635
      Left            =   30
      TabIndex        =   1
      Top             =   -90
      Width           =   3075
      Begin VB.ComboBox txtledgertype 
         Height          =   330
         ItemData        =   "frmSerialNo.frx":030A
         Left            =   1020
         List            =   "frmSerialNo.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   780
         Width           =   1590
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
         Left            =   450
         MaskColor       =   &H00000000&
         TabIndex        =   4
         Top             =   1155
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   1575
         TabIndex        =   3
         Top             =   1155
         Width           =   1035
      End
      Begin VB.TextBox txtAcctNarration 
         Height          =   315
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   2
         Top             =   825
         Visible         =   0   'False
         Width           =   315
      End
      Begin Crystal.CrystalReport rptLedger 
         Left            =   0
         Top             =   1470
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
         Left            =   30
         TabIndex        =   7
         Top             =   135
         Width           =   3015
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   -240
            MaxLength       =   50
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   435
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   2595
            Picture         =   "frmSerialNo.frx":0345
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   165
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
            Left            =   990
            MaxLength       =   10
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   180
            Width           =   1575
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Item Serial #:"
            Height          =   210
            Left            =   60
            TabIndex        =   11
            Top             =   210
            Width           =   915
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   210
         Left            =   585
         TabIndex        =   6
         Top             =   825
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmSerialNo"
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
On Error GoTo LocalErr
If txtledgertype = "Receipts" Then
            ls_ledgertype = "And TransType = 'G'"
        ElseIf txtledgertype = "Issue" Then
            ls_ledgertype = "And TransType = 'I'"
        Else
            ls_ledgertype = ""
        End If
   With rptLedger
        .DiscardSavedData = True
        .WindowTitle = Me.Caption
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Item Serial No'"
        .Formulas(3) = "Sign1 = '" & Gc_UserName & "'"
        .SelectionFormula = "{IC_Trans.itemserialno}= '" & txtLocCode & "'"
        If txtledgertype = "Receipts" Then
        .SelectionFormula = .SelectionFormula & " And {IC_Trans.TransType}= 'G'"
        ElseIf txtledgertype = "Issue" Then
        .SelectionFormula = .SelectionFormula & " And {IC_Trans.TransType}= 'I'"
        ElseIf txtledgertype = "Transfer" Then
        .SelectionFormula = .SelectionFormula & " And {IC_Trans.TransType}= 'T'"
        ElseIf txtledgertype = "Adjustment" Then
        .SelectionFormula = .SelectionFormula & " And {IC_Trans.TransType}= 'A'"
        End If
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "itemserialno.rpt"
        .Action = 1
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
    Gs_SQL = "SELECT    IC_Trans.ItemSerialNo as 'Serial No', IC_Supplier.Description as 'Party Description'  FROM    IC_Trans INNER JOIN  "
    Gs_SQL = Gs_SQL & " IC_Supplier ON IC_Trans.CompCode = IC_Supplier.Compcode AND IC_Trans.SupplierCode = IC_Supplier.SupplierCode "
    Gs_FindFld = " IC_Trans.ItemSerialNo"
    Gs_Subon = False
    Gs_OtherPara = " Where IC_Trans.Compcode = '" & Gs_compcode & "'"
    Gs_OrderBy = "Order by IC_Trans.ItemSerialNo"
    MyLookupOLDB.Caption = "Search Item Serial No"
    MyLookupOLDB.Show 1
    
    If Len(txtLocCode) > 0 Then txtLocCode_KeyDown vbKeyReturn, vbKeyShift
    
End Sub

Private Sub dtpto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtLocCode.SetFocus
End Sub

Private Sub txtLocCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 If KeyCode = vbKeyReturn And Len(txtLocCode.Text) > 0 Then
         lb_found = MySeek(txtLocCode.Text, "Itemserialno", PR_ICTran)
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtLocCode.SetFocus
             'txtLocDesc.Text = ""
         Else
            Text1.Text = PR_ICTran("itemserialno")
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
PR_ICTran.Open "Select *  from Ic_Trans where compcode ='" & Gs_compcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_ICTran.Close
End Sub



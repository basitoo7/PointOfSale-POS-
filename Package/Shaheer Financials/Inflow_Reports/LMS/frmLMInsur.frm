VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLMInsur 
   Caption         =   "Insurance"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLMInsur.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   3240
      Left            =   0
      TabIndex        =   1
      Top             =   570
      Width           =   5460
      Begin VB.TextBox txtRemarks 
         Height          =   525
         Left            =   1155
         MaxLength       =   255
         TabIndex        =   26
         Top             =   2655
         Width           =   4245
      End
      Begin VB.CommandButton Command1 
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
         Left            =   1965
         Picture         =   "frmLMInsur.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "SKIP"
         Top             =   1275
         Width           =   315
      End
      Begin VB.TextBox txtInsurCompCode 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1170
         MaxLength       =   3
         TabIndex        =   21
         Tag             =   "SKIP"
         Top             =   1275
         Width           =   765
      End
      Begin VB.ComboBox cmbInsurType 
         Height          =   330
         ItemData        =   "frmLMInsur.frx":047C
         Left            =   1170
         List            =   "frmLMInsur.frx":0486
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   210
         Width           =   2685
      End
      Begin VB.TextBox txtleaseno 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1170
         MaxLength       =   3
         TabIndex        =   6
         Tag             =   "SKIP"
         Top             =   930
         Width           =   765
      End
      Begin VB.CommandButton Command2 
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
         Left            =   1965
         Picture         =   "frmLMInsur.frx":04B5
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "SKIP"
         Top             =   930
         Width           =   315
      End
      Begin VB.CommandButton cmdlookup2 
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
         Left            =   1965
         Picture         =   "frmLMInsur.frx":0627
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   570
         Width           =   315
      End
      Begin VB.TextBox txtCustomerName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2325
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   570
         Width           =   3090
      End
      Begin MSComCtl2.DTPicker dtpDOIssue 
         Height          =   315
         Left            =   1155
         TabIndex        =   2
         Top             =   1965
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57278465
         CurrentDate     =   37293
      End
      Begin MSMask.MaskEdBox txtCustomerNo 
         Height          =   315
         Left            =   1170
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   570
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtAmount 
         Height          =   315
         Left            =   1155
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2310
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtLeaseName 
         Height          =   315
         Left            =   2325
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   915
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483644
         PromptInclude   =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtpDOExpiry 
         Height          =   315
         Left            =   4245
         TabIndex        =   17
         Top             =   1935
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57278465
         CurrentDate     =   37293
      End
      Begin MSMask.MaskEdBox txtInsurCode 
         Height          =   315
         Left            =   1170
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1620
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtInsurCompName 
         Height          =   315
         Left            =   2325
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1260
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483644
         PromptInclude   =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Remarks :"
         Height          =   210
         Left            =   390
         TabIndex        =   25
         Top             =   2685
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Insur. Comp :"
         Height          =   210
         Index           =   2
         Left            =   195
         TabIndex        =   23
         Top             =   1275
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Insur. Type :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   19
         Top             =   210
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Expiry Date :"
         Height          =   210
         Left            =   3285
         TabIndex        =   16
         Top             =   1950
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Amount :"
         Height          =   210
         Left            =   465
         TabIndex        =   14
         Top             =   2325
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Account # :"
         Height          =   210
         Index           =   0
         Left            =   300
         TabIndex        =   12
         Top             =   930
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Customer # :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   225
         TabIndex        =   11
         Top             =   555
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Insur.Policy # :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   90
         TabIndex        =   10
         Top             =   1635
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Issue Date :"
         Height          =   210
         Left            =   255
         TabIndex        =   9
         Top             =   1980
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Value Date :"
         Height          =   210
         Index           =   1
         Left            =   3075
         TabIndex        =   8
         Top             =   600
         Width           =   885
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   1005
      ButtonWidth     =   1376
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&New"
            Key             =   "Ctrl + N"
            Description     =   "Add"
            Object.ToolTipText     =   "Add new record"
            Object.Tag             =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Edit"
            Key             =   "Ctrl + E"
            Description     =   "Edit"
            Object.ToolTipText     =   "Edit an existing record"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Delete"
            Key             =   "Ctrl + D"
            Description     =   "Remove "
            Object.ToolTipText     =   "Remove an existing record."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Save"
            Key             =   "Shift+S"
            Description     =   "Save a new Record"
            Object.ToolTipText     =   "Save on disk"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Listing"
            Description     =   "Print Listing."
            Object.ToolTipText     =   "Print listing."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Re&fresh"
            Description     =   "Find a Record."
            Object.ToolTipText     =   "Find a record."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancel"
            Description     =   "Cancel Operation"
            Object.ToolTipText     =   "Cancel operation mode"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   14
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4710
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLMInsur.frx":0799
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLMInsur.frx":0BED
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLMInsur.frx":1041
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLMInsur.frx":1495
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLMInsur.frx":18E9
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLMInsur.frx":1D3D
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLMInsur.frx":2491
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmLMInsur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IF_BlnkInflow As Boolean
Dim Mode As String
Dim Is_CustomerNo As String
Dim aa As String

Public PO_CODE As Object
Public PO_DESC As Object

Dim PR_LMInsurance As New Recordset
Dim pr_Customer As New Recordset
Dim PR_AccountNo As New Recordset
Dim PR_InsurComp As New Recordset
Dim InsuranceType As String

Private Sub cmbInsurType_Click()
    Dim ls_Sql As String
    
    txtCustomerNo = ""
    txtCustomerName = ""
    txtleaseno = ""
    txtLeaseName = ""
    
    If cmbInsurType.ListIndex = 0 Then
        InsuranceType = "L"
    ElseIf cmbInsurType.ListIndex = 1 Then
        InsuranceType = "A"
    End If
    
    If cmbInsurType.ListIndex = 0 Then
        ls_Sql = "SELECT DISTINCT dbo.LM_LeaseInfo.CustomerNo, dbo.Customer.CustomerName"
        ls_Sql = ls_Sql & " FROM dbo.LM_LeaseInfo INNER JOIN"
        ls_Sql = ls_Sql & " dbo.Customer ON dbo.LM_LeaseInfo.Compcode = dbo.Customer.Compcode "
        ls_Sql = ls_Sql & " AND dbo.LM_LeaseInfo.BranchCode = dbo.Customer.BranchCode AND"
        ls_Sql = ls_Sql & " dbo.LM_LeaseInfo.CustomerNo = dbo.Customer.CustomerNo"
        ls_Sql = ls_Sql & " WHERE LM_LeaseInfo.CompCode + LM_LeaseInfo.BranchCode = '" & Gs_compcode + Gs_BranchCode & "' ORDER BY LM_LeaseInfo.CustomerNo"
        If pr_Customer.State = 1 Then pr_Customer.Close
        pr_Customer.Open ls_Sql, gc_dbcon, adOpenStatic, adLockReadOnly
        
        ls_Sql = "SELECT DISTINCT LM_LeaseInfo.CustomerNo + LM_LeaseInfo.LeaseNo AS FindFld, LM_LeaseInfo.CustomerNo, LM_LeaseInfo.LeaseNo, LM_LeaseInfo.LeaseRental FROM LM_LeaseInfo WHERE LM_LeaseInfo.CompCode + LM_LeaseInfo.BranchCode = '" & Gs_compcode + Gs_BranchCode & "' ORDER BY LM_LeaseInfo.LeaseNo"
        If PR_AccountNo.State = 1 Then PR_AccountNo.Close
        PR_AccountNo.Open ls_Sql, gc_dbcon, adOpenStatic, adLockReadOnly
    Else
        ls_Sql = "SELECT DISTINCT dbo.ADV_Master.CustomerNo, dbo.Customer.CustomerName"
        ls_Sql = ls_Sql & " FROM dbo.Customer INNER JOIN"
        ls_Sql = ls_Sql & " dbo.ADV_Master ON dbo.Customer.Compcode = dbo.ADV_Master.Compcode "
        ls_Sql = ls_Sql & " AND dbo.Customer.BranchCode = dbo.ADV_Master.BranchCode AND"
        ls_Sql = ls_Sql & " dbo.Customer.CustomerNo = dbo.ADV_Master.CustomerNo"
        ls_Sql = ls_Sql & " WHERE ADV_Master.AgrCalcType = 'A' AND ADV_Master.CompCode + ADV_Master.BranchCode = '" & Gs_compcode + Gs_BranchCode & "' ORDER BY ADV_Master.CustomerNo"
        If pr_Customer.State = 1 Then pr_Customer.Close
        pr_Customer.Open ls_Sql, gc_dbcon, adOpenStatic, adLockReadOnly
        
        ls_Sql = "SELECT DISTINCT ADV_Master.CustomerNo + ADV_Master.AccountNo AS FindFld, ADV_Master.CustomerNo, ADV_Master.AccountNo AS LeaseNo, '' AS LeaseRental FROM ADV_Master WHERE ADV_Master.AgrCalcType = 'A' AND ADV_Master.CompCode + ADV_Master.BranchCode = '" & Gs_compcode + Gs_BranchCode & "' ORDER BY ADV_Master.AccountNo"
        If PR_AccountNo.State = 1 Then PR_AccountNo.Close
        PR_AccountNo.Open ls_Sql, gc_dbcon, adOpenStatic, adLockReadOnly
    End If
End Sub

Private Sub cmdlookup2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCustomerNo
    Set PO_DESC = txtCustomerName
    
    GoTop pr_Customer
    MyLookup.Caption = "Customers"
    MyLookup.FillGrid pr_Customer, "CustomerNo", "CustomerName", txtCustomerNo.MaxLength
    MyLookup.Show 1
    
    If Len(txtCustomerNo) > 0 Then txtCustomerNo_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtInsurCompCode
    Set PO_DESC = txtInsurCompName

    GoTop PR_InsurComp
    MyLookup.Caption = "Insurance Companies"
    MyLookup.FillGrid PR_InsurComp, "InsurCode", "InsurDesc", txtInsurCompCode.MaxLength + 2
    MyLookup.Show 1
    If Len(txtPmtId) > 0 Then txtInsurCompCode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtleaseno
    Set PO_DESC = txtLeaseName

    PR_AccountNo.Filter = "CustomerNo = '" & txtCustomerNo & "'"
    GoTop PR_AccountNo
    MyLookup.Caption = "Account Nos"
    MyLookup.FillGrid PR_AccountNo, "LeaseNo", "LeaseRental", txtleaseno.MaxLength
    MyLookup.Show 1
    If Len(txtleaseno) > 0 Then txtleaseno_KeyDown vbKeyReturn, vbKeyShift
    PR_AccountNo.Filter = adFilterNone
End Sub

Private Sub dtpDOExpiry_KeyDown(KeyCode As Integer, Shift As Integer)
    If Lastkey(KeyCode) Then txtAmount.SetFocus
End Sub

Private Sub dtpDOIssue_KeyDown(KeyCode As Integer, Shift As Integer)
    If Lastkey(KeyCode) Then dtpDOExpiry.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF11 Then Call DentMode(Mode, 4, PR_LMInsurance, Me, cmbInsurType, txtCustomerNo, "X", "CompCount", 3, "CustomerNo", "CustomerName", 1, False, Toolbar1)
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
' Setting up Preveliges
  SetToolBar(1) = chkRights("LMSRNTLPT1")
  SetToolBar(2) = chkRights("LMSRNTLPT2")
  SetToolBar(3) = chkRights("LMSRNTLPT3")
  SetToolBar(4) = chkRights("LMSRNTLPT4")

  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  PR_LMInsurance.Open "SELECT LM_Insurance.*, CompCode + BranchCode + CustomerNo + AccountNo + InsurCode + RTrim(InsurNo) AS FindFld FROM LM_Insurance WHERE CompCode + BranchCode = '" & Gs_compcode + Gs_BranchCode & "'", gc_dbcon, adOpenDynamic, adLockOptimistic
  PR_InsurComp.Open "SELECT * FROM LM_InsurComp", gc_dbcon, adOpenStatic, adLockReadOnly
  
  Screen.MousePointer = vbDefault
  
  IF_BlnkInflow = IIf(PR_LMInsurance.EOF, True, False)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If pr_Customer.State = 1 Then pr_Customer.Close
    If PR_AccountNo.State = 1 Then PR_AccountNo.Close
    PR_InsurComp.Close
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If Lastkey(KeyCode) Then txtRemarks.SetFocus
End Sub

Private Sub txtCustomerNo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn And txtCustomerNo.Text <> "" Then
      txtCustomerNo = DoPad(txtCustomerNo, txtCustomerNo.MaxLength)
        If Not MySeek(txtCustomerNo.Text, "CustomerNo", pr_Customer) Then
              Call SetErr(Gs_RecNFMsg, vbCritical)
              txtCustomerNo.SetFocus
        Else
              txtCustomerName.Text = pr_Customer("CustomerName")
              txtleaseno.SetFocus
        End If
   ElseIf KeyCode = vbKeyF12 Then
      Call cmdlookup2_Click
   ElseIf KeyCode = vbKeyPageUp Then
      cmbInsurType.SetFocus
   End If
End Sub

Private Sub txtInsurCode_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And Len(txtInsurCode.Text) > 0 Then
    txtInsurCode = DoPad(txtInsurCode, txtInsurCode.MaxLength)
    If Mode <> "A" Then
         If Not MySeek(Gs_compcode + Gs_BranchCode + txtCustomerNo + txtleaseno + txtInsurCompCode + txtInsurCode, "FindFld", PR_LMInsurance) Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
         Else
             SetVal
             If dtpDOIssue.Enabled = True Then dtpDOIssue.SetFocus
         End If
    Else
        If Lastkey(KeyCode) Then dtpDOIssue.SetFocus
    End If
 End If
End Sub

Private Sub txtleaseno_KeyDown(KeyCode As Integer, Shift As Integer)

If Lastkey(KeyCode) And txtleaseno.Text <> "" Then
   txtleaseno = DoPad(txtleaseno, txtleaseno.MaxLength)
   If Not MySeek(txtCustomerNo + txtleaseno, "FindFld", PR_AccountNo) Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtleaseno.SetFocus
   Else
      txtLeaseName = PR_AccountNo("LeaseRental") & ""
      txtInsurCompCode.SetFocus
   End If
ElseIf KeyCode = vbKeyF12 Then
      Call Command2_Click
ElseIf KeyCode = vbKeyPageUp Then
      txtCustomerNo.SetFocus
End If
End Sub
Private Sub txtInsurCompCode_KeyDown(KeyCode As Integer, Shift As Integer)
If Lastkey(KeyCode) And txtInsurCompCode.Text <> "" Then
   txtInsurCompCode = DoPad(txtInsurCompCode, txtInsurCompCode.MaxLength)
   If Not MySeek(txtInsurCompCode, "InsurCode", PR_InsurComp) Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtInsurCompCode.SetFocus
   Else
       txtInsurCompName = PR_InsurComp.Fields("InsurDesc")
       txtInsurCode.SetFocus
   End If
ElseIf KeyCode = vbKeyF12 Then
      Call Command1_Click
ElseIf KeyCode = vbKeyPageUp Then
      txtleaseno.SetFocus
End If
End Sub

'Private Sub txtTransCode_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim lb_found As Boolean
'
' If KeyCode = vbKeyReturn And Len(txtTransCode.Text) > 0 Then
'
'         txtTransCode.Text = DoPad(txtTransCode.Text, txtTransCode.MaxLength)
'         lb_found = MySeek(txtTransCode.Text, "Transcode", PR_LeasePmt)
'
'       Select Case Mode
'            Case "A"
'                If lb_found Then
'                   Call SetErr(Gs_RecFdMsg, vbCritical)
'                    SetClear Me
'                   txtTransCode.SetFocus
'                Else
'                   dtpValueDate.SetFocus
'                End If
'            Case Else
'                If Not lb_found Then
'                   Call SetErr(Gs_RecNFMsg, vbCritical)
'                    SetClear Me
'                   txtTransCode.SetFocus
'                Else
'                   Call SetVal
'                   If Mode <> "D" Then dtpValueDate.SetFocus
'                End If
'            End Select
'   ElseIf KeyCode = vbKeyF12 Then
'       ' Call cmdLookup4_Click
'   End If
'  End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If Button.Index = 1 Then
    ElseIf Range(Button.Index, 2, 3) Then
        PR_LMInsurance.Requery
    End If
    
    If IF_BlnkInflow And Range(Button.Index, 2, 3) Then
       Call SetErr("Data not found.", vbCritical)
       Mode = ""
       Cancel = True
    Else
        Mode = DentMode(Mode, Button.Index, PR_LMInsurance, Me, cmbInsurType, cmbInsurType, "X", "CompCount", 3, "CustomerNo", "CustomerName", 1, False, Toolbar1)
    End If
    
End Sub

Public Sub SaveValues()
Dim cntsql As New ADODB.Command
IF_BlnkInflow = False

cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText
gc_dbcon.BeginTrans
     Select Case Mode
           Case "A"
              cntsql.CommandText = "INSERT into LM_Insurance(CompCode, BranchCode, CustomerNo, AccountNo, InsurCode, InsurType, InsurNo, InsurAmount, DOIssue, DOExpiry, Remarks, UserId, TransDate, TransTime) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtCustomerNo & "','" & txtleaseno & "','" & txtInsurCompCode & "','" & InsuranceType & "', '" & txtInsurCode.Text & "', " & Val(txtAmount) & ", '" & Format(dtpDOIssue.Value, "YYYY/MM/DD") & "','" & Format(dtpDOExpiry.Value, "YYYY/MM/DD") & "','" & Trim(txtRemarks) & "','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS AMPM") & "')"
              cntsql.Execute
           Case "E"
              cntsql.CommandText = "UPDATE LM_Insurance SET InsurAmount = " & Val(txtAmount) & ", DOIssue = '" & Format(dtpDOIssue.Value, "YYYY/MM/DD") & "', DOExpiry = '" & Format(dtpDOExpiry.Value, "YYYY/MM/DD") & "', Remarks = '" & Trim(txtRemarks.Text) & "' WHERE InsurType+Compcode+BranchCode+CustomerNo+AccountNo+InsurCode+InsurNo = '" & InsuranceType + Gs_compcode + Gs_BranchCode + txtCustomerNo + txtleaseno + txtInsurCompCode + txtInsurCode & "'"
              cntsql.Execute
           Case "D"
              cntsql.CommandText = "DELETE FROM LM_Insurance WHERE InsurType+Compcode+BranchCode+CustomerNo+AccountNo+InsurCode+InsurNo = '" & InsuranceType + Gs_compcode + Gs_BranchCode + txtCustomerNo + txtleaseno + txtInsurCompCode + txtInsurCode & "'"
              cntsql.Execute
     End Select
gc_dbcon.CommitTrans

PR_LMInsurance.Requery
End Sub
Private Sub SetVal()
    
    InsuranceType = PR_LMInsurance("InsurType") & ""
    If InsuranceType = "L" Then
        cmbInsurType.ListIndex = 0
    Else
        cmbInsurType.ListIndex = 1
    End If
    cmbInsurType_Click
    txtCustomerNo = PR_LMInsurance("CustomerNo") & ""
    If MySeek(txtCustomerNo.Text, "CustomerNo", pr_Customer) Then
        txtCustomerName.Text = pr_Customer.Fields("CustomerName")
    End If
    txtleaseno = PR_LMInsurance("AccountNo") & ""
    If MySeek(txtCustomerNo.Text + txtleaseno.Text, "FindFLD", PR_AccountNo) Then
        txtLeaseName.Text = PR_AccountNo.Fields("LeaseRental") & ""
    End If
    txtInsurCompCode.Text = PR_LMInsurance("InsurCode")
    If MySeek(txtInsurCompCode.Text, "InsurCode", PR_InsurComp) Then
        txtInsurCompName.Text = PR_InsurComp.Fields("InsurDesc") & ""
    End If
    txtInsurCode.Text = PR_LMInsurance.Fields("InsurNo") & ""
    dtpDOIssue.Value = IIf(PR_LMInsurance.Fields("DOIssue").Value = Null, Date, PR_LMInsurance.Fields("DOIssue").Value)
    dtpDOExpiry.Value = IIf(PR_LMInsurance.Fields("DOExpiry").Value = Null, Date, PR_LMInsurance.Fields("DOExpiry").Value)
    txtAmount = PR_LMInsurance.Fields("InsurAmount") & ""
    txtRemarks = PR_LMInsurance.Fields("Remarks") & ""
End Sub
Public Function ChkInputs() As Boolean
    
    If cmbInsurType.Text <> "" And txtCustomerNo.Text <> "" And txtleaseno.Text <> "" And txtInsurCompCode.Text <> "" And txtAmount.Text <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If

End Function

Public Sub FrmRefresh()
  PR_LMInsurance.Requery
  PR_InsurComp.Requery
  pr_Customer.Requery
  PR_AccountNo.Requery
End Sub

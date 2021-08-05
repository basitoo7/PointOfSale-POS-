VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLeaseAttrib 
   Caption         =   "Lease Attrib"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4845
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLeaseAttrib.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      TabIndex        =   1
      Top             =   570
      Width           =   4845
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   4785
         MaxLength       =   50
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   555
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Frame Frame3 
         Height          =   2430
         Left            =   2310
         TabIndex        =   13
         Top             =   900
         Width           =   2490
         Begin VB.CheckBox ChkIslamic 
            Alignment       =   1  'Right Justify
            Caption         =   "Islamic Portfolio"
            Height          =   330
            Left            =   420
            TabIndex        =   24
            Top             =   2040
            Width           =   1650
         End
         Begin VB.CheckBox CHKFDCF 
            Alignment       =   1  'Right Justify
            Caption         =   "FDCF Reporting  :"
            Height          =   330
            Left            =   420
            TabIndex        =   21
            Top             =   1770
            Width           =   1650
         End
         Begin VB.CheckBox CHKSBP 
            Alignment       =   1  'Right Justify
            Caption         =   "SBP Reporting    : "
            Height          =   330
            Left            =   420
            TabIndex        =   20
            Top             =   1225
            Width           =   1650
         End
         Begin VB.CheckBox CHKIFC 
            Alignment       =   1  'Right Justify
            Caption         =   "IFC Reporting      :"
            Height          =   330
            Left            =   420
            TabIndex        =   19
            Top             =   1500
            Width           =   1650
         End
         Begin VB.CheckBox ChkCIB 
            Alignment       =   1  'Right Justify
            Caption         =   "CIB Reporting     :"
            Height          =   330
            Left            =   420
            TabIndex        =   18
            Top             =   695
            Width           =   1650
         End
         Begin VB.CheckBox ChkSECP 
            Alignment       =   1  'Right Justify
            Caption         =   "SECP Reporting  :"
            Height          =   330
            Left            =   420
            TabIndex        =   17
            Top             =   960
            Width           =   1650
         End
         Begin VB.CheckBox txtprostat 
            Alignment       =   1  'Right Justify
            Caption         =   "Provision Status :"
            Height          =   330
            Left            =   420
            TabIndex        =   15
            Top             =   430
            Width           =   1650
         End
         Begin VB.CheckBox txtbilling 
            Alignment       =   1  'Right Justify
            Caption         =   "Biling                   :"
            Height          =   330
            Left            =   420
            TabIndex        =   14
            Top             =   165
            Width           =   1650
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Lease Agreement Status"
         ForeColor       =   &H00000080&
         Height          =   2415
         Left            =   75
         TabIndex        =   9
         Top             =   915
         Width           =   2190
         Begin MSMask.MaskEdBox txtlitigamt 
            Height          =   375
            Left            =   825
            TabIndex        =   23
            Tag             =   "SKIP"
            Top             =   1905
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   8
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.OptionButton txtlitigation 
            Alignment       =   1  'Right Justify
            Caption         =   "Litigation :"
            Height          =   285
            Left            =   450
            TabIndex        =   12
            Top             =   1380
            Width           =   1080
         End
         Begin VB.OptionButton txtactive 
            Alignment       =   1  'Right Justify
            Caption         =   "Active :"
            Height          =   285
            Left            =   615
            TabIndex        =   11
            Top             =   855
            Width           =   915
         End
         Begin VB.OptionButton txtterminated 
            Alignment       =   1  'Right Justify
            Caption         =   "Terminated :"
            Height          =   285
            Left            =   300
            TabIndex        =   10
            Top             =   375
            Width           =   1230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Amount :"
            Height          =   210
            Left            =   120
            TabIndex        =   22
            Top             =   1965
            Width           =   645
         End
      End
      Begin VB.CommandButton cmdLookup 
         Appearance      =   0  'Flat
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
         Left            =   2130
         Picture         =   "frmLeaseAttrib.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "SKIP"
         Top             =   165
         Width           =   315
      End
      Begin VB.TextBox txtleaseno 
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   4
         Tag             =   "SKIP"
         Top             =   525
         Width           =   465
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
         Left            =   1800
         Picture         =   "frmLeaseAttrib.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "SKIP"
         Top             =   525
         Width           =   315
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2490
         MaxLength       =   50
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   165
         Width           =   2265
      End
      Begin MSMask.MaskEdBox txtCustNO 
         Height          =   315
         Left            =   1305
         TabIndex        =   6
         Tag             =   "SKIP"
         Top             =   165
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Customer Code :"
         Height          =   315
         Left            =   90
         TabIndex        =   8
         Top             =   210
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lease # :"
         Height          =   210
         Index           =   0
         Left            =   600
         TabIndex        =   7
         Top             =   555
         Width           =   675
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   1005
      ButtonWidth     =   1217
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
            Description     =   "Add"
            Object.ToolTipText     =   "Add new record"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Edit"
            Description     =   "Edit"
            Object.ToolTipText     =   "Edit an existing record"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Delete"
            Description     =   "Remove "
            Object.ToolTipText     =   "Remove an existing record."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Save"
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
            Caption         =   "Refresh"
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
         Left            =   4920
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
               Picture         =   "frmLeaseAttrib.frx":05EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLeaseAttrib.frx":0A42
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLeaseAttrib.frx":0E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLeaseAttrib.frx":12EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLeaseAttrib.frx":173E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLeaseAttrib.frx":1B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLeaseAttrib.frx":22E6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmLeaseAttrib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkGls0 As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PI_AStatus As Integer
Dim pr_Customer As New Recordset
Dim PR_LMSInfo As New Recordset

Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCustNo
    Set PO_DESC = Text4
    GoTop pr_Customer
    MyLookup.Caption = "Customer"
    MyLookup.FillGrid pr_Customer, "CustomerNo", "CustomerName", txtCustNo.MaxLength
    MyLookup.Show 1
    If Len(txtCustNo) > 0 Then TxtCustNo_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtleaseno
    Set PO_DESC = Text1
    PR_LMSInfo.Filter = "CustomerNo = '" & txtCustNo & "'"
    MyLookup.Caption = "Lease Agreements"
    MyLookup.FillGrid PR_LMSInfo, "LeaseNo", "LeaseAmount", txtleaseno.MaxLength
    MyLookup.Show 1
    PR_LMSInfo.Filter = adFilterNone
    If Len(txtleaseno) > 0 Then txtleaseno_KeyDown vbKeyReturn, vbKeyShift
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF11 Then Call DentMode(Mode, 4, PR_LMSInfo, Me, txtCustNo, txtCustNo, PR_LMSInfo, "BranchCnt", 3, "BranchCode", "BranchDesc", 1, False, Toolbar1)
End Sub

Private Sub Form_Load()
  
' Setting up Preveliges
  
  SetToolBar(1) = chkRights("LMATTRIB01")
  SetToolBar(2) = chkRights("LMATTRIB02")
  SetToolBar(3) = chkRights("LMATTRIB03")
  SetToolBar(4) = chkRights("LMATTRIB04")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  Set PR_SyProc = New Recordset
  Set PR_Company = New Recordset
  pr_Customer.Open "Select Customer.* from Customer Left Outer Join Facilities On Customer.CustomerNo = Facilities.CustomerNo Where Customer.Compcode+Customer.BranchCode = '" & Gs_compcode + Gs_BranchCode & "' And Facilities.FacilityNo = '01' Order By Customer.CustomerNo", gc_dbcon, adOpenStatic, adLockOptimistic, 1
  PR_LMSInfo.Open "Select *,CustomerNo+LeaseNo As FindFld from LM_LeaseInfo where compcode + BranchCode ='" & Gs_compcode + Gs_BranchCode & "' Order by BranchCode,CustomerNo,LeaseNo", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PB_BlnkGls0 = IIf(PR_LMSInfo.EOF, True, False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pr_Customer.Close
    PR_LMSInfo.Close
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If PB_BlnkGls0 And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_LMSInfo, Me, txtCustNo, txtCustNo, PR_LMSInfo, "BranchCnt", 3, "BranchCode", "BranchDesc", 1, False, Toolbar1)
    End If
    
End Sub

Public Sub SaveValues()
PB_BlnkGls0 = False
     If Mode = "E" Or Mode = "A" Then
     PR_LMSInfo("ActiveStatus") = PI_AStatus
     If PI_AStatus = 2 Then PR_LMSInfo("litigation_amt") = Val(0 & txtlitigamt)
     PR_LMSInfo("ReminderStat") = txtbilling
     PR_LMSInfo("ProvisionStat") = txtprostat
     PR_LMSInfo("CIBStatus") = ChkCIB.Value
     PR_LMSInfo("SECPStatus") = ChkSECP.Value
     PR_LMSInfo("SBPStatus") = CHKSBP.Value
     PR_LMSInfo("IFCStatus") = CHKIFC.Value
     PR_LMSInfo("FDCFStatus") = CHKFDCF.Value
     PR_LMSInfo("Islamic") = ChkIslamic.Value
     PR_LMSInfo.Update
     PR_LMSInfo.Requery
     End If
End Sub
Private Sub SetVal()
txtterminated.Value = IIf(PR_LMSInfo("ActiveStatus") = 0, True, False)
txtactive.Value = IIf(PR_LMSInfo("ActiveStatus") = 1, True, False)
txtlitigation.Value = IIf(PR_LMSInfo("ActiveStatus") = 2, True, False)
If PR_LMSInfo("ActiveStatus") = 2 Then txtlitigamt = Val(0 & PR_LMSInfo("litigation_amt"))
txtprostat.Value = Val(0 & PR_LMSInfo("ProvisionStat"))
txtbilling.Value = Val(0 & PR_LMSInfo("ReminderStat"))
ChkCIB.Value = Val(0 & PR_LMSInfo("CIBStatus"))
ChkSECP.Value = Val(0 & PR_LMSInfo("SECPStatus"))
CHKSBP.Value = Val(0 & PR_LMSInfo("SBPStatus"))
CHKIFC.Value = Val(0 & PR_LMSInfo("IFCStatus"))
CHKFDCF.Value = Val(0 & PR_LMSInfo("FDCFStatus"))
ChkIslamic.Value = Val(0 & PR_LMSInfo("Islamic"))
End Sub
Public Function ChkInputs() As Boolean
    If txtCustNo <> "" And txtleaseno <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function
Private Sub txtactive_Click()
PI_AStatus = 1
txtbilling.Enabled = True
txtlitigamt.Enabled = False
txtlitigamt = ""
txtprostat.Enabled = True
ChkCIB.Enabled = True
ChkSECP.Enabled = True
CHKSBP.Enabled = True
CHKIFC.Enabled = True
CHKFDCF.Enabled = True
ChkIslamic.Enabled = True
End Sub
Private Sub TxtCustNo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If Lastkey(KeyCode) And Len(txtCustNo.Text) > 0 Then
        txtCustNo.Text = IIf(IsNumeric(txtCustNo), DoPad(txtCustNo, txtCustNo.MaxLength), UCase(txtCustNo))
        lb_found = MySeek(txtCustNo.Text, "CustomerNo", pr_Customer)
       
       Select Case Mode
            Case "A"
                If Not lb_found Then
                   Call SetErr("Record Not Found", vbCritical)
                   txtCustNo.SetFocus
                Else
                   Text4 = pr_Customer("CustomerName")
                   txtleaseno.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   txtCustNo.SetFocus
                Else
                   Text4 = pr_Customer("CustomerName")
                   txtleaseno.Enabled = True
                   txtleaseno.SetFocus
                End If
            End Select
       ElseIf KeyCode = vbKeyF12 Then
             cmdLookup_Click
       End If
  End Sub

Private Sub txtleaseno_KeyDown(KeyCode As Integer, Shift As Integer)

If Lastkey(KeyCode) And txtleaseno.Text <> "" Then
   txtleaseno = DoPad(txtleaseno, txtleaseno.MaxLength)
   lb_found = MySeek(Trim(txtCustNo) + Trim(txtleaseno), "FindFld", PR_LMSInfo)
   
   If lb_found Then
    If Mode = "E" Or Mode = "D" Or Mode = "" Then
       Call SetVal
    Else
    txtterminated.SetFocus
    End If
  Else
      Call SetErr("Lease not found for this Customer", vbCritical)
      txtleaseno.SetFocus
   End If
ElseIf KeyCode = vbKeyF12 Then
     Command1_Click
End If
End Sub

Private Sub txtlitigation_Click()
PI_AStatus = 2
txtlitigamt.Enabled = True
txtlitigamt.SetFocus
txtbilling.Value = 0
txtbilling.Enabled = False
txtprostat.Value = 0
txtprostat.Enabled = False
ChkCIB.Value = 0
ChkCIB.Enabled = False
ChkSECP.Value = 0
ChkSECP.Enabled = False
CHKSBP.Value = 0
CHKSBP.Enabled = False
CHKIFC.Value = 0
CHKIFC.Enabled = False
CHKFDCF.Value = 0
CHKFDCF.Enabled = False
ChkIslamic.Value = 0
ChkIslamic.Enabled = False

End Sub
Private Sub txtterminated_Click()
PI_AStatus = 0
txtbilling.Value = 0
txtlitigamt.Enabled = False
txtlitigamt = ""
txtbilling.Enabled = False
txtprostat.Value = 0
txtprostat.Enabled = False
ChkCIB.Value = 0
ChkCIB.Enabled = False
ChkSECP.Value = 0
ChkSECP.Enabled = False
CHKSBP.Value = 0
CHKSBP.Enabled = False
CHKIFC.Value = 0
CHKIFC.Enabled = False
CHKFDCF.Value = 0
CHKFDCF.Enabled = False
ChkIslamic.Value = 0
ChkIslamic.Enabled = False
End Sub

Public Sub FrmRefresh()
  pr_Customer.Requery
  PR_LMSInfo.Requery
End Sub

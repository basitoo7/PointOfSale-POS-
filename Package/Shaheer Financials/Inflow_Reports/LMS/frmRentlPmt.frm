VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRentlPmt 
   Caption         =   "Adjustment against Lease Payment"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRentlPmt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   0
      TabIndex        =   1
      Top             =   570
      Width           =   5460
      Begin VB.TextBox txtInstrmentNo 
         Height          =   315
         Left            =   1155
         MaxLength       =   50
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1650
         Width           =   4245
      End
      Begin VB.TextBox txtleaseno 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1170
         MaxLength       =   3
         TabIndex        =   7
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
         Picture         =   "frmRentlPmt.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmRentlPmt.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   570
         Width           =   315
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
         Picture         =   "frmRentlPmt.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1290
         Width           =   315
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2340
         MaxLength       =   50
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   570
         Width           =   3075
      End
      Begin MSMask.MaskEdBox text2 
         Height          =   315
         Left            =   2340
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1290
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483644
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   50
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
      Begin MSComCtl2.DTPicker dtpValueDate 
         Height          =   285
         Left            =   4155
         TabIndex        =   3
         Top             =   210
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   22806529
         CurrentDate     =   37293
      End
      Begin MSMask.MaskEdBox txtTransCode 
         Height          =   315
         Left            =   1170
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   210
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
      Begin MSMask.MaskEdBox txtCustomerNo 
         Height          =   315
         Left            =   1170
         TabIndex        =   4
         TabStop         =   0   'False
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
      Begin MSMask.MaskEdBox txtPmtId 
         Height          =   315
         Left            =   1170
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1290
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   3
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
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2010
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
      Begin MSMask.MaskEdBox textx 
         Height          =   315
         Left            =   4920
         TabIndex        =   22
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   930
         Visible         =   0   'False
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483644
         PromptInclude   =   0   'False
         MaxLength       =   50
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
      Begin MSMask.MaskEdBox text3 
         Height          =   315
         Left            =   2340
         TabIndex        =   23
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   930
         Width           =   1065
         _ExtentX        =   1879
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
      Begin MSMask.MaskEdBox text5 
         Height          =   315
         Left            =   3480
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   930
         Visible         =   0   'False
         Width           =   1065
         _ExtentX        =   1879
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
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Amount :"
         Height          =   210
         Left            =   465
         TabIndex        =   21
         Top             =   2010
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Inst'ment No :"
         Height          =   210
         Left            =   150
         TabIndex        =   20
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lease # :"
         Height          =   210
         Index           =   0
         Left            =   450
         TabIndex        =   17
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Customer # :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   225
         TabIndex        =   16
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Reference Id :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   105
         TabIndex        =   15
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "On A/c Of :"
         Height          =   210
         Left            =   300
         TabIndex        =   14
         Top             =   1320
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Value Date :"
         Height          =   210
         Left            =   3225
         TabIndex        =   11
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Value Date :"
         Height          =   210
         Index           =   1
         Left            =   3075
         TabIndex        =   10
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
      Width           =   5505
      _ExtentX        =   9710
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
               Picture         =   "frmRentlPmt.frx":0760
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentlPmt.frx":0BB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentlPmt.frx":1008
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentlPmt.frx":145C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentlPmt.frx":18B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentlPmt.frx":1D04
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentlPmt.frx":2458
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmRentlPmt"
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

Dim PR_LMShdl As New Recordset
Dim PR_LMSPmtId As New Recordset
Dim PR_LMSInfo As New Recordset
Dim PR_LeasePmt As New Recordset
Dim PR_Customer As New Recordset
Private Sub cmdlookup2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCustomerNo
    Set PO_DESC = Text1
    
    GoTop PR_Customer
    MyLookup.Caption = "Customers"
    MyLookup.FillGrid PR_Customer, "CustomerNo", "CustomerName", txtCustomerNo.MaxLength
    MyLookup.Show 1
    
    If Len(txtCustomerNo) > 0 Then txtCustomerNo_KeyDown vbKeyReturn, vbKeyShift
End Sub
Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtPmtId
    Set PO_DESC = Textx

    GoTop PR_LMSPmtId
    MyLookup.Caption = "Payment As"
    MyLookup.FillGrid PR_LMSPmtId, "IdCode", "IdDescrip", txtPmtId.MaxLength + 2
    MyLookup.Show 1
    If Len(txtPmtId) > 0 Then txtPmtId_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtleaseno
    Set PO_DESC = Text1

    PR_LMSInfo.Filter = "BranchCode = '" & Gs_BranchCode & "' And CustomerNo = '" & txtCustomerNo & "'"
    GoTop PR_LMSInfo
    MyLookup.Caption = "Lease Agreements"
    MyLookup.FillGrid PR_LMSInfo, "LeaseNo", "LeaseRental", txtleaseno.MaxLength
    MyLookup.Show 1
    If Len(txtleaseno) > 0 Then txtleaseno_KeyDown vbKeyReturn, vbKeyShift
    PR_LMSInfo.Filter = adFilterNone
End Sub
Private Sub dtpValueDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If LastKey(KeyCode) Then txtCustomerNo.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF11 Then Call DentMode(Mode, 4, PR_LeasePmt, Me, txtTransCode, dtpValueDate, ParaCntr_Rs, "LMPmtTransCode", 10, "Transcode", "CustomerNo", 0, False, Toolbar1)
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
  
  PR_LMSPmtId.Open "Select * From FCM_IDs where RecId = 'LMR' Order By IdCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
  PR_LMSInfo.Open "Select *,BranchCode+CustomerNo+LeaseNo As FindFld from LM_LeaseInfo where compcode ='" & Gs_compcode & "' And ActiveStatus >=1 Order by BranchCode,CustomerNo,LeaseNo", gc_dbcon, adOpenStatic, adLockOptimistic, 1
  PR_Customer.Open "Select Customer.*,Customer.BranchCode+Customer.CustomerNo As FindFld from Customer Inner Join Facilities On Customer.CustomerNo = Facilities.CustomerNo Where Customer.Compcode+Customer.BranchCode = '" & Gs_compcode + Gs_BranchCode & "' And Facilities.FacilityNo = '01' Order By Customer.CustomerNo,Customer.BranchCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
  PR_LeasePmt.Open "Select LM_Payments.* from LM_Payments Where Compcode = '" & Gs_compcode & "' Order By BranchCode,Transcode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_LMShdl.Open "Select *,BranchCode+CustomerNo+LeaseNo As FindFld from LM_Schedule Where Compcode = '" & Gs_compcode & "' Order By BranchCode,CustomerNo,LeaseNo,AccrualDate", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  
  Screen.MousePointer = vbDefault
  
  IF_BlnkInflow = IIf(PR_LeasePmt.EOF, True, False)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_LMSInfo.Close
    PR_Customer.Close
    PR_LeasePmt.Close
    PR_LMSPmtId.Close
    PR_LMShdl.Close
End Sub

Private Sub txtCustomerNo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn And txtCustomerNo.Text <> "" Then
      txtCustomerNo = DoPad(txtCustomerNo, txtCustomerNo.MaxLength)
        If Not MySeek(txtCustomerNo.Text, "CustomerNo", PR_Customer) Then
              Call SetErr(Gs_RecNFMsg, vbCritical)
              txtCustomerNo.SetFocus
        Else
              Text1.Text = PR_Customer("CustomerName")
              txtleaseno.SetFocus
        End If
   ElseIf KeyCode = vbKeyF12 Then
      Call cmdlookup2_Click
   ElseIf KeyCode = vbKeyPageUp Then
      dtpValueDate.SetFocus
   End If
End Sub
Private Sub txtInstrmentNo_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
    txtAmount.SetFocus
 ElseIf KeyCode = vbKeyPageUp Then
    txtPmtId.SetFocus
 End If
End Sub
Private Sub txtleaseno_KeyDown(KeyCode As Integer, Shift As Integer)

If LastKey(KeyCode) And txtleaseno.Text <> "" Then
   txtleaseno = DoPad(txtleaseno, txtleaseno.MaxLength)
   If Not MySeek(Gs_BranchCode + txtCustomerNo + txtleaseno, "FindFld", PR_LMSInfo) Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtleaseno.SetFocus
   Else
      Text3 = PR_LMSInfo("LeaseRental")
      txtPmtId.SetFocus
   End If
ElseIf KeyCode = vbKeyF12 Then
      Call Command2_Click
ElseIf KeyCode = vbKeyPageUp Then
      txtCustomerNo.SetFocus
End If
End Sub
Private Sub txtPmtId_KeyDown(KeyCode As Integer, Shift As Integer)
If LastKey(KeyCode) And txtPmtId.Text <> "" Then
   txtPmtId = UCase(txtPmtId)
   If Not MySeek(txtPmtId, "IdCode", PR_LMSPmtId) Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtPmtId.SetFocus
   Else
       Text2 = PR_LMSPmtId.Fields("IdDescrip")
       txtInstrmentNo.SetFocus
   End If
ElseIf KeyCode = vbKeyF12 Then
      Call Command1_Click
ElseIf KeyCode = vbKeyPageUp Then
      txtleaseno.SetFocus
End If
End Sub
Private Sub txtTransCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn And Len(txtTransCode.Text) > 0 Then
          
         txtTransCode.Text = DoPad(txtTransCode.Text, txtTransCode.MaxLength)
         lb_found = MySeek(txtTransCode.Text, "Transcode", PR_LeasePmt)
         
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                    SetClear Me
                   txtTransCode.SetFocus
                Else
                   dtpValueDate.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                    SetClear Me
                   txtTransCode.SetFocus
                Else
                   Call SetVal
                   If Mode <> "D" Then dtpValueDate.SetFocus
                End If
            End Select
   ElseIf KeyCode = vbKeyF12 Then
       ' Call cmdLookup4_Click
   End If
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If Button.Index = 1 Then
'       cmdLookup4.Enabled = False
    ElseIf Range(Button.Index, 2, 3) Then
        PR_LeasePmt.Requery
     '  cmdLookup4.Enabled = True
    End If
    
    If IF_BlnkInflow And Range(Button.Index, 2, 3) Then
       Call SetErr("Data not found.", vbCritical)
       Mode = ""
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_LeasePmt, Me, txtTransCode, dtpValueDate, ParaCntr_Rs, "LMPmtTransCode", 10, "Transcode", "CustomerNo", 0, False, Toolbar1)
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
              cntsql.CommandText = "INSERT into LM_Payments(Compcode,BranchCode,Transcode,CustomerNo,LeaseNo,ChqDate,RelzDate,PaymentId,InstrNo,PaidAmount,UserId,TransDate,TransTime) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtTransCode & "','" & txtCustomerNo & "','" & txtleaseno & "','" & Format(dtpValueDate.Value, "YYYY/MM/DD") & "','" & Format(dtpValueDate.Value, "YYYY/MM/DD") & "','" & txtPmtId & "','" & txtInstrmentNo & "'," & Val(0 & txtAmount) & ",'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "')"
              cntsql.Execute
           Case "E"
              cntsql.CommandText = "UPDATE LM_Payments SET CustomerNo = '" & txtCustomerNo & "',LeaseNo = '" & txtleaseno & "',PaymentID = '" & txtPmtId & "',RelzDate = '" & Format(dtpValueDate.Value, "YYYY/MM/DD") & "',ChqDate = '" & Format(dtpValueDate.Value, "YYYY/MM/DD") & "',InstrNo = " & Val(0 & txtInstrmentNo) & ",PaidAmount = " & Val(0 & txtAmount) & " WHERE Compcode+BranchCode+Transcode = '" & Gs_compcode + Gs_BranchCode + txtTransCode & "'"
              cntsql.Execute
           Case "D"
           Dim IntoLoop As Boolean
           Dim ln_lessamt As Double
           Dim ln_lessamt2 As Double
           Dim ln_PrftAmt As Double
           Dim ln_CostAmt As Double
           Dim ln_Rental As Double
           
           IntoLoop = False
              cntsql.CommandText = "DELETE FROM LM_Payments WHERE Compcode+BranchCode+Transcode = '" & Gs_compcode + Gs_BranchCode + txtTransCode & "'"
              cntsql.Execute
            
            If PR_LeasePmt("RecPosted") = 1 Then
              If MySeek(Gs_BranchCode + txtCustomerNo + txtleaseno, "FindFld", PR_LMShdl) Then
                  Do While PR_LMShdl("FindFld") = Gs_BranchCode + txtCustomerNo + txtleaseno
                     If PR_LMShdl("RentalPaidAmt") <= 0 Then
SetAgain:
                        PR_LMShdl.MovePrevious
                        If PR_LMShdl("RentalPaidAmt") > 0 Then
                           ln_lessamt = Val(0 & txtAmount)
                           ln_lessamt2 = Val(0 & txtAmount)
                           Do While PR_LMShdl("FindFld") = Gs_BranchCode + txtCustomerNo + txtleaseno
                              IntoLoop = True
                              ln_PrftAmt = PR_LMShdl("ProfitPaidAmt")
                              ln_CostAmt = PR_LMShdl("CostPaidAmt")
                              ln_Rental = PR_LMShdl("RentalPaidAmt")
                              
                              PR_LMShdl("ProfitPaidAmt") = PR_LMShdl("ProfitPaidAmt") - IIf(ln_lessamt2 >= PR_LMShdl("ProfitPaidAmt"), PR_LMShdl("ProfitPaidAmt"), IIf(ln_lessamt2 > 0, ln_lessamt, 0))
                              ln_lessamt2 = ln_lessamt2 - IIf(ln_lessamt2 >= ln_PrftAmt, ln_PrftAmt, IIf(ln_lessamt2 > 0, ln_lessamt2, 0))
                              PR_LMShdl("CostPaidAmt") = PR_LMShdl("CostPaidAmt") - IIf(ln_lessamt2 >= PR_LMShdl("CostPaidAmt"), PR_LMShdl("CostPaidAmt"), IIf(ln_lessamt2 > 0, ln_lessamt2, 0))
                              ln_lessamt2 = ln_lessamt2 - IIf(ln_lessamt2 >= ln_CostAmt, ln_CostAmt, IIf(ln_lessamt2 > 0, ln_lessamt2, 0))
                              
                              PR_LMShdl("RentalPaidAmt") = PR_LMShdl("RentalPaidAmt") - IIf(ln_lessamt >= PR_LMShdl("RentalPaidAmt"), PR_LMShdl("RentalPaidAmt"), ln_lessamt)
                              PR_LMShdl("PaidDate") = IIf(PR_LMShdl("RentalPaidAmt") <= 0, Null, PR_LMShdl("PaidDate"))
                              ln_lessamt = ln_lessamt - IIf(ln_lessamt >= ln_Rental, ln_Rental, ln_lessamt)
                              PR_LMShdl.Update
                              
                              If ln_lessamt > 0 Then PR_LMShdl.MovePrevious
                              If ln_lessamt <= 0 Then Exit Do
                           Loop
                           Exit Do
                        End If
                     End If
                     PR_LMShdl.MoveNext
                  Loop
                  If Not IntoLoop Then GoTo SetAgain
              End If
            End If
     End Select
gc_dbcon.CommitTrans

PR_LeasePmt.Requery
End Sub
Private Sub SetVal()
     txtleaseno = PR_LeasePmt("LeaseNo") & ""
     txtCustomerNo = PR_LeasePmt("CustomerNo") & ""
     txtInstrmentNo = Val(0 & PR_LeasePmt("InstrNo"))
     txtAmount = Val(0 & PR_LeasePmt("PaidAmount"))
     txtPmtId = UCase(PR_LeasePmt("PaymentId")) & ""
     dtpValueDate = PR_LeasePmt("RelzDate")
     
     If MySeek(txtCustomerNo.Text, "CustomerNo", PR_Customer) Then
        Text1 = PR_Customer.Fields("CustomerName")
     End If
     
     If MySeek(txtPmtId.Text, "IDCode", PR_LMSPmtId) Then
        Text2 = PR_LMSPmtId.Fields("IDDescrip")
     End If
     
End Sub
Public Function ChkInputs() As Boolean
    
    If Len(txtTransCode.Text) = txtTransCode.MaxLength And txtCustomerNo.Text <> "" And txtleaseno.Text <> "" And txtAmount.Text <> "" And txtPmtId <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If

End Function


Public Sub FrmRefresh()
  PR_LMSPmtId.Requery
  PR_LMSInfo.Requery
  PR_Customer.Requery
  PR_LeasePmt.Requery
  PR_LMShdl.Requery
End Sub

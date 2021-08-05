VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmHRMBank1 
   Caption         =   "Bank Setup"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHRMBank1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   4950
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
      Height          =   3165
      Left            =   15
      TabIndex        =   1
      Top             =   570
      Width           =   4935
      Begin VB.TextBox txtVchrTypedesc2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2370
         MaxLength       =   64
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   2760
         Width           =   2475
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
         Left            =   2040
         Picture         =   "frmHRMBank1.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2760
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
         Left            =   2040
         Picture         =   "frmHRMBank1.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2400
         Width           =   315
      End
      Begin VB.TextBox txtvchrtypedesc1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2370
         MaxLength       =   64
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   2400
         Width           =   2475
      End
      Begin VB.TextBox txtvchrtypedesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2355
         MaxLength       =   64
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   2025
         Width           =   2475
      End
      Begin VB.CommandButton Command4 
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
         Left            =   2025
         Picture         =   "frmHRMBank1.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2025
         Width           =   315
      End
      Begin VB.CommandButton Command5 
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
         Left            =   2835
         Picture         =   "frmHRMBank1.frx":0760
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1305
         Width           =   315
      End
      Begin VB.TextBox txtglcode 
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
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   11
         ToolTipText     =   "Account No"
         Top             =   1320
         Width           =   1500
      End
      Begin VB.TextBox txtglactdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         MaxLength       =   64
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1680
         Width           =   3510
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   4620
         MaxLength       =   35
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   195
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.TextBox TxtManager 
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   600
         Width           =   3495
      End
      Begin VB.CommandButton cmdLookup 
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
         Left            =   2070
         Picture         =   "frmHRMBank1.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   240
         Width           =   315
      End
      Begin MSMask.MaskEdBox txtLocation 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         MaxLength       =   4
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
      Begin MSMask.MaskEdBox txtVchrType 
         Height          =   315
         Left            =   1320
         TabIndex        =   16
         Top             =   2025
         Width           =   675
         _ExtentX        =   1191
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtVchrType1 
         Height          =   315
         Left            =   1335
         TabIndex        =   21
         Top             =   2400
         Width           =   675
         _ExtentX        =   1191
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtVchrType2 
         Height          =   315
         Left            =   1335
         TabIndex        =   25
         Top             =   2760
         Width           =   675
         _ExtentX        =   1191
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
         PromptChar      =   " "
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Vchr Type (Adj) :"
         Height          =   255
         Left            =   30
         TabIndex        =   26
         Top             =   2775
         Width           =   1275
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Vchr Type (Sale) :"
         Height          =   255
         Left            =   30
         TabIndex        =   22
         Top             =   2415
         Width           =   1275
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Vchr Type (GRN) :"
         Height          =   255
         Left            =   45
         TabIndex        =   17
         Top             =   2055
         Width           =   1275
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "GL Description # :"
         Height          =   255
         Left            =   105
         TabIndex        =   14
         Top             =   1725
         Width           =   1155
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "GL Account # :"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1335
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Code :"
         Height          =   210
         Left            =   180
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         Height          =   210
         Left            =   375
         TabIndex        =   8
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Manager Name :"
         Height          =   210
         Left            =   105
         TabIndex        =   7
         Top             =   990
         Width           =   1170
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   1005
      ButtonWidth     =   1217
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
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
               Picture         =   "frmHRMBank1.frx":0A44
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHRMBank1.frx":0E98
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHRMBank1.frx":12EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHRMBank1.frx":1740
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHRMBank1.frx":1B94
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHRMBank1.frx":1FE8
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHRMBank1.frx":273C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmHRMBank1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkLoca As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_VchType As New Recordset

Dim Pr_Gldetail As New Recordset
Dim PR_Bank As New Recordset


Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtLocation
    Set PO_DESC = txtDesc
    
    GoTop PR_Bank
    MyLookup.Caption = "Banks"
    MyLookup.FillGrid PR_Bank, "BankCode", "BankName", txtLocation.MaxLength
    MyLookup.Show 1
    
    If Len(txtLocation) > 0 Then TxtLocation_KeyDown vbKeyReturn, vbKeyShift

End Sub


Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtVchrType1
    Set PO_DESC = txtvchrtypedesc1
    
    GoTop PR_VchType
    PR_VchType.Filter = "Branchcode = '" & Gs_BranchCode & "'"
    MyLookup.Caption = "Voucher Types"
    MyLookup.FillGrid PR_VchType, "VchrType", "VchrDescrip", 5
    MyLookup.Show 1
    PR_VchType.Filter = adFilterNone
    If Len(txtVchrType1) > 0 Then txtVchrType1_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtVchrType2
    Set PO_DESC = txtVchrTypedesc2
    
    GoTop PR_VchType
    PR_VchType.Filter = "Branchcode = '" & Gs_BranchCode & "'"
    MyLookup.Caption = "Voucher Types"
    MyLookup.FillGrid PR_VchType, "VchrType", "VchrDescrip", 5
    MyLookup.Show 1
    PR_VchType.Filter = adFilterNone
    If Len(txtVchrType2) > 0 Then txtVchrType2_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Command4_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtVchrType
    Set PO_DESC = Text1
    
    GoTop PR_VchType
    PR_VchType.Filter = "Branchcode = '" & Gs_BranchCode & "'"
    MyLookup.Caption = "Voucher Types"
    MyLookup.FillGrid PR_VchType, "VchrType", "VchrDescrip", 5
    MyLookup.Show 1
    PR_VchType.Filter = adFilterNone
    If Len(txtVchrType) > 0 Then txtVchrType_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Command5_Click()

    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtglcode
    Set PO_DESC = txtglactdesc
    Gs_SQL = "Select Accountno 'Account No', Acct_Desc  'Description' from gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_Subon = True
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
    Gs_OrderBy = "Order by Acct_Desc,AccountNo"
    MyLookupOLDB.Caption = "Account Nos."
    MyLookupOLDB.Show 1
    If Len(txtglcode) > 0 Then txtglcode_KeyDown vbKeyReturn, vbKeyShift
    

End Sub

Private Sub Form_Load()
  
  SetToolBar(1) = chkRights("FCMBANK001")
  SetToolBar(2) = chkRights("FCMBANK002")
  SetToolBar(3) = chkRights("FCMBANK003")
  SetToolBar(4) = chkRights("FCMBANK004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  

  PR_Bank.Open "Select * from SysBanks Order By BankCode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_VchType.Open "Select *,branchcode+vchrtype as findfld from GlVchrType where glVchrType.CompCode ='" & Gs_compcode & "' and VchrType <> '0OB' order by GlVchrType.VchrType ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PB_BlnkLoca = IIf(PR_Bank.EOF, True, False)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_Bank.Close
    PR_VchType.Close
End Sub


Private Sub txtadjvchrtype_Change()

End Sub

Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) Then
     TxtManager.SetFocus
  End If
End Sub

Private Sub txtglcode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtglcode <> "" Then
Pr_Gldetail.Open "select * from gl_detail where accountno = '" & txtglcode & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not Pr_Gldetail.EOF Then
        txtglactdesc = Pr_Gldetail("Acct_Desc")
Else
    Call MsgBox("Gl Account not found", vbCritical)
     txtglcode.SetFocus
End If
Pr_Gldetail.Close
End If
End Sub

Private Sub TxtLocation_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If Lastkey(KeyCode) Then
         
      txtLocation.Text = IIf(IsNumeric(txtLocation.Text), DoPad(UCase(txtLocation.Text), txtLocation.MaxLength), UCase(txtLocation.Text))
      lb_found = MySeek(txtLocation.Text, "BankCode", PR_Bank)
         
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   Cancel = True
                    SetClear Me
                   txtLocation.SetFocus
                Else
                   txtDesc.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   Cancel = True
                    SetClear Me
                   txtLocation.SetFocus
                Else
                   Call SetVal
                   If Mode <> "D" Then
                    '  txtLocation.Enabled = False
                      txtDesc.SetFocus
                   End If
                End If
            End Select
ElseIf KeyCode = vbKeyF12 Then
        Call cmdLookup_Click
End If
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
      cmdLookup.Enabled = False
    Else
      cmdLookup.Enabled = True
    End If
    
    If PB_BlnkLoca And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_Bank, Me, txtLocation, txtDesc, "X", "CompCount", 3, "BankCode", "BankName", 1, False, Toolbar1)
    End If
    
End Sub

Public Sub SaveValues()
Dim cntsql As New ADODB.Command
PB_BlnkLoca = False
Dim ls_btype As String

cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText
     Select Case Mode
           Case "A"
              cntsql.CommandText = "INSERT into SysBanks(BankCode,BankName,ManagerName,GlAccount,vchrtype,vchrtype1,vchrtype2) VALUES ('" & txtLocation.Text & "','" & txtDesc.Text & "','" & TxtManager.Text & "','" & txtglcode & "','" & txtVchrType & "','" & txtVchrType1 & "','" & txtVchrType2 & "')"
              cntsql.Execute
           Case "E"
              cntsql.CommandText = "UPDATE SysBanks SET BankName= '" & txtDesc.Text & "',ManagerName = '" & TxtManager & "',GlAccount = '" & txtglcode & "',vchrtype = '" & txtVchrType & "',vchrtype1 = '" & txtVchrType1 & "',vchrtype2 = '" & txtVchrType2 & "' WHERE  BankCode= '" & txtLocation.Text & "'"
              cntsql.Execute
           Case "D"
              cntsql.CommandText = "DELETE FROM SysBanks WHERE BankCode = '" & txtLocation.Text & "'"
              cntsql.Execute
     End Select
PR_Bank.Requery
End Sub

Private Sub SetVal()
     txtDesc = PR_Bank("BankName") & ""
     TxtManager = PR_Bank("ManagerName") & ""
     txtglcode = PR_Bank("GlAccount") & ""
     If txtglcode <> "" Then Call txtglcode_KeyDown(vbKeyReturn, vbKeyShift)
     txtVchrType = PR_Bank("vchrtype") & ""
     If txtVchrType <> "" Then Call txtVchrType_KeyDown(vbKeyReturn, vbKeyShift)
     txtVchrType1 = PR_Bank("vchrtype1") & ""
     If txtVchrType1 <> "" Then Call txtVchrType1_KeyDown(vbKeyReturn, vbKeyShift)
     txtVchrType2 = PR_Bank("vchrtype2") & ""
     If txtVchrType2 <> "" Then Call txtVchrType2_KeyDown(vbKeyReturn, vbKeyShift)
     
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtLocation.Text) = txtLocation.MaxLength And txtDesc <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub FrmRefresh()
  Pr_Gldetail.Requery
  PR_Bank.Requery
  PR_VchType.Requery
End Sub

Private Sub txtVchrType_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn And txtVchrType <> "" Then
    txtVchrType = UCase(txtVchrType)
    lb_found = MySeek(Trim(Gs_BranchCode) + Trim(txtVchrType.Text), "FindFld", PR_VchType)
    If Not lb_found Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtVchrType.SetFocus
    Else
       txtvchrtypedesc = PR_VchType("VchrDescrip")
    End If
  End If
End Sub

Private Sub txtVchrType1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn And txtVchrType1 <> "" Then
    txtVchrType1 = UCase(txtVchrType1)
    lb_found = MySeek(Trim(Gs_BranchCode) + Trim(txtVchrType1.Text), "FindFld", PR_VchType)
    If Not lb_found Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtVchrType1.SetFocus
    Else
       txtvchrtypedesc1 = PR_VchType("VchrDescrip")
    End If
  End If
End Sub

Private Sub txtVchrType2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn And txtVchrType2 <> "" Then
    txtVchrType2 = UCase(txtVchrType2)
    lb_found = MySeek(Trim(Gs_BranchCode) + Trim(txtVchrType2.Text), "FindFld", PR_VchType)
    If Not lb_found Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtVchrType2.SetFocus
    Else
       txtVchrTypedesc2 = PR_VchType("VchrDescrip")
    End If
  End If
End Sub


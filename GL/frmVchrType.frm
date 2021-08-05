VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVchrType 
   Caption         =   "Voucher Types"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVchrType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4875
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
      Height          =   2265
      Left            =   0
      TabIndex        =   0
      Top             =   570
      Width           =   4875
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   4425
         Picture         =   "frmVchrType.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   180
         Width           =   315
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4440
         MaxLength       =   50
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1890
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   4290
         Picture         =   "frmVchrType.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1410
         Width           =   315
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1485
         TabIndex        =   13
         Top             =   1020
         Width           =   2895
         Begin VB.OptionButton optJV 
            Caption         =   "&JV"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   4
            Top             =   30
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton optCash 
            Caption         =   "&Cash"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1980
            TabIndex        =   6
            Top             =   30
            Width           =   855
         End
         Begin VB.OptionButton optBank 
            Caption         =   "&Bank"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   960
            TabIndex        =   5
            Top             =   30
            Width           =   795
         End
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
         Left            =   2040
         Picture         =   "frmVchrType.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   180
         Width           =   315
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1500
         TabIndex        =   11
         Top             =   1830
         Width           =   1785
         Begin VB.OptionButton optYear 
            Caption         =   "Yearly"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   960
            TabIndex        =   8
            Top             =   120
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton optMonth 
            Caption         =   "Monthly"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   7
            Top             =   120
            Width           =   855
         End
      End
      Begin MSMask.MaskEdBox txtVchrDesc 
         Height          =   315
         Left            =   1500
         TabIndex        =   3
         Top             =   600
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "c"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtVchrType 
         Height          =   315
         Left            =   1500
         TabIndex        =   2
         Tag             =   "SKIP"
         Top             =   180
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
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
      Begin MSMask.MaskEdBox txtAccountNo 
         Height          =   315
         Left            =   1470
         TabIndex        =   15
         Top             =   1410
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "c"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtbranchcode 
         Height          =   315
         Left            =   3810
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         ToolTipText     =   "Default Currency"
         Top             =   180
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin VB.Label Label18 
         Caption         =   "Branch # :"
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         Top             =   165
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Voucher Type :"
         Height          =   210
         Left            =   270
         TabIndex        =   21
         Top             =   210
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Account No :"
         Height          =   210
         Left            =   450
         TabIndex        =   16
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type Class :"
         Height          =   210
         Left            =   495
         TabIndex        =   14
         Top             =   1020
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         Height          =   210
         Left            =   495
         TabIndex        =   10
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Frequency :"
         Height          =   210
         Left            =   525
         TabIndex        =   9
         Top             =   1890
         Width           =   870
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   4875
      _ExtentX        =   8599
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
               Picture         =   "frmVchrType.frx":0760
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVchrType.frx":0BB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVchrType.frx":1008
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVchrType.frx":145C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVchrType.frx":18B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVchrType.frx":1D04
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVchrType.frx":2458
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmVchrType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lb_BlnkMast As Boolean
Dim Mode As String

Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Branch As New Recordset
Dim PR_Detail As New Recordset
Dim PR_GlType As New Recordset


Private Sub cmdLookup_Click()
   
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtvchrType
    Set PO_DESC = txtVchrDesc
    
    GoTop PR_GlType
    PR_GlType.Filter = "BranchCode = '" & txtbranchcode & "'"
    MyLookup.Caption = "Voucher Types"
    MyLookup.FillGrid PR_GlType, "VchrType", "VchrDescrip", 5
    MyLookup.Show 1
    PR_GlType.Filter = adFilterNone
    If Len(txtvchrType) > 0 Then txtVchrType_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Command1_Click()
  Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtAccountNo
    Set PO_DESC = Text1
    Gs_SQL = "Select Accountno 'Account No', Acct_Desc  'Description' from gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
    Gs_OrderBy = "Order by Acct_Desc,AccountNo"
    MyLookupOLDB.Caption = "Account Nos."
    MyLookupOLDB.Show 1
    If Len(txtAccountNo) > 0 Then TxtAccountNo_KeyDown vbKeyReturn, vbKeyShift
    If Len(txtAccountNo) > 0 Then TxtAccountNo_KeyDown vbKeyReturn, vbKeyShift
End Sub
Private Sub Command4_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbranchcode
    Set PO_DESC = Text1
    GoTop PR_Branch
    MyLookup.Caption = "Branches"
    MyLookup.FillGrid PR_Branch, "BranchCode", "BranchDesc", txtbranchcode.MaxLength
    MyLookup.Show 1
    
    If Len(txtbranchcode) > 0 Then txtBranchCode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub txtBranchCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If Lastkey(KeyCode) And txtbranchcode.Text <> "" Then
         txtbranchcode = DoPad(txtbranchcode, 3)
         lb_found = MySeek(txtbranchcode.Text, "BranchCode", PR_Branch)
        
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtbranchcode.SetFocus
         Else
             If txtvchrType.Enabled Then txtvchrType.SetFocus
             If Not txtvchrType.Enabled Then txtVchrDesc.SetFocus
         End If
 ElseIf KeyCode = vbKeyF12 Then
     Call Command4_Click
 End If
End Sub

Private Sub Form_Load()
  SetToolBar(1) = chkRights("GLVCHR0001")
  SetToolBar(2) = chkRights("GLVCHR0002")
  SetToolBar(3) = chkRights("GLVCHR0003")
  SetToolBar(4) = chkRights("GLVCHR0004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  PR_Branch.Open "Select * from SysBranch Where Compcode = '" & Gs_compcode & "' order by Branchcode", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_Detail.Open "SELECT AccountNo, Acct_Desc FROM Gl_Detail WHERE CompCode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockOptimistic, 1
  PR_GlType.Open "Select *,Branchcode+VchrType As FindFld from GlVchrType where glVchrType.CompCode ='" & Gs_compcode & "' order by GlVchrType.VchrType ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  lb_BlnkMast = IIf(PR_GlType.EOF, True, False)
   
  cmdLookup.Enabled = Not (lb_BlnkMast)
  txtbranchcode = Gs_BranchCode
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_Branch.Close
    PR_GlType.Close
    PR_Detail.Close
End Sub

Private Sub optBank_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtAccountNo.SetFocus
End If
End Sub

Private Sub optCash_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtAccountNo.SetFocus
End If
End Sub

Private Sub optJV_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtAccountNo.SetFocus
End If
End Sub

Private Sub TxtAccountNo_KeyDown(KeyCode As Integer, Shift As Integer)
   If Lastkey(KeyCode) And txtAccountNo.Text <> "" Then
         lb_found = MySeek(txtAccountNo, "AccountNo", PR_Detail)
     
         If Not lb_found Then
            Call SetErr(Gs_RecNFMsg, vbCritical)
            txtAccountNo.SetFocus
         Else
            optMonth.SetFocus
         End If
  ElseIf KeyCode = vbKeyF12 Then
        Call Command1_Click
  End If
End Sub

Private Sub txtVchrDesc_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
    optJV.SetFocus
 End If
End Sub

Private Sub txtVchrType_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn Then
 
         PR_GlType.Requery
         txtvchrType.Text = UCase(txtvchrType.Text)
         lb_found = MySeek(txtbranchcode + txtvchrType.Text, "FindFld", PR_GlType)
         
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                  ' 'Cancel = True
                   txtvchrType.SetFocus
                Else
                   txtVchrDesc.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   'Cancel = True
                   txtvchrType.SetFocus
                Else
                   Call SetVal
                   If Mode <> "D" Then
                     ' txtVchrType.Enabled = False
                      txtVchrDesc.SetFocus
                   End If
                End If
            End Select
   ElseIf KeyCode = vbKeyF12 Then
        Call cmdLookup_Click
   End If
  
      
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If lb_BlnkMast And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_GlType, frmVchrType, txtvchrType, txtVchrDesc, "x", "CompCount", 3, "VchrType", "VchrDesc", 1, False, Toolbar1)
       If Mode = "A" Then cmdLookup.Enabled = False Else cmdLookup.Enabled = True
    End If
    
End Sub

Public Sub SaveValues()
On Error GoTo LocalErr
Dim cntsql As New ADODB.Command
lb_BlnkMast = False

           Dim ln_freq As Integer
           If optMonth.Value = True Then
           ln_freq = 1
           Else
           ln_freq = 0
           End If

            Dim ls_Vtype As String
            If optCash.Value = True Then
            ls_Vtype = "C"
            ElseIf optBank.Value = True Then
            ls_Vtype = "B"
            Else
            ls_Vtype = "J"
            End If

cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText
gc_dbcon.BeginTrans
     Select Case Mode
           Case "A"
              cntsql.CommandText = "INSERT into GlVchrType(compcode,BranchCode,vchrtype,vchrdescrip,AccountNo,vchrfrequency,VchrBase,userid,adddate,addtime) VALUES ('" & Gs_compcode & "','" & txtbranchcode & "','" & txtvchrType.Text & "','" & txtVchrDesc.Text & "','" & txtAccountNo & "'," & Val(ln_freq) & ",'" & UCase(ls_Vtype) & "','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "')"
              cntsql.Execute
              
              cntsql.CommandText = "INSERT into Gl_VchrCntrs(compcode,BranchCode,vchrtype,VchrYear) VALUES ('" & Gs_compcode & "','" & txtbranchcode & "','" & txtvchrType.Text & "'," & Year(Gs_Fnperiod) & ")"
              cntsql.Execute
           Case "E"
              cntsql.CommandText = "UPDATE GlVchrType SET vchrdescrip = '" & txtVchrDesc & "',AccountNo = '" & txtAccountNo & "',  vchrfrequency ='" & ln_freq & "',  VchrBase ='" & ls_Vtype & "' WHERE  compcode+Branchcode+vchrtype= '" & Gs_compcode + txtbranchcode + txtvchrType.Text & "'"
              cntsql.Execute
              
           Case "D"
            cntsql.CommandText = "DELETE FROM GlVchrType WHERE compcode+BranchCode = '" & Gs_compcode + txtbranchcode & "'and VchrType='" & txtvchrType.Text & "'"
            cntsql.Execute
            
            cntsql.CommandText = "DELETE FROM Gl_VchrCntrs WHERE compcode+BranchCode = '" & Gs_compcode + txtbranchcode & "'and VchrType='" & txtvchrType.Text & "' And VchrYear = " & Year(Gs_Fnperiod) & " "
            cntsql.Execute
     End Select
gc_dbcon.CommitTrans
PR_GlType.Requery

Exit Sub
LocalErr:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub SetVal()
     txtvchrType = PR_GlType("VchrType")
     txtVchrDesc = PR_GlType("VchrDescrip")
     txtbranchcode = PR_GlType("Branchcode") & ""
     txtAccountNo = PR_GlType("AccountNo") & ""
     
     If PR_GlType("VchrFrequency") = 1 Then
         optMonth.Value = True
     Else
         optYear.Value = True
     End If
     
     If PR_GlType("VchrBase") = "C" Then
         optCash.Value = True
     ElseIf PR_GlType("vchrBase") = "B" Then
         optBank.Value = True
     Else
         optJV.Value = True
     End If
  
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtvchrType.Text) = txtvchrType.MaxLength And txtVchrDesc.Text <> "" Then
       ChkInputs = True
    Else
       Call SetErr("Incomplete Data found", vbCritical)
       ChkInputs = False
    End If
End Function

Private Sub txtcompname_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
    txtcompaddr1.SetFocus
 End If
End Sub

Public Sub FrmRefresh()
  PR_Detail.Requery
  PR_GlType.Requery
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRptRouting 
   Caption         =   "GL. Accounts Routing to Customized Reports"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   Icon            =   "frmRptRouting.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   5130
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   0
      TabIndex        =   1
      Top             =   570
      Width           =   5115
      Begin VB.OptionButton Option1 
         Caption         =   "Net"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1410
         TabIndex        =   10
         Top             =   1530
         Value           =   -1  'True
         Width           =   555
      End
      Begin VB.OptionButton optDebit 
         Caption         =   "Debit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         TabIndex        =   11
         Top             =   1530
         Width           =   675
      End
      Begin VB.OptionButton optCredit 
         Caption         =   "Credit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3030
         TabIndex        =   12
         Top             =   1530
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   4620
         Picture         =   "frmRptRouting.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1080
         Width           =   315
      End
      Begin VB.CommandButton cmdLookup0 
         Height          =   315
         Left            =   1860
         Picture         =   "frmRptRouting.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   300
         Width           =   315
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   2400
         Picture         =   "frmRptRouting.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   690
         Width           =   315
      End
      Begin VB.TextBox txtRptDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2220
         MaxLength       =   64
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   300
         Width           =   2715
      End
      Begin VB.TextBox txtGroupDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2760
         MaxLength       =   64
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   690
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   4770
         MaxLength       =   50
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1530
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSMask.MaskEdBox txtAccountNo 
         Height          =   315
         Left            =   1260
         TabIndex        =   8
         Top             =   1080
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRptCode 
         Height          =   315
         Left            =   1260
         TabIndex        =   2
         Tag             =   "SKIP"
         Top             =   300
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         MaxLength       =   2
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
      Begin MSMask.MaskEdBox txtGroupCode 
         Height          =   315
         Left            =   1260
         TabIndex        =   5
         Tag             =   "SKIP"
         Top             =   690
         Width           =   1155
         _ExtentX        =   2037
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
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   1815
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   3201
         _Version        =   393216
         Cols            =   4
         AllowBigSelection=   0   'False
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Balance Type :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   1500
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Account No :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   255
         TabIndex        =   17
         Top             =   1110
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Report Code :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   16
         Top             =   300
         Width           =   990
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Group Code :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   15
         Top             =   690
         Width           =   960
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5130
      _ExtentX        =   9049
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
               Picture         =   "frmRptRouting.frx":0760
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRptRouting.frx":0BB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRptRouting.frx":1008
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRptRouting.frx":145C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRptRouting.frx":18B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRptRouting.frx":1D04
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRptRouting.frx":2458
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmRptRouting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pb_BlnkVchr As Boolean
Dim Mode As String
Dim Ls_Rptcode As String
Dim PI_CurRow    As Integer
Dim PI_SrNo      As Integer
Dim PS_RowClicked As String
Dim LS_status As String

Public PO_CODE As Object
Public PO_DESC As Object

Dim PR_GlRptRef As New Recordset
Dim PR_GlRptDetl As New Recordset
Dim PR_Gl_Detail As New Recordset
Dim PR_GlrptRouting As New Recordset

Private Sub cmdLookup0_Click()
    Set PO_CODE = Nothing
    Set PO_DESC = Nothing
    Set PO_AnyForm = Nothing
    
    Set PO_AnyForm = Me
    Set PO_CODE = txtRptCode
    Set PO_DESC = txtRptDesc
    
    GoTop PR_GlRptRef
    MyLookup.Caption = "Customized Reports."
    MyLookup.FillGrid PR_GlRptRef, "ReportCode", "RptDescrip", 5
    MyLookup.Show 1
    
    If Len(txtRptCode.Text) > 0 Then txtRptCode_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Command2_Click()
    
    Set PO_CODE = Nothing
    Set PO_DESC = Nothing
    Set PO_AnyForm = Nothing
    
    Set PO_AnyForm = Me
    Set PO_CODE = txtgroupcode
    Set PO_DESC = txtgroupdesc
    
    GoTop PR_GlRptDetl
    MyLookup.Caption = "Group List."
    MyLookup.FillGrid PR_GlRptDetl, "GroupCode", "GroupDesc", 10
    MyLookup.Show 1
    
    If Len(txtgroupcode.Text) > 0 Then txtgroupcode_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccountno
    Set PO_DESC = Text1
    
    GoTop PR_Gl_Detail
    MyLookup.Caption = "Account Nos."
    MyLookup.FillGrid PR_Gl_Detail, "AccountNo", "Acct_Desc", Len(PR_Gl_Detail.Fields("AccountNo"))
    MyLookup.Show 1
    
    If Len(txtaccountno.Text) > 0 Then txtAccountNo_KeyDown vbKeyReturn, vbKeyShift
    
End Sub

Private Sub Form_Load()
Ls_Rptcode = ""

  SetToolBar(1) = chkRights("GLRPTRUT01")
  SetToolBar(2) = chkRights("GLRPTRUT02")
  SetToolBar(3) = chkRights("GLRPTRUT03")
  SetToolBar(4) = chkRights("GLRPTRUT04")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)

  PR_GlRptRef.Open "Select * from GlRpts_Ref where CompCode ='" & Gs_compcode & "' and ReportBase = 'O'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_GlRptDetl.Open "Select *,ReportCode+GroupCode as RptDtlSeek from GlGroupDetl where CompCode ='" & Gs_compcode & "' Order by Reportcode,GroupCode", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_Gl_Detail.Open "Select * from Gl_Detail where compcode='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_GlrptRouting.Open "Select *,ReportCode+Groupcode As FindFld from Gl_RptRouting where compcode='" & Gs_compcode & "' Order by Reportcode,Groupcode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  
  Pb_BlnkVchr = IIf(PR_GlRptRef.EOF, True, False)
  
  InitializeGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
  PR_GlRptDetl.Close
  PR_Gl_Detail.Close
  PR_GlrptRouting.Close
  PR_GlRptRef.Close
End Sub
Private Sub grid1_DblClick()
    With Grid1
        If .Row > 0 Then
            PI_CurRow = .Row
        End If
        txtgroupcode = .TextMatrix(.Row, 1)
        txtaccountno = .TextMatrix(.Row, 2)
        PS_RowClicked = "Y"
    End With

End Sub

Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
    With Grid1
        If KeyCode = vbKeyDelete Then
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            .RemoveItem .Row
        End If
    End With

End Sub

Private Sub optCredit_Click()
  LS_status = "C"
  Call AddGrid
  txtgroupcode.SetFocus
End Sub

Private Sub optDebit_Click()
  LS_status = "D"
  Call AddGrid
  txtgroupcode.SetFocus
End Sub

Private Sub Option1_Click()
  LS_status = "N"
  Call AddGrid
  txtgroupcode.SetFocus
End Sub

Private Sub txtAccountNo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
Dim Ln_Found As Integer
Ln_Found = 0
If KeyCode = vbKeyReturn And txtaccountno <> "" Then
        lb_found = MySeek(txtaccountno, "AccountNo", PR_Gl_Detail)
        
        If Not lb_found Then
            Call SetErr(Gs_RecNFMsg, vbCritical)
            txtaccountno.SetFocus
        Else
         With Grid1
            If Val(txtaccountno) > 0 Then
                For ln_cnt = 1 To .Rows - 1
                    If .TextMatrix(ln_cnt, 2) = txtaccountno Then
                        Call SetErr("Account No. already exists.", vbCritical)
                        lb_found = True
                        txtaccountno.SetFocus
                        Exit Sub
                    End If
                Next
            End If
          End With
          Grid1.SetFocus
        End If
ElseIf KeyCode = vbKeyF12 Then
     Command1_Click
End If
End Sub

Private Sub txtAccountNo_LostFocus()
    Call txtAccountNo_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub txtgroupcode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

If Lastkey(KeyCode) And txtgroupcode.Text <> "" Then
        txtgroupcode.Text = DoPad(txtgroupcode.Text, 10)
        lb_found = MySeek(txtRptCode + txtgroupcode, "RptDtlSeek", PR_GlRptDetl)
        
        If lb_found Then
            txtgroupdesc.Text = PR_GlRptDetl("groupdesc")
            txtaccountno.SetFocus
        Else
            Call SetErr(Gs_RecNFMsg, vbCritical)
        End If
ElseIf KeyCode = vbKeyF12 Then
     Command2_Click
End If

End Sub
Private Sub txtRptCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

If KeyCode = vbKeyReturn And txtRptCode <> "" Then

        txtRptCode = DoPad(txtRptCode.Text, 2)
        lb_found = MySeek(txtRptCode.Text, "ReportCode", PR_GlRptRef)
        
        If Not lb_found Then
          Call SetErr(Gs_RecNFMsg, vbCritical)
          txtRptCode.SetFocus
          Exit Sub
        Else
             PR_GlRptDetl.Filter = "Reportcode = '" & txtRptCode.Text & "'"
             GoTop PR_GlRptDetl
        End If
        
        lb_found = MySeek(txtRptCode.Text, "ReportCode", PR_GlrptRouting)
        Select Case Mode
          Case "A"
            If lb_found Then
               Call SetErr(Gs_RecFdMsg, vbCritical)
               txtRptCode.SetFocus
            Else
               txtgroupcode.SetFocus
            End If
          Case Else
            If Not lb_found Then
               Call SetErr(Gs_RecNFMsg, vbCritical)
               txtRptCode.SetFocus
            Else
               txtRptDesc.Text = PR_GlRptRef("rptdescrip")
               LoadGRNTrans
               If Mode <> "D" Then
                 txtgroupcode.SetFocus
               End If
            End If
        End Select
ElseIf KeyCode = vbKeyF12 Then
     cmdLookup0_Click
End If

  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    If Pb_BlnkVchr And Range(Button.Index, 2, 3) Then
       Call SetErr("Data not found.", vbCritical)
       Mode = ""
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_GlRptDetl, frmRptRouting, txtRptCode, txtgroupcode, "x", "CompCount", 3, "ReportCode", "GroupCode", 1, False, Toolbar1)
    End If
        

End Sub

Public Sub SaveValues()
On Error GoTo LocalErr
Dim cntsql As New ADODB.Command
PB_BlnkComp = False

cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText

gc_dbcon.BeginTrans
     Select Case Mode
           Case "A"
              With Grid1
                .Row = 1
              For ln_cnt = 1 To Grid1.Rows - 1
                 cntsql.CommandText = "INSERT into Gl_RptRouting(compcode,reportcode,groupCode,accountNo,Acct_status) VALUES ('" & Gs_compcode & "','" & txtRptCode.Text & "','" & .TextMatrix(ln_cnt, 1) & "','" & .TextMatrix(ln_cnt, 2) & "','" & .TextMatrix(ln_cnt, 3) & "')"
                 cntsql.Execute
              Next
              End With
           Case "E"
            cntsql.CommandText = "DELETE FROM Gl_RptRouting WHERE compcode = '" & Gs_compcode & "'and reportcode='" & txtRptCode.Text & "'"
            cntsql.Execute
              
              With Grid1
                .Row = 1
              For ln_cnt = 1 To Grid1.Rows - 1
                  cntsql.CommandText = "INSERT into Gl_RptRouting(compcode,reportcode,groupCode,accountno,Acct_status) VALUES ('" & Gs_compcode & "','" & txtRptCode.Text & "','" & .TextMatrix(ln_cnt, 1) & "','" & .TextMatrix(ln_cnt, 2) & "','" & .TextMatrix(ln_cnt, 3) & "')"
                  cntsql.Execute
              Next
              End With
           Case "D"
                cntsql.CommandText = "DELETE FROM Gl_RptRouting WHERE compcode = '" & Gs_compcode & "'and reportcode='" & txtRptCode.Text & "'"
                cntsql.Execute

     End Select
     
gc_dbcon.CommitTrans
PR_GlrptRouting.Requery
     
     PI_SrNo = 0
     PS_RowClicked = ""
     InitializeGrid
     
Exit Sub
LocalErr:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Private Sub SetVal()
On Error GoTo LocalErr
InitializeGrid

Exit Sub
LocalErr:
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtRptCode.Text) > 0 And PI_SrNo > 0 Then
       ChkInputs = True
    Else
       Call SetErr("Incomplete Data found", vbCritical)
       ChkInputs = False
    End If
End Function
Public Sub InitializeGrid()
    With Grid1
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Group Code |<Account No     |<Type  "
        .ColWidth(1) = 1500
        .ColWidth(2) = 2000
        .ColWidth(3) = 700
        .Redraw = True
    End With
End Sub

Private Sub LoadGRNTrans()
Dim lb_found As Boolean
Dim ln_cnt   As Integer
Dim temp As String
ln_cnt = 1
    
    lb_found = MySeek(txtRptCode.Text, "ReportCode", PR_GlrptRouting)

    If lb_found Then
        With Grid1
            Do While LTrim(RTrim(PR_GlrptRouting("Reportcode").Value)) = txtRptCode.Text
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = ln_cnt
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = Trim(PR_GlrptRouting.Fields("Groupcode"))
                .TextMatrix(.Row, 2) = Trim(PR_GlrptRouting.Fields("AccountNo"))
                .TextMatrix(.Row, 3) = PR_GlrptRouting.Fields("Acct_Status")
                .Rows = .Rows + 1
                ln_cnt = ln_cnt + 1
                PR_GlrptRouting.MoveNext
                If PR_GlrptRouting.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
            txtRptCode.SetFocus
        End With
    Else
        Call SetErr("Transactions not found.", vbCritical)
        txtRptCode.SetFocus
    End If
End Sub

Private Sub AddGrid()
If txtgroupcode.Text <> "" And txtaccountno <> "" Then

        If PS_RowClicked = "" Then
            If PI_SrNo = 0 Then
                PI_SrNo = 1
            Else
                PI_SrNo = PI_SrNo + 1
            End If
        End If
        
            With Grid1
                If PS_RowClicked = "" Then
                    If Not PI_SrNo = 1 Then .Rows = .Rows + 1
                    .Row = .Rows - 1
                Else
                    .Row = PI_CurRow
                End If
                
                If PS_RowClicked = "" Then
                    .TextMatrix(.Row, 0) = PI_SrNo
                Else
                    .TextMatrix(.Row, 0) = PI_CurRow
                End If

                .TextMatrix(.Row, 1) = Trim(txtgroupcode.Text)
                .TextMatrix(.Row, 2) = Trim(txtaccountno.Text)
                .TextMatrix(.Row, 3) = LS_status
                
                 PS_RowClicked = ""
                 txtaccountno = ""
                 LS_status = ""
              End With
Option1.Value = False
optDebit.Value = False
optCredit.Value = False
End If
End Sub


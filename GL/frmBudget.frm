VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBudget 
   Caption         =   "Budgeting "
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   Icon            =   "frmBudget.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   0
      TabIndex        =   1
      Top             =   585
      Width           =   5505
      Begin VB.TextBox txtPeriod 
         BackColor       =   &H00FFFF00&
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
         Height          =   315
         Left            =   1260
         MaxLength       =   4
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   300
         Width           =   615
      End
      Begin MSMask.MaskEdBox txtDebit 
         Height          =   315
         Left            =   1260
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1275
         _ExtentX        =   2249
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
      Begin VB.TextBox txtAccountno 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   705
         Width           =   3735
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   315
         Left            =   5010
         Picture         =   "frmBudget.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   705
         Width           =   315
      End
      Begin VB.TextBox txtaccdesc 
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
         Left            =   1260
         MaxLength       =   64
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1080
         Width           =   4080
      End
      Begin MSMask.MaskEdBox txtCredit 
         Height          =   315
         Left            =   4080
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1275
         _ExtentX        =   2249
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
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
         TabIndex        =   12
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label Label2 
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
         Left            =   195
         TabIndex        =   11
         Top             =   780
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Period :"
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
         Left            =   600
         TabIndex        =   10
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Debit Amount :"
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
         Left            =   90
         TabIndex        =   9
         Top             =   1620
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Credit Amount :"
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
         Left            =   2940
         TabIndex        =   8
         Top             =   1620
         Width           =   1110
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
         Left            =   2880
         Top             =   -120
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
               Picture         =   "frmBudget.frx":047C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBudget.frx":08D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBudget.frx":0D24
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBudget.frx":1178
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBudget.frx":15CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBudget.frx":1A20
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBudget.frx":2174
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmBudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pb_BlnkVchr As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_GlBudget As Recordset
Dim PR_GlDetail As Recordset

Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtAccountNo
    Set PO_DESC = txtaccdesc
    
    If Mode = "A" Then
        GoTop PR_GlDetail
        MyLookup.Caption = "Account Nos."
        MyLookup.FillGrid PR_GlDetail, "AccountNo", "Acct_Desc", Len(PR_GlDetail.Fields("AccountNo"))
    Else
        GoTop PR_GlBudget
        MyLookup.Caption = "Account Nos."
        MyLookup.FillGrid PR_GlBudget, "AccountNo", "Acct_Desc", Len(PR_GlDetail.Fields("AccountNo"))
    End If
    MyLookup.Show 1
    
    If Len(txtAccountNo.Text) > 0 Then txtAccountNo_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Form_Load()
  txtPeriod = LTrim(Str(Year(Gs_Fnperiod)))
  
  SetToolBar(1) = chkRights("GLBUDGT001")
  SetToolBar(2) = chkRights("GLBUDGT002")
  SetToolBar(3) = chkRights("GLBUDGT003")
  SetToolBar(4) = chkRights("GLBUDGT004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  Set PR_GlBudget = New Recordset
  Set PR_GlDetail = New Recordset

'  PR_GlBudget.Open "Select Gl_budget.*,Fn_Year+AccountNo As SeekKey from Gl_budget where CompCode ='" & Gs_compcode & "' order by 6", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  
  PR_GlBudget.Open "SELECT Gl_budget.AccountNo,Gl_budget.Fn_Year, Gl_budget.Dr_amount, Gl_budget.Cr_amount, Gl_budget.Fn_Year+Gl_budget.AccountNo As SeekKey,Gl_Detail.Acct_Desc FROM Gl_budget INNER JOIN Gl_Detail ON (Gl_budget.AccountNo = Gl_Detail.AccountNo) AND (Gl_budget.Compcode = Gl_Detail.compcode) where Gl_Detail.CompCode ='" & Gs_compcode & "' order by 5", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_GlDetail.Open "Select * from Gl_detail where CompCode ='" & Gs_compcode & "'", gc_dbcon, adOpenDynamic, adLockOptimistic, 1

  
  Pb_BlnkVchr = IIf(PR_GlBudget.EOF, True, False)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_GlDetail.Close
    PR_GlBudget.Close
End Sub

Private Sub txtAccountNo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If LastKey(KeyCode) Then
      lb_found = MySeek(txtPeriod + txtAccountNo, "SeekKey", PR_GlBudget)
      
      Select Case Mode
          Case "A"
               If lb_found Then
                  Call SetErr(Gs_RecFdMsg, vbCritical)
                  txtAccountNo.SetFocus
               Else
                  lb_found = MySeek(txtAccountNo, "AccountNo", PR_GlDetail)
                  If Not lb_found Then
                     Call SetErr(Gs_RecNFMsg, vbCritical)
                     txtAccountNo.SetFocus
                  Else
                     txtaccdesc.Text = PR_GlDetail.Fields("Acct_Desc")
                     txtCredit.Enabled = False
                     txtDebit.SetFocus
                  End If
               End If
          Case Else
                  If Not lb_found Then
                     Call SetErr(Gs_RecNFMsg, vbCritical)
                     txtAccountNo.SetFocus
                  Else
                     Call SetVal
                     lb_found = MySeek(txtAccountNo, "AccountNo", PR_GlDetail)
                     txtaccdesc.Text = PR_GlDetail.Fields("Acct_Desc")
                     
                     If Mode = "E" Then
                        If Val(txtDebit.Text) > 0 Then
                           txtDebit.SetFocus
                        Else
                           txtCredit.SetFocus
                        End If
                     End If
                  End If
      End Select
  ElseIf KeyCode = vbKeyF12 Then
       cmdLookup_Click
  End If
End Sub

Private Sub txtCredit_KeyDown(KeyCode As Integer, Shift As Integer)
 If LastKey(KeyCode) Then
    If Val(txtCredit.Text) <= 0 Then
       txtCredit.Text = ""
       txtCredit.Enabled = False
       txtDebit.Enabled = True
       txtDebit.SetFocus
       txtDebit.Text = ""
    End If
 End If
End Sub

Private Sub txtDebit_KeyDown(KeyCode As Integer, Shift As Integer)
 If LastKey(KeyCode) Then
    If Val(txtDebit.Text) <= 0 Then
       txtDebit.Text = ""
       txtDebit.Enabled = False
       txtCredit.Enabled = True
       txtCredit.SetFocus
       txtCredit.Text = ""
    End If
 End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    If Pb_BlnkVchr And Range(Button.Index, 2, 3) Then
       Call SetErr("Data not found,", vbCritical)
       Mode = ""
    Else
       Mode = DentMode(Mode, Button.Index, PR_GlBudget, frmBudget, txtAccountNo, txtAccountNo, "Compcount", "x", 3, "txtAccountNo", "TxtAccountNo", 1, False, Toolbar1)
    End If
        
End Sub

Public Sub SaveValues()
Dim cntsql As New ADODB.Command
Pb_BlnkVchr = False

cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText

     Select Case Mode
           Case "A"
              cntsql.CommandText = "INSERT into Gl_budget(compcode,accountno,fn_year,dr_amount,cr_amount) VALUES ('" & Gs_compcode & "','" & txtAccountNo.Text & "','" & txtPeriod.Text & "','" & Val(txtDebit.Text) & "','" & Val(txtCredit.Text) & "')"
              cntsql.Execute
           Case "E"
              cntsql.CommandText = "update Gl_budget set dr_amount='" & Val(txtDebit.Text) & "',cr_amount= '" & Val(txtCredit.Text) & "'"
              cntsql.Execute
              
           Case "D"
              cntsql.CommandText = "DELETE FROM Gl_budget WHERE compcode = '" & Gs_compcode & "'and accountno='" & txtAccountNo.Text & "' and fn_year='" & txtPeriod.Text & "'"
              cntsql.Execute
     End Select
     PR_GlBudget.Requery
End Sub

Private Sub SetVal()
     txtAccountNo.Text = PR_GlBudget("accountno")
     txtDebit.Text = PR_GlBudget("dr_amount")
     txtCredit.Text = PR_GlBudget("cr_amount")
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtAccountNo.Text) > 0 And (Val(txtDebit.Text) > 0 Or Val(txtCredit.Text) > 0) Then
       ChkInputs = True
    Else
       Call SetErr("Incomplete Data found", vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub FrmRefresh()
  PR_GlBudget.Requery
  PR_GlDetail.Requery
End Sub

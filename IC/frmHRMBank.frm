VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHRMBank 
   Caption         =   "Bank Setup"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHRMBank.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5745
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
      Height          =   1800
      Left            =   15
      TabIndex        =   1
      Top             =   570
      Width           =   5685
      Begin VB.TextBox txtglDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3135
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1335
         Width           =   2445
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
         Left            =   2790
         Picture         =   "frmHRMBank.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1335
         Width           =   315
      End
      Begin VB.TextBox txtGlCode 
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1350
         Width           =   1470
      End
      Begin VB.TextBox Textx 
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
         Width           =   4290
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   600
         Width           =   4290
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
         Picture         =   "frmHRMBank.frx":047C
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "GL Code :"
         Height          =   210
         Left            =   555
         TabIndex        =   12
         Top             =   1380
         Width           =   720
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
      Width           =   5745
      _ExtentX        =   10134
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
               Picture         =   "frmHRMBank.frx":05EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHRMBank.frx":0A42
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHRMBank.frx":0E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHRMBank.frx":12EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHRMBank.frx":173E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHRMBank.frx":1B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmHRMBank.frx":22E6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmHRMBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkLoca As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_VchType As New Recordset
Dim pr_dumy As New Recordset
Dim PR_Bank As New Recordset

Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtLocation
    Set PO_DESC = txtdesc
    
    GoTop PR_Bank
    MyLookup.Caption = "Banks"
    MyLookup.FillGrid PR_Bank, "BankCode", "BankName", txtLocation.MaxLength
    MyLookup.Show 1
    
    If Len(txtLocation) > 0 Then TxtLocation_KeyDown vbKeyReturn, vbKeyShift

End Sub


Private Sub Command1_Click()

    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtGlCode
    Set PO_DESC = txtglDesc
    Gs_SQL = "Select  Accountno 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtGlCode) > 0 Then txtglcode_KeyDown vbKeyReturn, vbKeyShift

End Sub


Private Sub txtglcode_KeyDown(KeyCode As Integer, Shift As Integer)
If txtGlCode <> "" Then
    
    
        ls_sql = "Select Accountno 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  accountno = '" & txtGlCode & "' "
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
                ''Cancel = True
            Else
                txtglDesc = pr_dumy("description")
            End If
         pr_dumy.Close
    
End If
End Sub



Private Sub Form_Load()
  
  SetToolBar(1) = chkRights("SRBST00001")
  SetToolBar(2) = chkRights("SRBST00002")
  SetToolBar(3) = chkRights("SRBST00003")
  SetToolBar(4) = chkRights("SRBST00004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  

  PR_Bank.Open "Select * from SysBanks where compcode = '" & Gs_compcode & "' Order By BankCode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1

  
  PB_BlnkLoca = IIf(PR_Bank.EOF, True, False)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_Bank.Close
End Sub


Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) Then
     TxtManager.SetFocus
  End If
End Sub

Private Sub txtGlCode_LostFocus()
If Len(Trim(txtGlCode)) > 0 Then
    Call txtglcode_KeyDown(vbKeyReturn, vbKeyShift)
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
                   ''Cancel = True
                    SetClear Me
                   txtLocation.SetFocus
                Else
                   txtdesc.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   'Cancel = True
                    SetClear Me
                   txtLocation.SetFocus
                Else
                   Call SetVal
                   If Mode <> "D" Then
                    '  txtLocation.Enabled = False
                      txtdesc.SetFocus
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
       'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_Bank, Me, txtLocation, txtdesc, "X", "CompCount", 3, "BankCode", "BankName", 1, False, Toolbar1)
    End If
    
End Sub

Public Sub SaveValues()
Dim cntsql As New ADODB.Command
PB_BlnkLoca = False
Dim ls_btype As String

cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText
ls_btype = "L"
     Select Case Mode
           Case "A"
              cntsql.CommandText = "INSERT into SysBanks(Compcode,BankCode,BankName,ManagerName,BankType,GlCode) VALUES ('" & Gs_compcode & "','" & txtLocation.Text & "','" & txtdesc.Text & "','" & TxtManager.Text & "','" & ls_btype & "','" & txtGlCode & "')"
              cntsql.Execute
           Case "E"
              cntsql.CommandText = "UPDATE SysBanks SET BankName= '" & txtdesc.Text & "',ManagerName = '" & TxtManager & "',BankType = '" & ls_btype & "',GlCode = '" & txtGlCode & "' WHERE  BankCode= '" & txtLocation.Text & "' and compcode = '" & Gs_compcode & "'"
              cntsql.Execute
           Case "D"
              cntsql.CommandText = "DELETE FROM SysBanks WHERE BankCode = '" & txtLocation.Text & "' and compcode= '" & Gs_compcode & "'"
              cntsql.Execute
     End Select
PR_Bank.Requery
End Sub

Private Sub SetVal()
     txtdesc = PR_Bank("BankName") & ""
     TxtManager = PR_Bank("ManagerName") & ""
     txtGlCode = Trim(PR_Bank("GLCode") & "")
         If Len(txtGlCode) > 0 Then txtglcode_KeyDown vbKeyReturn, vbKeyShift
     
   '  txtVchrType = PR_Bank("BankVchrtype") & ""
   '  txtadjVchrType = PR_Bank("adjVchrtype") & ""
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtLocation.Text) = txtLocation.MaxLength And txtdesc <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub FrmRefresh()
  PR_GlDetail.Requery
  PR_Bank.Requery
  PR_VchType.Requery
End Sub

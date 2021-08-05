VERSION 5.00
Begin VB.Form frmImportGlHeads 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Company Name"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   Icon            =   "frmImportGLHeads.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Insert Voucher Types"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   930
      Left            =   45
      TabIndex        =   24
      Top             =   4380
      Width           =   4995
      Begin VB.CheckBox ChkVoucherType 
         Caption         =   "Insert Voucher Types"
         Height          =   420
         Left            =   165
         TabIndex        =   26
         Top             =   300
         Width           =   3525
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Insert"
         Height          =   360
         Left            =   3840
         TabIndex        =   25
         Top             =   315
         Width           =   1065
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Run All Above Process"
      Height          =   360
      Left            =   2775
      TabIndex        =   21
      Top             =   6390
      Width           =   2190
   End
   Begin VB.Frame Frame5 
      Caption         =   "Insert from Company"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   60
      TabIndex        =   16
      Top             =   5400
      Width           =   4995
      Begin VB.TextBox txtCompdesc 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2130
         MaxLength       =   64
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   390
         Width           =   2790
      End
      Begin VB.TextBox txtfromCompcode 
         Height          =   300
         Left            =   1275
         MaxLength       =   3
         TabIndex        =   19
         Tag             =   "SKIPN"
         Top             =   390
         Width           =   510
      End
      Begin VB.CommandButton cmdLookup0 
         Height          =   315
         Left            =   1800
         Picture         =   "frmImportGLHeads.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   390
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Company Code:"
         Height          =   195
         Left            =   135
         TabIndex        =   17
         Top             =   420
         Width           =   1125
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Profit and Loss"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   930
      Left            =   45
      TabIndex        =   13
      Top             =   3390
      Width           =   4995
      Begin VB.ComboBox txtPL 
         Height          =   315
         ItemData        =   "frmImportGLHeads.frx":047C
         Left            =   855
         List            =   "frmImportGLHeads.frx":048C
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   375
         Width           =   2790
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Insert"
         Height          =   360
         Left            =   3840
         TabIndex        =   14
         Top             =   345
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Base On:"
         Height          =   210
         Left            =   135
         TabIndex        =   15
         Top             =   405
         Width           =   675
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Balance Sheet"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   870
      Left            =   45
      TabIndex        =   9
      Top             =   2430
      Width           =   4995
      Begin VB.CommandButton Command5 
         Caption         =   "&Insert"
         Height          =   360
         Left            =   3840
         TabIndex        =   11
         Top             =   345
         Width           =   1065
      End
      Begin VB.ComboBox txtBS 
         Height          =   315
         ItemData        =   "frmImportGLHeads.frx":04C2
         Left            =   855
         List            =   "frmImportGLHeads.frx":04D2
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   375
         Width           =   2790
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Base On:"
         Height          =   210
         Left            =   135
         TabIndex        =   12
         Top             =   405
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Insert Chart of Accouts"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   885
      Left            =   45
      TabIndex        =   5
      Top             =   1440
      Width           =   4995
      Begin VB.ComboBox txtlevel 
         Height          =   315
         ItemData        =   "frmImportGLHeads.frx":0508
         Left            =   855
         List            =   "frmImportGLHeads.frx":051B
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   375
         Width           =   2790
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Insert"
         Height          =   360
         Left            =   3840
         TabIndex        =   6
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Base On:"
         Height          =   210
         Left            =   135
         TabIndex        =   8
         Top             =   405
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Deletion all heads "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1425
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   4995
      Begin VB.CheckBox ChkGLTran 
         Caption         =   "GL Transaction"
         Height          =   300
         Left            =   2385
         TabIndex        =   22
         Top             =   315
         Width           =   2265
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Delete"
         Height          =   360
         Left            =   3840
         TabIndex        =   4
         Top             =   975
         Width           =   1065
      End
      Begin VB.CheckBox ChkPL 
         Caption         =   "Profit and Loss"
         Height          =   300
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   2265
      End
      Begin VB.CheckBox ChkBS 
         Caption         =   "Balance Sheet"
         Height          =   300
         Left            =   480
         TabIndex        =   2
         Top             =   675
         Width           =   2265
      End
      Begin VB.CheckBox chkcoa 
         Caption         =   "Chart of Account"
         Height          =   300
         Left            =   480
         TabIndex        =   1
         Top             =   300
         Width           =   2265
      End
   End
End
Attribute VB_Name = "frmImportGlHeads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ls_sql As String
Dim pr_dumy As New Recordset
Public PO_CODE As Object
Public PO_DESC As Object

Private Sub cmdLookup0_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtfromCompcode
    Set PO_DESC = txtCompdesc
        
        Gs_SQL = "SELECT  Compcode,Compname FROM syscomp "
        Gs_FindFld = "compname"
        Gs_OrderBy = "Order by compname"
        MyLookupOLDB.Caption = "System Company"
        MyLookupOLDB.Show 1
        
        If Trim(txtfromCompcode) = Gs_compcode Then
            Call MsgBox("Same company code not allowed", vbCritical)
            txtCompdesc = ""
            txtfromCompcode = ""
        End If
        SendKeys "{Tab}"
End Sub

Private Sub Command1_Click()
Call DeleteProcess
End Sub

Private Sub DeleteProcess()
If chkcoa.Value = 1 Then
    ls_sql = "delete from gl_sub0 where compcode = '" & Gs_compcode & "'"
    gc_dbcon.Execute ls_sql
    ls_sql = "delete from gl_sub1 where compcode = '" & Gs_compcode & "'"
    gc_dbcon.Execute ls_sql
    ls_sql = "delete from gl_sub2 where compcode = '" & Gs_compcode & "'"
    gc_dbcon.Execute ls_sql
    ls_sql = "delete from gl_Detail where compcode = '" & Gs_compcode & "'"
    gc_dbcon.Execute ls_sql
End If


If ChkGLTran.Value = 1 Then
    ls_sql = "delete from gl_trans where compcode = '" & Gs_compcode & "'"
    gc_dbcon.Execute ls_sql
    ls_sql = "delete from gl_ref where compcode = '" & Gs_compcode & "'"
    gc_dbcon.Execute ls_sql
End If


If ChkPL.Value = 1 Then
    ls_sql = "delete from GL_PLSheet1 where compcode = '" & Gs_compcode & "'"
    gc_dbcon.Execute ls_sql
    ls_sql = "delete from GL_PLSheet2 where compcode = '" & Gs_compcode & "'"
    gc_dbcon.Execute ls_sql
    ls_sql = "delete from GL_PLSheet3 where compcode = '" & Gs_compcode & "'"
    gc_dbcon.Execute ls_sql
    ls_sql = "delete from GL_PLSheet3Detail where compcode = '" & Gs_compcode & "'"
    gc_dbcon.Execute ls_sql
End If


If ChkBS.Value = 1 Then
    ls_sql = "delete from GL_BSheet1 where compcode = '" & Gs_compcode & "'"
    gc_dbcon.Execute ls_sql
    ls_sql = "delete from GL_BSheet2 where compcode = '" & Gs_compcode & "'"
    gc_dbcon.Execute ls_sql
    ls_sql = "delete from GL_BSheet3 where compcode = '" & Gs_compcode & "'"
    gc_dbcon.Execute ls_sql
    ls_sql = "delete from GL_BSheet3Detail where compcode = '" & Gs_compcode & "'"
    gc_dbcon.Execute ls_sql
End If


If chkcoa.Value = 1 Or ChkBS.Value = 1 Or ChkPL.Value = 1 Or ChkGLTran.Value = 1 Then
Call MsgBox("Chart of Account/ GL Transaction successfully deleted", vbInformation)
End If

End Sub
Private Sub InsertCOA()

If txtlevel.Text = "All" Then
    ls_sql = "insert into gl_sub0 (Compcode, Acct_sub0, Acct_Desc, Sub1_Cunt, UserId, AddDate, AddTime) select  '" & Gs_compcode & "' as  Compcode, Acct_sub0, Acct_Desc, Sub1_Cunt, '" & Gc_UserId & "' as  UserId, '" & Format(Date, "YYYY/MM/DD") & "' as   AddDate, AddTime from gl_sub0 where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql
    
    ls_sql = "insert into gl_sub1 (Compcode, Acct_sub0,Acct_sub1, Acct_Desc, Sub2_Cunt, UserId, AddDate, AddTime) select  '" & Gs_compcode & "' as  Compcode, Acct_sub0, Acct_sub1, Acct_Desc, Sub2_Cunt, '" & Gc_UserId & "' as  UserId, '" & Format(Date, "YYYY/MM/DD") & "' as   AddDate, AddTime from gl_sub1 where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql
    
    ls_sql = "insert into gl_sub2 (Compcode, Acct_sub1,Acct_sub2, Acct_Desc, Sub3_Cunt, UserId, AddDate, AddTime) select  '" & Gs_compcode & "' as  Compcode, Acct_sub1, Acct_sub2, Acct_Desc, Sub3_Cunt, '" & Gc_UserId & "' as  UserId, '" & Format(Date, "YYYY/MM/DD") & "' as   AddDate, AddTime from gl_sub2 where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql
    
    ls_sql = "insert into gl_Detail (compcode, Acct_Sub, Acct_Detail, AccountNo, Acct_Desc, Acct_Type, Acct_Base, Acct_Status, Crncy_Code, Bs_DrLineNo, Bs_CrLineNo, Pf_DrLineNo,  Pf_CrLineNo, OldAccount, UserId, AddDate, AddTime) select  '" & Gs_compcode & "' as  Compcode, Acct_Sub, Acct_Detail, AccountNo, Acct_Desc, Acct_Type, Acct_Base, Acct_Status, Crncy_Code, Bs_DrLineNo, Bs_CrLineNo, Pf_DrLineNo,  Pf_CrLineNo, OldAccount, '" & Gc_UserId & "' as  UserId, '" & Format(Date, "YYYY/MM/DD") & "' as   AddDate, AddTime from gl_detail where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql
End If


If txtlevel.Text = "Control Accounts" Then
    ls_sql = "insert into gl_sub0 (Compcode, Acct_sub0, Acct_Desc, Sub1_Cunt, UserId, AddDate, AddTime) select  '" & Gs_compcode & "' as  Compcode, Acct_sub0, Acct_Desc, Sub1_Cunt, '" & Gc_UserId & "' as  UserId, '" & Format(Date, "YYYY/MM/DD") & "' as   AddDate, AddTime from gl_sub0 where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql
End If

If txtlevel.Text = "Detail Accounts" Then
    ls_sql = "insert into gl_sub1 (Compcode, Acct_sub0,Acct_sub1, Acct_Desc, Sub2_Cunt, UserId, AddDate, AddTime) select  '" & Gs_compcode & "' as  Compcode, Acct_sub0, Acct_sub1, Acct_Desc, Sub2_Cunt, '" & Gc_UserId & "' as  UserId, '" & Format(Date, "YYYY/MM/DD") & "' as   AddDate, AddTime from gl_sub1 where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql
End If

If txtlevel.Text = "Sub Ledger Accounts" Then
    ls_sql = "insert into gl_sub2 (Compcode, Acct_sub1,Acct_sub2, Acct_Desc, Sub3_Cunt, UserId, AddDate, AddTime) select  '" & Gs_compcode & "' as  Compcode, Acct_sub1, Acct_sub2, Acct_Desc, Sub3_Cunt, '" & Gc_UserId & "' as  UserId, '" & Format(Date, "YYYY/MM/DD") & "' as   AddDate, AddTime from gl_sub2 where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql
End If

If txtlevel.Text = "Subsidiary Sub Ledger A/c" Then
    ls_sql = "insert into gl_Detail (compcode, Acct_Sub, Acct_Detail, AccountNo, Acct_Desc, Acct_Type, Acct_Base, Acct_Status, Crncy_Code, Bs_DrLineNo, Bs_CrLineNo, Pf_DrLineNo,  Pf_CrLineNo, OldAccount, UserId, AddDate, AddTime) select  '" & Gs_compcode & "' as  Compcode, Acct_Sub, Acct_Detail, AccountNo, Acct_Desc, Acct_Type, Acct_Base, Acct_Status, Crncy_Code, Bs_DrLineNo, Bs_CrLineNo, Pf_DrLineNo,  Pf_CrLineNo, OldAccount, '" & Gc_UserId & "' as  UserId, '" & Format(Date, "YYYY/MM/DD") & "' as   AddDate, AddTime from gl_detail where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql
End If


If txtlevel.Text <> "" Then
    Call MsgBox("Chart of Account successfully Inserted", vbInformation)
End If

End Sub
Private Sub InsertBS()
If txtBS.Text = "All" Then
    ls_sql = "Insert into GL_BSheet1 (Compcode,Bcode,BDesc) select '" & Gs_compcode & "' as Compcode,bcode,bdesc from GL_BSheet1 where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql

    ls_sql = "Insert into GL_BSheet2 (Compcode,Bcode,BNCode,BNDesc) select '" & Gs_compcode & "' as Compcode,bcode,BnCode,BNdesc from GL_BSheet2 where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql

    ls_sql = "Insert into GL_BSheet3 (Compcode,Bcode,BnCode,BnIcode,BniDesc) select '" & Gs_compcode & "' as Compcode,bcode,bncode,bnicode,bnidesc from GL_BSheet3 where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql
    
    ls_sql = "Insert into GL_BSheet3Detail (Compcode,Bcode,bncode,bnicode,accountno) select '" & Gs_compcode & "' as Compcode,Bcode,bncode,bnicode,accountno from GL_BSheet3Detail where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql

End If

If txtBS.Text = "Heads Setup" Then
    ls_sql = "Insert into GL_BSheet1 (Compcode,Bcode,BDesc) select '" & Gs_compcode & "' as Compcode,bcode,bdesc from GL_BSheet1 where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql

End If

If txtBS.Text = "Notes Setup" Then
    ls_sql = "Insert into GL_BSheet2 (Compcode,Bcode,BNCode,BNDesc) select '" & Gs_compcode & "' as Compcode,bcode,BnCode,BNdesc from GL_BSheet2 where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql

End If

If txtBS.Text = "Notes Items Setup" Then
    ls_sql = "Insert into GL_BSheet3 (Compcode,Bcode,BnCode,BnIcode,BniDesc) select '" & Gs_compcode & "' as Compcode,bcode,bncode,bnicode,bnidesc from GL_BSheet3 where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql
    
    ls_sql = "Insert into GL_BSheet3Detail (Compcode,Bcode,bncode,bnicode,accountno) select '" & Gs_compcode & "' as Compcode,Bcode,bncode,bnicode,accountno from GL_BSheet3Detail where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql

End If

If txtBS.Text <> "" Then
    Call MsgBox("Balance Sheet successfully Inserted", vbInformation)
End If

End Sub

Private Sub InsertPL()
If txtPL.Text = "All" Then
    ls_sql = "Insert into GL_PLSheet1 (Compcode,PLcode,PLDesc) select '" & Gs_compcode & "' as Compcode,PLcode,PLdesc from GL_PLSheet1 where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql

    ls_sql = "Insert into GL_PLSheet2 (Compcode,PLcode,PLNCode,PLNDesc) select '" & Gs_compcode & "' as Compcode,PLcode,PLnCode,PLNdesc from GL_PLSheet2 where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql

    ls_sql = "Insert into GL_PLSheet3 (Compcode,PLcode,PLnCode,PLnIcode,PLniDesc) select '" & Gs_compcode & "' as Compcode,PLcode,PLncode,PLnicode,PLnidesc from GL_PLSheet3 where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql
    
    ls_sql = "Insert into GL_PLSheet3Detail (Compcode,PLcode,PLncode,PLnicode,accountno) select '" & Gs_compcode & "' as Compcode,PLcode,PLncode,PLnicode,accountno from GL_PLSheet3Detail where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql

End If

If txtPL.Text = "Heads Setup" Then
    ls_sql = "Insert into GL_PLSheet1 (Compcode,PLcode,PLDesc) select '" & Gs_compcode & "' as Compcode,PLcode,PLdesc from GL_PLSheet1 where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql

End If

If txtPL.Text = "Notes Setup" Then
    ls_sql = "Insert into GL_PLSheet2 (Compcode,PLcode,PLNCode,PLNDesc) select '" & Gs_compcode & "' as Compcode,PLcode,PLnCode,PLNdesc from GL_PLSheet2 where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql

End If

If txtPL.Text = "Notes Items Setup" Then
    ls_sql = "Insert into GL_PLSheet3 (Compcode,PLcode,PLnCode,PLnIcode,PLniDesc) select '" & Gs_compcode & "' as Compcode,PLcode,PLncode,PLnicode,PLnidesc from GL_PLSheet3 where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql
    
    ls_sql = "Insert into GL_PLSheet3Detail (Compcode,PLcode,PLncode,PLnicode,accountno) select '" & Gs_compcode & "' as Compcode,PLcode,PLncode,PLnicode,accountno from GL_PLSheet3Detail where compcode = '" & txtfromCompcode & "'"
    gc_dbcon.Execute ls_sql
End If

If txtPL.Text <> "" Then
    Call MsgBox("Profit and Loss successfully Inserted", vbInformation)
End If

End Sub
Private Sub InsertVoucherType()
ls_sql = "Insert into GlVchrType(CompCode, BranchCode, VchrType, VchrDescrip, AccountNo, VchrFrequency, VchrBase, UserId, AddDate, AddTime) select '" & Gs_compcode & "' as CompCode, BranchCode, VchrType, VchrDescrip, AccountNo, VchrFrequency, VchrBase, UserId, AddDate, AddTime FROM GlVchrType  where compcode = '" & txtfromCompcode & "'"
gc_dbcon.Execute ls_sql
Call MsgBox("Voucher Types successfully Inserted", vbInformation)
End Sub
Private Sub Command2_Click()
Dim res
res = MsgBox("Are you sure to run all above process", vbYesNo + vbInformation)
If vbYes Then
chkcoa.Value = 1
ChkBS.Value = 1
ChkPL.Value = 1
ChkGLTran.Value = 1

txtlevel.Text = "All"
txtBS.Text = "All"
txtPL.Text = "All"

Call DeleteProcess
Call InsertCOA
Call InsertBS
Call InsertPL
Call InsertVoucherType

End If
End Sub

Private Sub Command3_Click()
If txtfromCompcode <> "" And checkvalidate = True Then
    If ChkVoucherType.Value = 1 Then
    Call InsertVoucherType
    End If
Else
   Call MsgBox("Select Valid Company Code From!!!", vbCritical)
   txtfromCompcode.SetFocus
End If


End Sub

Private Sub Command4_Click()
If txtfromCompcode <> "" And checkvalidate = True Then
    Call InsertCOA
Else
   Call MsgBox("Select Valid Company Code From!!!", vbCritical)
   txtfromCompcode.SetFocus
End If
End Sub

Private Sub Command5_Click()
If txtfromCompcode <> "" And checkvalidate = True Then
    Call InsertBS
Else
   Call MsgBox("Select Valid Company Code From!!!", vbCritical)
   txtfromCompcode.SetFocus
End If

End Sub

Private Sub Command8_Click()
If txtfromCompcode <> "" And checkvalidate = True Then
    Call InsertPL
Else
   Call MsgBox("Select Valid Company Code From!!!", vbCritical)
   txtfromCompcode.SetFocus
End If

End Sub

Private Sub Form_Load()
txtlevel.Text = "All"
txtBS.Text = "All"
txtPL.Text = "All"
Me.Caption = Gs_CompName
End Sub


Private Sub txtfromCompcode_Validate(Cancel As Boolean)
If txtfromCompcode <> "" Then
    txtfromCompcode = DoPad(txtfromCompcode, txtfromCompcode.MaxLength)
    
        pr_dumy.Open "Select compcode,compname from syscomp where compcode = '" & txtfromCompcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Company code not found", vbCritical)
                'Cancel = True
            Else
                txtCompdesc = pr_dumy("compname")
            End If
         pr_dumy.Close

End If
End Sub


Function checkvalidate() As Boolean
    
        pr_dumy.Open "Select compcode,compname from syscomp where compcode = '" & txtfromCompcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Company code not found", vbCritical)
                txtfromCompcode.SetFocus
                checkvalidate = False
            Else
                txtCompdesc = pr_dumy("compname")
                checkvalidate = True
            End If
         pr_dumy.Close

End Function


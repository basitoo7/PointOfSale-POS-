VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmARSaleType 
   Caption         =   "Sale Types"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmARSaleType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   2235
      Left            =   30
      TabIndex        =   9
      Top             =   570
      Width           =   6600
      Begin VB.TextBox txtusedcounter 
         Height          =   315
         Left            =   6000
         MaxLength       =   50
         TabIndex        =   22
         Tag             =   "SKIP"
         Top             =   180
         Width           =   510
      End
      Begin VB.CheckBox ChkCounterstatus 
         Caption         =   "Starting New Counter"
         Height          =   360
         Left            =   2535
         TabIndex        =   21
         Top             =   120
         Width           =   2385
      End
      Begin VB.TextBox txtcode 
         BackColor       =   &H00FFFF00&
         Height          =   330
         Left            =   1605
         MaxLength       =   3
         TabIndex        =   20
         Tag             =   "SKIPN"
         Top             =   150
         Width           =   540
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   2535
         MaxLength       =   50
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   135
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   1605
         MaxLength       =   50
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   540
         Width           =   4905
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
         Left            =   2190
         Picture         =   "frmARSaleType.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   150
         Width           =   315
      End
      Begin VB.Frame Frame1 
         ForeColor       =   &H00000080&
         Height          =   1365
         Left            =   45
         TabIndex        =   12
         Top             =   825
         Width           =   6510
         Begin VB.TextBox txtAccountDesc3 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3615
            TabIndex        =   19
            Top             =   915
            Width           =   2850
         End
         Begin VB.TextBox txtAccountDesc2 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3615
            TabIndex        =   18
            Top             =   555
            Width           =   2850
         End
         Begin VB.TextBox txtAccountDesc1 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3615
            TabIndex        =   17
            Top             =   195
            Width           =   2850
         End
         Begin VB.CommandButton CmdAccount2 
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
            Left            =   3270
            Picture         =   "frmARSaleType.frx":047C
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   555
            Width           =   315
         End
         Begin VB.CommandButton CmdAccount3 
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
            Left            =   3270
            Picture         =   "frmARSaleType.frx":05EE
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   930
            Width           =   315
         End
         Begin VB.CommandButton CmdAccount1 
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
            Left            =   3270
            Picture         =   "frmARSaleType.frx":0760
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   195
            Width           =   315
         End
         Begin VB.TextBox txtAccount3 
            Height          =   315
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   7
            Tag             =   "SKIP"
            Top             =   945
            Width           =   1695
         End
         Begin VB.TextBox txtAccount2 
            Height          =   315
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   5
            Tag             =   "SKIP"
            Top             =   570
            Width           =   1695
         End
         Begin VB.TextBox txtAccount1 
            Height          =   315
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   3
            Tag             =   "SKIP"
            Top             =   195
            Width           =   1695
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "SED A/c :"
            Height          =   225
            Left            =   135
            TabIndex        =   15
            Top             =   975
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Tax A/c :"
            Height          =   180
            Left            =   90
            TabIndex        =   14
            Top             =   570
            Width           =   1395
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Sale A/c :"
            Height          =   225
            Left            =   330
            TabIndex        =   13
            Top             =   210
            Width           =   1155
         End
      End
      Begin VB.Label Label6 
         Caption         =   "Used Counter:"
         Height          =   255
         Left            =   4965
         TabIndex        =   23
         Top             =   210
         Width           =   1050
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         Height          =   210
         Index           =   0
         Left            =   645
         TabIndex        =   11
         Top             =   555
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Code :"
         Height          =   210
         Left            =   1050
         TabIndex        =   10
         Top             =   180
         Width           =   495
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   1005
      ButtonWidth     =   1217
      ButtonHeight    =   953
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
            Caption         =   "&Find"
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
               Picture         =   "frmARSaleType.frx":08D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmARSaleType.frx":0D26
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmARSaleType.frx":117A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmARSaleType.frx":15CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmARSaleType.frx":1A22
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmARSaleType.frx":1E76
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmARSaleType.frx":25CA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmARSaleType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkSupp As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim OptVal As String
Dim PR_Dumy As New Recordset
Dim PR_SaleType As New Recordset


Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtcode
    Set PO_DESC = txtDesc
    Gs_SQL = "Select  StypeCode 'Sale Code' ,StypeDesc as Description from AccountSetupSale"
    Gs_FindFld = "StypeDesc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by StypeDesc"
    
    MyLookupOLDB.Caption = "Sale Types"
    MyLookupOLDB.Show 1
    
    If Len(txtcode) > 0 Then Call txtCode_KeyDown(vbKeyReturn, vbKeyShift)
    
End Sub
Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
If txtcode <> "" And KeyCode = vbKeyReturn Then
        txtcode = DoPad(txtcode, txtcode.MaxLength)
          ls_sql = "Select  StypeCode 'Sale Code' ,StypeDesc as Description from AccountSetupSale where compcode = '" & Gs_compcode & "' and stypecode = '" & txtcode & "'"
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Sale Type not found", vbCritical)
                Cancel = True
                txtcode.SetFocus
            Else
                txtDesc = PR_Dumy("description")
                txtDesc.SetFocus
                
                
            End If
         PR_Dumy.Close
         If Mode <> "A" Then Call SetVal

End If
End Sub
Private Sub CmdAccount1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccount1
    Set PO_DESC = txtAccountDesc1
    Gs_SQL = "Select  Accountno 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtaccount1) > 0 Then Call txtaccount1_KeyDown(vbKeyReturn, vbKeyShift)
    
End Sub
Private Sub CmdAccount2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtAccount2
    Set PO_DESC = txtAccountDesc2
    Gs_SQL = "Select  Accountno 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtAccount2) > 0 Then Call txtaccount2_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub CmdAccount3_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtAccount3
    Set PO_DESC = txtAccountDesc3
    Gs_SQL = "Select  Accountno 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtAccount3) > 0 Then Call txtaccount3_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub txtaccount1_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount1 <> "" And KeyCode = vbKeyReturn Then
          ls_sql = "Select Accountno 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  accountno = '" & txtaccount1 & "' "
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
                Cancel = True
                txtaccount1.SetFocus
            Else
                txtAccountDesc1 = PR_Dumy("description")
                txtAccount2.SetFocus
            End If
         PR_Dumy.Close
End If
End Sub
Private Sub SetVal()
Dim pr_dumyloadvalue As New Recordset
pr_dumyloadvalue.Open "Select * from AccountSetupSale where compcode = '" & Gs_compcode & "' and stypecode = '" & txtcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumyloadvalue.EOF Then

txtaccount1 = Trim(pr_dumyloadvalue("SaleAccount") & "")
If txtaccount1 <> "" Then Call txtaccount1_KeyDown(vbKeyReturn, vbKeyShift)
txtAccount2 = Trim(pr_dumyloadvalue("TaxAccount") & "")
If txtAccount2 <> "" Then Call txtaccount2_KeyDown(vbKeyReturn, vbKeyShift)
txtAccount3 = Trim(pr_dumyloadvalue("SedAccount") & "")
If txtAccount3 <> "" Then Call txtaccount3_KeyDown(vbKeyReturn, vbKeyShift)
ChkCounterstatus.Value = Val(0 & pr_dumyloadvalue("CounterStatus"))
txtusedcounter = Trim(pr_dumyloadvalue("usedCounter") & "")
End If
pr_dumyloadvalue.Close
End Sub


Private Sub txtaccount2_KeyDown(KeyCode As Integer, Shift As Integer)
If txtAccount2 <> "" And KeyCode = vbKeyReturn Then
          ls_sql = "Select Accountno 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  accountno = '" & txtAccount2 & "' "
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
                Cancel = True
                txtAccount2.SetFocus
            Else
                txtAccountDesc2 = PR_Dumy("description")
                txtAccount3.SetFocus
            End If
         PR_Dumy.Close
End If
End Sub
Private Sub txtaccount3_KeyDown(KeyCode As Integer, Shift As Integer)
If txtAccount3 <> "" And KeyCode = vbKeyReturn Then
          ls_sql = "Select Accountno 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  accountno = '" & txtAccount3 & "' "
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
                Cancel = True
                txtAccount3.SetFocus
            Else
                txtAccountDesc3 = PR_Dumy("description")
                 
            End If
         PR_Dumy.Close
End If
End Sub
Private Sub Form_Load()
  SetToolBar(1) = chkRights("SRSTS00001")
  SetToolBar(2) = chkRights("SRSTS00002")
  SetToolBar(3) = chkRights("SRSTS00003")
  SetToolBar(4) = chkRights("SRSTS00004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  
  PR_SaleType.Open "Select AccountSetupSale.* from AccountSetupSale order by stypecode ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
   
  
  PB_BlnkSupp = IIf(PR_SaleType.EOF, True, False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_SaleType.Close
  
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Range(Button.Index, 2, 3) Then
      cmdLookup.Enabled = True
    ElseIf Button.Index = 1 Then
      cmdLookup.Enabled = False
      txtcode = maxtranscode
    End If
    
    If PB_BlnkSupp And Range(Button.Index, 2, 3) Then
       Call SetErr("Data not found. ", vbCritical)
       Mode = ""
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_SaleType, Me, txtDesc, txtDesc, "X", "CompCount", 3, " Code", "Description", 1, False, Toolbar1)
    End If
    
End Sub

Public Sub SaveValues()
Dim ln_cnt As Integer
Dim ls_CodeID As String
PB_BlnkSupp = False
Dim ls_sql As String

     Select Case Mode
           Case "A"
               ls_sql = "INSERT into AccountSetupSale(Compcode,Stypecode,StypeDesc,SaleAccount,TaxAccount,SedAccount,CounterStatus,usedcounter) VALUES ('" & Gs_compcode & "','" & txtcode.Text & "','" & txtDesc.Text & "','" & txtaccount1.Text & "','" & txtAccount2.Text & "','" & txtAccount3.Text & "'," & ChkCounterstatus.Value & ",'" & txtusedcounter & "')"
               gc_dbcon.Execute ls_sql
           
           Case "E"
               ls_sql = "UPDATE AccountSetupSale SET Stypedesc= '" & txtDesc.Text & "', SaleAccount ='" & txtaccount1.Text & "', TaxAccount ='" & txtAccount2.Text & "' , SedAccount ='" & txtAccount3.Text & "' , CounterStatus = " & ChkCounterstatus.Value & ", UsedCounter = '" & txtusedcounter & "' WHERE compcode= '" & Gs_compcode & "' and  STypeCode= '" & txtcode.Text & "'"
               gc_dbcon.Execute ls_sql
              
           Case "D"
               ls_sql = "DELETE FROM AccountSetupSale WHERE compcode= '" & Gs_compcode & "' and  STypeCode= '" & txtcode.Text & "'"
               gc_dbcon.Execute ls_sql
           
     End Select
PR_SaleType.Requery

If Mode = "A" Then
txtcode = maxtranscode
End If
End Sub
Private Function maxtranscode() As String
PR_Dumy.Open "select max(stypeCode) as transcode from AccountSetupSale where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not PR_Dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & PR_Dumy("transcode")) + 1)), 3)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 3)
End If
PR_Dumy.Close
End Function
Public Function ChkInputs() As Boolean
    If Len(txtcode.Text) = txtcode.MaxLength And txtDesc <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub ClearVal()
End Sub

Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtaccount1.SetFocus
End Sub

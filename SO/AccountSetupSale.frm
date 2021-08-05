VERSION 5.00
Begin VB.Form frmAccountSetupSale 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Account Setup"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   Icon            =   "AccountSetupSale.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000080&
      Height          =   4185
      Left            =   15
      TabIndex        =   7
      Top             =   -60
      Width           =   4755
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   0
         MaxLength       =   50
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   0
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Reset"
         Height          =   345
         Left            =   30
         TabIndex        =   27
         Top             =   3690
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Cancel"
         Height          =   345
         Left            =   3360
         TabIndex        =   26
         Top             =   3690
         Width           =   1200
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Save"
         Height          =   345
         Left            =   2115
         TabIndex        =   25
         Top             =   3690
         Width           =   1200
      End
      Begin VB.Frame Frame2 
         Caption         =   "Customer A/C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   15
         TabIndex        =   22
         Top             =   2805
         Width           =   4710
         Begin VB.OptionButton Option2 
            Caption         =   "Base on Account Setup"
            Height          =   300
            Left            =   2130
            TabIndex        =   24
            Top             =   270
            Value           =   -1  'True
            Width           =   2040
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Base on Customer"
            Height          =   300
            Left            =   495
            TabIndex        =   23
            Top             =   255
            Width           =   1635
         End
      End
      Begin VB.TextBox txtaccount5 
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
         Left            =   2430
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1710
         Width           =   1920
      End
      Begin VB.TextBox txtaccount6 
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
         Left            =   2430
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2085
         Width           =   1920
      End
      Begin VB.CommandButton CmdAccount6 
         Height          =   315
         Left            =   4380
         Picture         =   "AccountSetupSale.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2070
         Width           =   315
      End
      Begin VB.CommandButton CmdAccount5 
         Height          =   315
         Left            =   4365
         Picture         =   "AccountSetupSale.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1695
         Width           =   315
      End
      Begin VB.CommandButton CmdAccount7 
         Height          =   315
         Left            =   4380
         Picture         =   "AccountSetupSale.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2460
         Width           =   315
      End
      Begin VB.TextBox txtaccount7 
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
         Left            =   2445
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2475
         Width           =   1920
      End
      Begin VB.TextBox txtaccount1 
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
         Left            =   2415
         MaxLength       =   50
         TabIndex        =   0
         Top             =   195
         Width           =   1920
      End
      Begin VB.TextBox txtaccount2 
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
         Left            =   2415
         MaxLength       =   50
         TabIndex        =   1
         Top             =   570
         Width           =   1920
      End
      Begin VB.TextBox txtaccount3 
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
         Left            =   2415
         MaxLength       =   50
         TabIndex        =   2
         Top             =   945
         Width           =   1920
      End
      Begin VB.CommandButton CmdAccount1 
         Height          =   315
         Left            =   4335
         Picture         =   "AccountSetupSale.frx":0760
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   195
         Width           =   315
      End
      Begin VB.CommandButton CmdAccount3 
         Height          =   315
         Left            =   4365
         Picture         =   "AccountSetupSale.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   930
         Width           =   315
      End
      Begin VB.CommandButton CmdAccount2 
         Height          =   315
         Left            =   4350
         Picture         =   "AccountSetupSale.frx":0A44
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   555
         Width           =   315
      End
      Begin VB.CommandButton CmdAccount4 
         Height          =   315
         Left            =   4365
         Picture         =   "AccountSetupSale.frx":0BB6
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1320
         Width           =   315
      End
      Begin VB.TextBox txtaccount4 
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
         Left            =   2430
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1335
         Width           =   1920
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer A/C :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   210
         TabIndex        =   21
         Top             =   1740
         Width           =   2145
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Misc A/c :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1035
         TabIndex        =   20
         Top             =   2115
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Other A/c :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   930
         TabIndex        =   19
         Top             =   2505
         Width           =   1440
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cost of Goods :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1170
         TabIndex        =   15
         Top             =   210
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Finished Goods(Inventory) :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   195
         TabIndex        =   14
         Top             =   600
         Width           =   2145
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Discount A/c :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1020
         TabIndex        =   13
         Top             =   975
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Sale Tax A/c :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   915
         TabIndex        =   12
         Top             =   1365
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmAccountSetupSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim PR_Dumy As New Recordset
Public PO_CODE As Object
Public PO_DESC As Object


Private Sub Command10_Click()
Call restvalues
End Sub

Private Sub restvalues()
Dim pr_dumyloadvalue As New Recordset
pr_dumyloadvalue.Open "Select * from AccountSetupSale", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumyloadvalue.EOF Then
txtaccount1 = Trim(pr_dumyloadvalue("CostofGoods") & "")
txtaccount2 = Trim(pr_dumyloadvalue("FinishedGoods") & "")
txtaccount3 = Trim(pr_dumyloadvalue("DiscountACT") & "")
txtaccount4 = Trim(pr_dumyloadvalue("SaleTacACT") & "")
txtaccount5 = Trim(pr_dumyloadvalue("CustomerACT") & "")
txtaccount6 = Trim(pr_dumyloadvalue("MiscACT") & "")
txtaccount7 = Trim(pr_dumyloadvalue("OtherACT") & "")
If Val(0 & pr_dumyloadvalue("OtherACT")) = 0 Then
    Option2.Value = True
Else
    Option1.Value = True
End If
End If
pr_dumyloadvalue.Close
End Sub

Private Sub Command8_Click()
Dim ls_sql As String
Dim ls_custoption As Integer
If Option1.Value = True Then ls_custoption = 0 Else ls_custoption = 1

ls_sql = "delete from AccountSetupSale where compcode = '" & Gs_compcode & "'"
gc_dbcon.Execute ls_sql

ls_sql = "insert into  AccountSetupSale (Compcode, CostofGoods, FinishedGoods, DiscountACT, SaleTacACT, CustomerACT, MiscACT, OtherACT, CustomerType)"
ls_sql = ls_sql & " values ('" & Gs_compcode & "','" & txtaccount1 & "','" & txtaccount2 & "','" & txtaccount3 & "','" & txtaccount4 & "','" & txtaccount5 & "','" & txtaccount6 & "','" & txtaccount7 & "'," & ls_custoption & ")"
gc_dbcon.Execute ls_sql
Call MsgBox("Successfully Updated !!!", vbInformation)
Call restvalues
End Sub

Private Sub Command9_Click()
txtaccount1 = ""
txtaccount2 = ""
txtaccount3 = ""
txtaccount4 = ""
txtaccount5 = ""
txtaccount6 = ""
txtaccount7 = ""
End Sub

Private Sub Form_Load()
Call restvalues
End Sub
Private Sub CmdAccount1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccount1
    Set PO_DESC = Text1
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
    Set PO_CODE = txtaccount2
    Set PO_DESC = Text1
    Gs_SQL = "Select  Accountno 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtaccount2) > 0 Then Call txtaccount2_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub CmdAccount3_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccount3
    Set PO_DESC = Text1
    Gs_SQL = "Select  Accountno 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtaccount3) > 0 Then Call txtaccount3_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub CmdAccount4_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccount4
    Set PO_DESC = Text1
    Gs_SQL = "Select  Accountno 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtaccount4) > 0 Then Call txtaccount4_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub CmdAccount5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccount5
    Set PO_DESC = Text1
    Gs_SQL = "Select  Accountno 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtaccount5) > 0 Then Call txtaccount5_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub CmdAccount6_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccount6
    Set PO_DESC = Text1
    Gs_SQL = "Select  Accountno 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtaccount6) > 0 Then Call txtaccount6_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub CmdAccount7_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccount7
    Set PO_DESC = Text1
    Gs_SQL = "Select  Accountno 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtaccount7) > 0 Then Call txtaccount7_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Function SearchAccount(ls_account As String) As Boolean
        ls_sql = "Select Accountno 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  accountno = '" & ls_account & "' "
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
                Cancel = True
                SearchAccount = False
            Else
                Text1 = PR_Dumy("description")
                SearchAccount = True
            End If
         PR_Dumy.Close
End Function

Private Sub txtaccount1_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount1 <> "" And KeyCode = vbKeyReturn Then
    If SearchAccount(txtaccount1) Then
    txtaccount2.SetFocus
    Else
    txtaccount1.SetFocus
    End If
End If
End Sub
Private Sub txtaccount2_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount2 <> "" And KeyCode = vbKeyReturn Then
    If SearchAccount(txtaccount2) Then
    txtaccount3.SetFocus
    Else
    txtaccount2.SetFocus
    End If
End If
End Sub
Private Sub txtaccount3_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount3 <> "" And KeyCode = vbKeyReturn Then
    If SearchAccount(txtaccount3) Then
    txtaccount4.SetFocus
    Else
    txtaccount3.SetFocus
    End If
End If

End Sub
Private Sub txtaccount4_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount4 <> "" And KeyCode = vbKeyReturn Then
    If SearchAccount(txtaccount4) Then
    txtaccount5.SetFocus
    Else
    txtaccount4.SetFocus
    End If
End If
End Sub
Private Sub txtaccount5_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount5 <> "" And KeyCode = vbKeyReturn Then
    If SearchAccount(txtaccount5) Then
    txtaccount6.SetFocus
    Else
    txtaccount5.SetFocus
    End If
End If
End Sub
Private Sub txtaccount6_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount6 <> "" And KeyCode = vbKeyReturn Then
    If SearchAccount(txtaccount6) Then
    txtaccount7.SetFocus
    Else
    txtaccount6.SetFocus
    End If
End If


End Sub
Private Sub txtaccount7_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount7 <> "" And KeyCode = vbKeyReturn Then
    If SearchAccount(txtaccount7) Then
    Option2.SetFocus
    Else
    txtaccount7.SetFocus
    End If
End If

End Sub




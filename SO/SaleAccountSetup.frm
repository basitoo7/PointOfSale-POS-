VERSION 5.00
Begin VB.Form frmSOAccountSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sale Accounts Setup"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6600
   Icon            =   "SaleAccountSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000080&
      Height          =   3495
      Left            =   15
      TabIndex        =   1
      Top             =   -60
      Width           =   6540
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3555
         MaxLength       =   50
         TabIndex        =   30
         Top             =   2550
         Width           =   2865
      End
      Begin VB.TextBox txtAccount7 
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
         Left            =   1830
         MaxLength       =   50
         TabIndex        =   29
         Top             =   2565
         Width           =   1365
      End
      Begin VB.CommandButton CMDAccount7 
         Height          =   315
         Left            =   3210
         Picture         =   "SaleAccountSetup.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2550
         Width           =   315
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3570
         MaxLength       =   50
         TabIndex        =   26
         Top             =   1440
         Width           =   2865
      End
      Begin VB.TextBox txtAccount4 
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
         Left            =   1830
         MaxLength       =   50
         TabIndex        =   25
         Top             =   1470
         Width           =   1365
      End
      Begin VB.CommandButton Cmdaccount4 
         Height          =   315
         Left            =   3225
         Picture         =   "SaleAccountSetup.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1440
         Width           =   315
      End
      Begin VB.CommandButton CMDAccount6 
         Height          =   315
         Left            =   3225
         Picture         =   "SaleAccountSetup.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2175
         Width           =   315
      End
      Begin VB.TextBox txtAccount6 
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
         Left            =   1830
         MaxLength       =   50
         TabIndex        =   21
         Top             =   2205
         Width           =   1365
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3570
         MaxLength       =   50
         TabIndex        =   20
         Top             =   2175
         Width           =   2865
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3570
         MaxLength       =   50
         TabIndex        =   18
         Top             =   1815
         Width           =   2865
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
         Left            =   1830
         MaxLength       =   50
         TabIndex        =   17
         Top             =   1845
         Width           =   1365
      End
      Begin VB.CommandButton CmdAccount5 
         Height          =   315
         Left            =   3225
         Picture         =   "SaleAccountSetup.frx":0760
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1815
         Width           =   315
      End
      Begin VB.CommandButton CmdAccount3 
         Height          =   315
         Left            =   3225
         Picture         =   "SaleAccountSetup.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1080
         Width           =   315
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
         Left            =   1830
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1110
         Width           =   1365
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3570
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1080
         Width           =   2865
      End
      Begin VB.CommandButton CmdAccount2 
         Height          =   315
         Left            =   3225
         Picture         =   "SaleAccountSetup.frx":0A44
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   720
         Width           =   315
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
         Left            =   1830
         MaxLength       =   50
         TabIndex        =   9
         Top             =   750
         Width           =   1365
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3570
         MaxLength       =   50
         TabIndex        =   8
         Top             =   720
         Width           =   2865
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3570
         MaxLength       =   50
         TabIndex        =   7
         Top             =   345
         Width           =   2865
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Reset"
         Height          =   345
         Left            =   90
         TabIndex        =   6
         Top             =   3030
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Cancel"
         Height          =   345
         Left            =   5505
         TabIndex        =   5
         Top             =   3030
         Width           =   930
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Save"
         Height          =   345
         Left            =   4545
         TabIndex        =   4
         Top             =   3045
         Width           =   930
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
         Left            =   1830
         MaxLength       =   50
         TabIndex        =   0
         Top             =   375
         Width           =   1365
      End
      Begin VB.CommandButton CmdAccount1 
         Height          =   315
         Left            =   3225
         Picture         =   "SaleAccountSetup.frx":0BB6
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   345
         Width           =   315
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Cost of Sale :"
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
         Left            =   525
         TabIndex        =   31
         Top             =   2595
         Width           =   1260
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Cash Account :"
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
         Left            =   630
         TabIndex        =   27
         Top             =   1500
         Width           =   1155
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Other Account2 :"
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
         Left            =   540
         TabIndex        =   23
         Top             =   2235
         Width           =   1260
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit Disc Account :"
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
         Left            =   210
         TabIndex        =   19
         Top             =   1875
         Width           =   1590
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Sed Account :"
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
         Left            =   630
         TabIndex        =   15
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "GST Account :"
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
         Left            =   630
         TabIndex        =   11
         Top             =   780
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Disc Account :"
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
         Left            =   60
         TabIndex        =   3
         Top             =   405
         Width           =   1725
      End
   End
End
Attribute VB_Name = "frmSOAccountSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PR_dumy As New Recordset
Public PO_CODE As Object
Public PO_DESC As Object
Public ls_transtype As String



Private Sub Cmdaccount4_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtAccount4
    Set PO_DESC = Text8
    Gs_SQL = "Select  AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtAccount4) > 0 Then Call txtaccount4_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub CmdAccount5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccount5
    Set PO_DESC = Text5
    Gs_SQL = "Select  AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtaccount5) > 0 Then Call txtaccount5_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub CMDAccount6_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtAccount6
    Set PO_DESC = Text3
    Gs_SQL = "Select  AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtAccount6) > 0 Then Call txtaccount6_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub CMDAccount7_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtAccount7
    Set PO_DESC = Text7
    Gs_SQL = "Select  AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtAccount7) > 0 Then Call txtaccount7_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command10_Click()
Call restvalues
End Sub

Private Sub restvalues()
Dim pr_dumyloadvalue As New Recordset
pr_dumyloadvalue.Open "Select * from SO_SaleAccountSetup where compcode = '" & Gs_compcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumyloadvalue.EOF Then
txtaccount1 = Trim(pr_dumyloadvalue("CashdiscAccount") & "")
If txtaccount1 <> "" Then Call txtaccount1_KeyDown(vbKeyReturn, vbKeyShift)
txtaccount2 = Trim(pr_dumyloadvalue("GSTAccount") & "")
If txtaccount2 <> "" Then Call txtaccount2_KeyDown(vbKeyReturn, vbKeyShift)
txtaccount3 = Trim(pr_dumyloadvalue("SEdAccount") & "")
If txtaccount3 <> "" Then Call txtaccount3_KeyDown(vbKeyReturn, vbKeyShift)
txtAccount4 = Trim(pr_dumyloadvalue("CashAccount") & "")
If txtAccount4 <> "" Then Call txtaccount4_KeyDown(vbKeyReturn, vbKeyShift)
txtaccount5 = Trim(pr_dumyloadvalue("CreditDiscAccount") & "")
If txtaccount5 <> "" Then Call txtaccount5_KeyDown(vbKeyReturn, vbKeyShift)
txtAccount6 = Trim(pr_dumyloadvalue("OtherAccount2") & "")
If txtAccount6 <> "" Then Call txtaccount6_KeyDown(vbKeyReturn, vbKeyShift)
txtAccount7 = Trim(pr_dumyloadvalue("CostofSale") & "")
If txtAccount7 <> "" Then Call txtaccount7_KeyDown(vbKeyReturn, vbKeyShift)

End If
pr_dumyloadvalue.Close
End Sub

Private Sub Command8_Click()
Dim ls_sql As String


ls_sql = "delete from SO_SaleAccountSetup where compcode = '" & Gs_compcode & "' "
gc_dbcon.Execute ls_sql

ls_sql = "insert into  SO_SaleAccountSetup (Compcode, CashdiscAccount,GSTAccount,SEDAccount,CashAccount,CreditDiscAccount,OtherAccount2,CostofSale)"
ls_sql = ls_sql & " values ('" & Gs_compcode & "','" & txtaccount1 & "','" & txtaccount2 & "','" & txtaccount3 & "','" & txtAccount4 & "','" & txtaccount5 & "','" & txtAccount6 & "','" & txtAccount7 & "')"
gc_dbcon.Execute ls_sql

Call MsgBox("Successfully Updated !!!", vbInformation)
Call restvalues

End Sub

Private Sub Command9_Click()
txtaccount1 = ""
Text1.Text = ""

txtaccount2 = ""
Text2.Text = ""

txtaccount3 = ""
Text4.Text = ""

txtAccount4 = ""
Text8.Text = ""

txtaccount5 = ""
Text5.Text = ""

txtAccount6 = ""
Text3.Text = ""

txtAccount7 = ""
Text7.Text = ""

End Sub

Private Sub Form_Load()
Call restvalues
End Sub
Private Sub CmdAccount1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccount1
    Set PO_DESC = Text1
    Gs_SQL = "Select  AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtaccount1) > 0 Then Call txtaccount1_KeyDown(vbKeyReturn, vbKeyShift)
    
End Sub

Private Sub txtaccount1_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount1 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  AccountNo = '" & txtaccount1 & "' "
          PR_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text1 = PR_dumy("description")
            End If
         PR_dumy.Close

End If

End Sub
Private Sub CmdAccount2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccount2
    Set PO_DESC = Text2
    Gs_SQL = "Select  AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtaccount2) > 0 Then Call txtaccount2_KeyDown(vbKeyReturn, vbKeyShift)
    
End Sub

Private Sub txtaccount2_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount2 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  AccountNo = '" & txtaccount2 & "' "
          PR_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text2 = PR_dumy("description")
                
            End If
         PR_dumy.Close

End If

End Sub

Private Sub CmdAccount3_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccount3
    Set PO_DESC = Text4
    Gs_SQL = "Select  AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtaccount3) > 0 Then Call txtaccount3_KeyDown(vbKeyReturn, vbKeyShift)
    
End Sub

Private Sub txtaccount3_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount3 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  AccountNo = '" & txtaccount3 & "' "
          PR_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text4 = PR_dumy("description")
            End If
         PR_dumy.Close

End If

End Sub
Private Sub txtaccount4_KeyDown(KeyCode As Integer, Shift As Integer)
If txtAccount4 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  AccountNo = '" & txtAccount4 & "' "
          PR_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text8 = PR_dumy("description")
                
            End If
         PR_dumy.Close

End If

End Sub


Private Sub txtaccount5_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount5 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  AccountNo = '" & txtaccount5 & "' "
          PR_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text5 = PR_dumy("description")
                
            End If
         PR_dumy.Close

End If
End Sub

Private Sub txtaccount6_KeyDown(KeyCode As Integer, Shift As Integer)
If txtAccount6 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  AccountNo = '" & txtAccount6 & "' "
          PR_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text3 = PR_dumy("description")
                
            End If
         PR_dumy.Close

End If

End Sub


Private Sub txtaccount7_KeyDown(KeyCode As Integer, Shift As Integer)
If txtAccount7 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select AccountNo 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  AccountNo = '" & txtAccount7 & "' "
          PR_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text7 = PR_dumy("description")
                
            End If
         PR_dumy.Close

End If

End Sub

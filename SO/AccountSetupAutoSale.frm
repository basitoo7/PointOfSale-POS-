VERSION 5.00
Begin VB.Form frmAccountSetupAutoSale 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Account Setup"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   Icon            =   "AccountSetupAutoSale.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000080&
      Height          =   2295
      Left            =   15
      TabIndex        =   4
      Top             =   -60
      Width           =   4755
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   0
         MaxLength       =   50
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   0
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Reset"
         Height          =   345
         Left            =   30
         TabIndex        =   15
         Top             =   1860
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Cancel"
         Height          =   345
         Left            =   3360
         TabIndex        =   14
         Top             =   1860
         Width           =   1200
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Save"
         Height          =   345
         Left            =   2115
         TabIndex        =   13
         Top             =   1860
         Width           =   1200
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
         Picture         =   "AccountSetupAutoSale.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   195
         Width           =   315
      End
      Begin VB.CommandButton CmdAccount3 
         Height          =   315
         Left            =   4335
         Picture         =   "AccountSetupAutoSale.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   930
         Width           =   315
      End
      Begin VB.CommandButton CmdAccount2 
         Height          =   315
         Left            =   4335
         Picture         =   "AccountSetupAutoSale.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   555
         Width           =   315
      End
      Begin VB.CommandButton CmdAccount4 
         Height          =   315
         Left            =   4335
         Picture         =   "AccountSetupAutoSale.frx":0760
         Style           =   1  'Graphical
         TabIndex        =   5
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
      Begin VB.Label Label3 
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
         Left            =   1170
         TabIndex        =   12
         Top             =   210
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Sale Account :"
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
         TabIndex        =   11
         Top             =   600
         Width           =   2145
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Tax Account :"
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
         TabIndex        =   10
         Top             =   975
         Width           =   1335
      End
      Begin VB.Label Label6 
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
         Left            =   915
         TabIndex        =   9
         Top             =   1365
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmAccountSetupAutoSale"
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
pr_dumyloadvalue.Open "Select * from AccountSetupSale where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumyloadvalue.EOF Then
txtaccount1 = Trim(pr_dumyloadvalue("CashAccount") & "")
txtaccount2 = Trim(pr_dumyloadvalue("SaleAccount") & "")
txtaccount3 = Trim(pr_dumyloadvalue("TaxAccount") & "")
txtaccount4 = Trim(pr_dumyloadvalue("SedAccount") & "")
End If
pr_dumyloadvalue.Close
End Sub

Private Sub Command8_Click()
Dim ls_sql As String


ls_sql = "delete from AccountSetupSale where compcode = '" & Gs_compcode & "'"
gc_dbcon.Execute ls_sql

ls_sql = "insert into  AccountSetupSale (Compcode, CashAccount, SaleAccount, TaxAccount, SedAccount)"
ls_sql = ls_sql & " values ('" & Gs_compcode & "','" & txtaccount1 & "','" & txtaccount2 & "','" & txtaccount3 & "','" & txtaccount4 & "')"
gc_dbcon.Execute ls_sql

Call MsgBox("Successfully Updated !!!", vbInformation)
Call restvalues

End Sub

Private Sub Command9_Click()
txtaccount1 = ""
txtaccount2 = ""
txtaccount3 = ""
txtaccount4 = ""
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
    Else
    txtaccount4.SetFocus
    End If
End If
End Sub




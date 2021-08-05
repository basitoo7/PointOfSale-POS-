VERSION 5.00
Begin VB.Form frmVendorAccountSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendor Account Setup"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6045
   Icon            =   "VendorAccountSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000080&
      Height          =   1515
      Left            =   15
      TabIndex        =   1
      Top             =   -60
      Width           =   5985
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3180
         MaxLength       =   50
         TabIndex        =   7
         Top             =   345
         Width           =   2685
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Reset"
         Height          =   345
         Left            =   90
         TabIndex        =   6
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Cancel"
         Height          =   345
         Left            =   4980
         TabIndex        =   5
         Top             =   1080
         Width           =   930
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Save"
         Height          =   345
         Left            =   4020
         TabIndex        =   4
         Top             =   1095
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
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   0
         Top             =   375
         Width           =   1365
      End
      Begin VB.CommandButton CmdAccount1 
         Height          =   315
         Left            =   2835
         Picture         =   "VendorAccountSetup.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   345
         Width           =   315
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Vendor Account :"
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
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmVendorAccountSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pr_dumy As New Recordset
Public PO_CODE As Object
Public PO_DESC As Object


Private Sub Command10_Click()
Call restvalues
End Sub

Private Sub restvalues()
Dim pr_dumyloadvalue As New Recordset
pr_dumyloadvalue.Open "Select * from VendorAccountSetup where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumyloadvalue.EOF Then
txtaccount1 = Trim(pr_dumyloadvalue("VendorAccount") & "")
If txtaccount1 <> "" Then Call txtaccount1_KeyDown(vbKeyReturn, vbKeyShift)
End If
pr_dumyloadvalue.Close
End Sub

Private Sub Command8_Click()
Dim ls_sql As String


ls_sql = "delete from vendorAccountSetup where compcode = '" & Gs_compcode & "'"
gc_dbcon.Execute ls_sql

ls_sql = "insert into  VendorAccountSetup (Compcode, VendorAccount)"
ls_sql = ls_sql & " values ('" & Gs_compcode & "','" & txtaccount1 & "')"
gc_dbcon.Execute ls_sql

Call MsgBox("Successfully Updated !!!", vbInformation)
Call restvalues

End Sub

Private Sub Command9_Click()
txtaccount1 = ""
Text1.Text = ""
End Sub

Private Sub Form_Load()
Call restvalues
End Sub
Private Sub CmdAccount1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccount1
    Set PO_DESC = Text1
    Gs_SQL = "Select  Acct_sub1+Acct_sub2 'Account Code' ,Acct_Desc as Description from Gl_sub2"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtaccount1) > 0 Then Call txtaccount1_KeyDown(vbKeyReturn, vbKeyShift)
    
End Sub
Function SearchAccount(ls_account As String) As Boolean
        ls_sql = "Select Accountno 'Account Code' ,Acct_Desc as Description from Gl_Detail where compcode = '" & Gs_compcode & "' and  accountno = '" & ls_account & "' "
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
                'Cancel = True
                SearchAccount = False
            Else
                Text1 = pr_dumy("description")
                SearchAccount = True
            End If
         pr_dumy.Close
End Function

Private Sub txtaccount1_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount1 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select Acct_sub1+Acct_sub2 'Account Code' ,Acct_Desc as Description from Gl_sub2 where compcode = '" & Gs_compcode & "' and  Acct_sub1+Acct_sub2 = '" & txtaccount1 & "' "
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text1 = pr_dumy("description")
            End If
         pr_dumy.Close

End If

End Sub





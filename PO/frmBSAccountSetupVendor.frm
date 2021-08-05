VERSION 5.00
Begin VB.Form frmBSAccountSetupVendor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clients Balance Sheet Setup"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6045
   Icon            =   "frmBSAccountSetupVendor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000080&
      Height          =   1500
      Left            =   15
      TabIndex        =   1
      Top             =   -60
      Width           =   5985
      Begin VB.CommandButton CmdAccount2 
         Height          =   315
         Left            =   2730
         Picture         =   "frmBSAccountSetupVendor.frx":030A
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
         Left            =   1335
         MaxLength       =   50
         TabIndex        =   9
         Top             =   750
         Width           =   1365
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3075
         MaxLength       =   50
         TabIndex        =   8
         Top             =   720
         Width           =   2865
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3075
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
         Top             =   1065
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Cancel"
         Height          =   345
         Left            =   4980
         TabIndex        =   5
         Top             =   1065
         Width           =   930
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Save"
         Height          =   345
         Left            =   4020
         TabIndex        =   4
         Top             =   1065
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
         Left            =   1335
         MaxLength       =   50
         TabIndex        =   0
         Top             =   375
         Width           =   1365
      End
      Begin VB.CommandButton CmdAccount1 
         Height          =   315
         Left            =   2730
         Picture         =   "frmBSAccountSetupVendor.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   345
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Notes :"
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
         Left            =   135
         TabIndex        =   11
         Top             =   765
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Main Head :"
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
         Left            =   135
         TabIndex        =   3
         Top             =   375
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmBSAccountSetupVendor"
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
pr_dumyloadvalue.Open "Select * from VendorBSSetup where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumyloadvalue.EOF Then
txtaccount1 = Trim(pr_dumyloadvalue("Bsmainhead") & "")
If txtaccount1 <> "" Then Call txtaccount1_KeyDown(vbKeyReturn, vbKeyShift)
txtaccount2 = Trim(pr_dumyloadvalue("bsnotes") & "")
If txtaccount2 <> "" Then Call txtaccount2_KeyDown(vbKeyReturn, vbKeyShift)


End If
pr_dumyloadvalue.Close
End Sub

Private Sub Command8_Click()
Dim ls_sql As String


ls_sql = "delete from VendorBSSetup where compcode = '" & Gs_compcode & "'"
gc_dbcon.Execute ls_sql

ls_sql = "insert into  VendorBSSetup (Compcode, BsMainHead,BSNotes)"
ls_sql = ls_sql & " values ('" & Gs_compcode & "','" & txtaccount1 & "','" & txtaccount2 & "')"
gc_dbcon.Execute ls_sql

Call MsgBox("Successfully Updated !!!", vbInformation)
Call restvalues

End Sub

Private Sub Command9_Click()
txtaccount1 = ""
Text1.Text = ""

txtaccount2 = ""
Text2.Text = ""

End Sub

Private Sub Form_Load()
Call restvalues
End Sub
Private Sub CmdAccount1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccount1
    Set PO_DESC = Text1
    Gs_SQL = "Select  Bcode 'Account Code', BDesc as Description from GL_BSheet1"
    Gs_FindFld = "BDesc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by BDesc"
    
    MyLookupOLDB.Caption = "Balance Sheet Main Heads"
    MyLookupOLDB.Show 1
    
    If Len(txtaccount1) > 0 Then Call txtaccount1_KeyDown(vbKeyReturn, vbKeyShift)
    
End Sub

Private Sub txtaccount1_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount1 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select BCode 'Account Code' ,BDesc as Description from GL_BSheet1 where compcode = '" & Gs_compcode & "' and  bcode = '" & txtaccount1 & "'"
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text1 = PR_Dumy("description")
            End If
         PR_Dumy.Close

End If

End Sub
Private Sub CmdAccount2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccount2
    Set PO_DESC = Text2
    Gs_SQL = "Select  Bncode 'Account Code' ,BnDesc as Description from Gl_BSheet2"
    Gs_FindFld = "BNDesc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' and Bcode = '" & txtaccount1 & "'"
    Gs_OrderBy = "Order by BNdesc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtaccount2) > 0 Then Call txtaccount2_KeyDown(vbKeyReturn, vbKeyShift)
    
End Sub

Private Sub txtaccount2_KeyDown(KeyCode As Integer, Shift As Integer)
If txtaccount2 <> "" And KeyCode = vbKeyReturn Then
         ls_sql = "Select BNCode 'Account Code' ,BNDesc as Description from Gl_BSheet2 where compcode = '" & Gs_compcode & "' and  bncode = '" & txtaccount2 & "' and  bcode = '" & txtaccount1 & "' "
          PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If PR_Dumy.EOF Then
                Call MsgBox("Account code not found", vbCritical)
            Else
                Text2 = PR_Dumy("description")
                
            End If
         PR_Dumy.Close

End If

End Sub






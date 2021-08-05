VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmExpenseType 
   Caption         =   "Expense Type"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8430
   Icon            =   "frmExpenseType.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1425
      Left            =   45
      TabIndex        =   1
      Top             =   570
      Width           =   8355
      Begin VB.TextBox txtsub0 
         Height          =   315
         Left            =   2115
         MaxLength       =   13
         TabIndex        =   9
         Top             =   975
         Width           =   1515
      End
      Begin VB.TextBox txtSubDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3990
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   960
         Width           =   4290
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   3660
         Picture         =   "frmExpenseType.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   960
         Width           =   315
      End
      Begin VB.TextBox txtMTCode 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   2115
         MaxLength       =   3
         TabIndex        =   6
         Tag             =   "SKIPN"
         Top             =   240
         Width           =   525
      End
      Begin VB.TextBox txtDesc 
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
         Left            =   2115
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   600
         Width           =   6135
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   315
         Left            =   2640
         Picture         =   "frmExpenseType.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "GL Account# :"
         Height          =   195
         Left            =   990
         TabIndex        =   10
         Top             =   1005
         Width           =   1050
      End
      Begin VB.Label Label2 
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
         Left            =   1155
         TabIndex        =   4
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Expense Code :"
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
         Left            =   915
         TabIndex        =   2
         Top             =   240
         Width           =   1140
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8430
      _ExtentX        =   14870
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
               Picture         =   "frmExpenseType.frx":05EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExpenseType.frx":0A42
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExpenseType.frx":0E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExpenseType.frx":12EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExpenseType.frx":173E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExpenseType.frx":1B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExpenseType.frx":22E6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmExpenseType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PO_CODE As Object
Public PO_DESC As Object
Dim Mode As String
Dim lb_found As Boolean
Dim PR_ItemClass As New Recordset
Dim PR_GlSub2 As New Recordset

Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtmtcode
    Set PO_DESC = txtDesc
    Gs_SQL = "Select ECode, Description from GL_ExenseType"
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Expense Type"
    MyLookupOLDB.Show 1
    If txtmtcode <> "" Then Call txtmtCode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub


Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtsub0.SetFocus
End If
End Sub

Private Sub txtDesc_LostFocus()
If txtDesc <> "" Then
txtDesc = UCase(txtDesc)
End If
End Sub

Private Sub txtsub0_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Val(txtsub0.Text) <> 0 Then
      PR_GlSub2.Open "Select * from Gl_Detail where Accountno = '" & txtsub0.Text & "' and CompCode = '" & Gs_compcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If PR_GlSub2.EOF Then
             Call SetErr("GL Code Not Found !!!", vbCritical)
             If txtsub0.Enabled Then txtsub0.SetFocus
             txtSubDesc.Text = ""
         Else
             txtSubDesc.Text = PR_GlSub2("Acct_Desc")
             
         End If
 PR_GlSub2.Close
End If
End Sub

Private Sub Command4_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtsub0
    Set PO_DESC = txtSubDesc
    Gs_SQL = "Select accountno  'Account No', Acct_Desc  'Description' from  gl_detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
    Gs_OrderBy = "Order by Acct_Desc"
    MyLookupOLDB.Caption = "GL Accounts."
    MyLookupOLDB.Show 1
    
   If Len(txtsub0) > 0 Then txtsub0_KeyDown vbKeyReturn, vbKeyShift

End Sub



Private Sub Form_Load()
  
  SetToolBar(1) = chkRights("SRITC00001")
  SetToolBar(2) = chkRights("SRITC00002")
  SetToolBar(3) = chkRights("SRITC00003")
  SetToolBar(4) = chkRights("SRITC00004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
   
End Sub
Private Function maxtranscode() As String
Dim PR_Dumy As New Recordset
PR_Dumy.Open "select max(ECode) as transcode from GL_ExenseType where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not PR_Dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & PR_Dumy("transcode")) + 1)), 3)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 3)
End If
PR_Dumy.Close
End Function

Private Sub txtmtCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtmtcode <> "" Then
        txtmtcode.Text = DoPad(UCase(txtmtcode.Text), txtmtcode.MaxLength)
        PR_ItemClass.Open "Select * from  GL_ExenseType where compcode = '" & Gs_compcode & "' and ECode = '" & txtmtcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        Select Case Mode
            Case "A"
                If Not PR_ItemClass.EOF Then
                   Call MsgBox(Gs_RecFdMsg, vbCritical)
                  Cancel = True
                   Call ClearVal
                   txtmtcode.SetFocus
                Else
                   txtDesc.SetFocus
                End If
            Case Else
                If PR_ItemClass.EOF Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   Call ClearVal
                   txtmtcode.SetFocus
                Else
                   Call SetVal
                   txtDesc.SetFocus
                  
                End If
            End Select
  PR_ItemClass.Close
ElseIf KeyCode = vbKeyReturn And txtmtcode = "" Then
    txtmtcode = ""
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
      cmdLookup.Enabled = False
      txtmtcode.Locked = True
    Else
      txtmtcode.Locked = False
      cmdLookup.Enabled = True
    End If
    If PB_BlnkItmClass And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, Me, Me, txtmtcode, txtDesc, "X", "CompCount", 3, "ItemClass", "Description", 1, False, Toolbar1)
    End If
    If Mode = "A" Then
    txtmtcode = maxtranscode
    txtDesc.SetFocus
    End If
   
End Sub

Public Sub SaveValues()
'On Error GoTo LocalErr
Dim ls_sql As String



gc_dbcon.BeginTrans
     Select Case Mode
           Case "A"
              ls_sql = "INSERT into GL_ExenseType(compcode,ECode,Description,Glcode) VALUES ('" & Gs_compcode & "','" & txtmtcode.Text & "','" & RepApp(txtDesc.Text) & "','" & txtsub0.Text & "')"
              gc_dbcon.Execute ls_sql
             
           Case "E"
              ls_sql = "UPDATE GL_ExenseType SET Description= '" & RepApp(txtDesc.Text) & "',glcode = '" & txtsub0 & "' WHERE  compcode = '" & Gs_compcode & "' and ECode= '" & txtmtcode.Text & "'"
              gc_dbcon.Execute ls_sql
           Case "D"
              ls_sql = "DELETE FROM GL_ExenseType WHERE ECode = '" & txtmtcode.Text & "' and compcode = '" & Gs_compcode & "'"
              gc_dbcon.Execute ls_sql
           
     End Select
gc_dbcon.CommitTrans


If Mode = "A" Then
    txtmtcode = maxtranscode
    End If
Exit Sub

LocalErr:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
End Sub
Public Sub ClearVal()
     txtmtcode = ""
     txtDesc = ""
     
End Sub

Private Sub SetVal()
     txtDesc = PR_ItemClass("Description")
     txtsub0 = Trim(PR_ItemClass("Glcode") & "")
     If txtsub0 <> "" Then Call txtsub0_KeyDown(vbKeyReturn, vbKeyShift)
     
End Sub
Public Function ChkInputs() As Boolean
    If Trim(txtmtcode.Text) = "" Then
       Call MsgBox("Enter Expense Code !!!", vbCritical)
        txtmtcode.SetFocus
       ChkInputs = False
    ElseIf Trim(txtDesc.Text) = "" Then
       Call MsgBox("Enter Expense Description !!!", vbCritical)
        txtDesc.SetFocus
        ChkInputs = False
    Else
        ChkInputs = True
    End If
End Function

Public Sub SetFrmEnv(ls_mode As String)
    txtDesc.Enabled = IIf(ls_mode <> "D", True, False)
End Sub


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemCategory 
   Caption         =   "Departments Setup"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   Icon            =   "frmItemCategory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   45
      TabIndex        =   1
      Top             =   570
      Width           =   7845
      Begin VB.TextBox txtgprofit 
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
         Left            =   3870
         MaxLength       =   50
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   195
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtConsumptionAct 
         Height          =   315
         Left            =   2100
         MaxLength       =   13
         TabIndex        =   17
         Top             =   1380
         Width           =   1515
      End
      Begin VB.TextBox txtConsumptionDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3990
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1365
         Width           =   4290
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   3645
         Picture         =   "frmItemCategory.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1365
         Width           =   315
      End
      Begin VB.TextBox txtSaleAct 
         Height          =   315
         Left            =   2100
         MaxLength       =   13
         TabIndex        =   14
         Top             =   1785
         Width           =   1515
      End
      Begin VB.TextBox txtSaleActDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3990
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1770
         Width           =   4305
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   3645
         Picture         =   "frmItemCategory.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1770
         Width           =   315
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   3645
         Picture         =   "frmItemCategory.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   975
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtPurchaseactdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3975
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   975
         Visible         =   0   'False
         Width           =   4290
      End
      Begin VB.TextBox txtPurchaseACT 
         Height          =   315
         Left            =   2100
         MaxLength       =   13
         TabIndex        =   9
         Top             =   990
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox txtempdisc 
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
         Left            =   7485
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtItemCode 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   1515
         MaxLength       =   3
         TabIndex        =   6
         Tag             =   "SKIPN"
         Top             =   225
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
         Left            =   1515
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   585
         Width           =   6240
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   315
         Left            =   2040
         Picture         =   "frmItemCategory.frx":0760
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   225
         Width           =   315
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "G-Profit"
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
         Left            =   3270
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Left            =   4560
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Left            =   8160
         TabIndex        =   21
         Top             =   165
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "GL Purchase Account# :"
         Height          =   195
         Left            =   255
         TabIndex        =   20
         Top             =   1020
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "GL Consumption Account# :"
         Height          =   195
         Left            =   45
         TabIndex        =   19
         Top             =   1425
         Width           =   2010
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "GLSale Account# :"
         Height          =   195
         Left            =   645
         TabIndex        =   18
         Top             =   1845
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Emp Disc:"
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
         Left            =   6750
         TabIndex        =   8
         Top             =   165
         Visible         =   0   'False
         Width           =   705
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
         Left            =   525
         TabIndex        =   4
         Top             =   615
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Department Code :"
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
         Left            =   105
         TabIndex        =   2
         Top             =   225
         Width           =   1335
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   1058
      ButtonWidth     =   1402
      ButtonHeight    =   1005
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
               Picture         =   "frmItemCategory.frx":08D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemCategory.frx":0D26
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemCategory.frx":117A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemCategory.frx":15CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemCategory.frx":1A22
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemCategory.frx":1E76
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemCategory.frx":25CA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmItemCategory"
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
Dim pr_dumy As New Recordset
Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtitemcode
    Set PO_DESC = txtDesc
    Gs_SQL = "Select CatCode, Description from IC_ItemCategory "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Item Class"
    MyLookupOLDB.Show 1
    If txtitemcode <> "" Then Call txtItemcode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtSaleAct
    Set PO_DESC = txtSaleActDesc
    Gs_SQL = "Select  Accountno 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtSaleAct) > 0 Then Call txtSaleAct_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub txtempdisc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtempDisc.SetFocus
End Sub

Private Sub txtSaleAct_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Val(txtSaleAct.Text) <> 0 Then
      PR_GlSub2.Open "Select * from  Gl_Detail where Accountno = '" & txtSaleAct & "' and CompCode = '" & Gs_compcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If PR_GlSub2.EOF Then
             Call SetErr("GL Code Not Found !!!", vbCritical)
             If txtSaleAct.Enabled Then txtSaleAct.SetFocus
             txtSaleActDesc.Text = ""
         Else
             txtSaleActDesc.Text = PR_GlSub2("Acct_Desc")
             If txtSaleAct.Enabled Then txtSaleAct.SetFocus
         End If
 PR_GlSub2.Close
End If
End Sub


Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtPurchaseACT
    Set PO_DESC = txtPurchaseactdesc
    Gs_SQL = "Select  Accountno 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtPurchaseACT) > 0 Then Call txtPurchaseACT_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command4_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtConsumptionAct
    Set PO_DESC = txtConsumptionDesc
    Gs_SQL = "Select  Accountno 'Account Code' ,Acct_Desc as Description from Gl_Detail"
    Gs_FindFld = "Acct_Desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
    Gs_OrderBy = "Order by Acct_desc"
    
    MyLookupOLDB.Caption = "Accounts"
    MyLookupOLDB.Show 1
    
    If Len(txtConsumptionAct) > 0 Then Call txtConsumptionAct_KeyDown(vbKeyReturn, vbKeyShift)

End Sub
Private Sub txtConsumptionAct_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Val(txtConsumptionAct.Text) <> 0 Then
      PR_GlSub2.Open "Select * from  Gl_Detail where Accountno = '" & txtConsumptionAct & "' and CompCode = '" & Gs_compcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If PR_GlSub2.EOF Then
             Call SetErr("GL Code Not Found !!!", vbCritical)
             If txtConsumptionAct.Enabled Then txtConsumptionAct.SetFocus
             txtConsumptionDesc.Text = ""
         Else
             txtConsumptionDesc.Text = PR_GlSub2("Acct_Desc")
             If txtSaleAct.Enabled Then txtSaleAct.SetFocus
         End If
 PR_GlSub2.Close
End If
End Sub


Private Sub txtPurchaseACT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Val(txtPurchaseACT.Text) <> 0 Then
      PR_GlSub2.Open "Select * from  Gl_Detail where Accountno = '" & txtPurchaseACT & "' and CompCode = '" & Gs_compcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If PR_GlSub2.EOF Then
             Call SetErr("GL Code Not Found !!!", vbCritical)
             If txtPurchaseACT.Enabled Then txtPurchaseACT.SetFocus
             txtPurchaseactdesc.Text = ""
         Else
             txtPurchaseactdesc.Text = PR_GlSub2("Acct_Desc")
             If txtConsumptionAct.Enabled Then txtConsumptionAct.SetFocus
         End If
 PR_GlSub2.Close
End If
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
Dim pr_dumy As New Recordset
pr_dumy.Open "select max(CatCode) as transcode from IC_ItemCategory where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 3)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 3)
End If
pr_dumy.Close
End Function

Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtPurchaseACT.SetFocus
End If
End Sub



Private Sub txtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtitemcode <> "" Then
        txtitemcode.Text = DoPad(UCase(txtitemcode.Text), txtitemcode.MaxLength)
        PR_ItemClass.Open "Select * from  IC_ItemCategory where compcode = '" & Gs_compcode & "' and CatCode = '" & txtitemcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        Select Case Mode
            Case "A"
                If Not PR_ItemClass.EOF Then
                   Call MsgBox(Gs_RecFdMsg, vbCritical)
                  'Cancel = True
                   Call ClearVal
                   txtitemcode.SetFocus
                Else
                   txtDesc.SetFocus
                End If
            Case Else
                If PR_ItemClass.EOF Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   Call ClearVal
                   txtitemcode.SetFocus
                Else
                   Call SetVal
                   txtDesc.SetFocus
                   
                End If
            End Select
  PR_ItemClass.Close
ElseIf KeyCode = vbKeyReturn And txtitemcode = "" Then
    txtitemcode = ""
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
      cmdLookup.Enabled = False
      txtitemcode.Locked = True
    Else
      txtitemcode.Locked = False
      cmdLookup.Enabled = True
    End If
    If PB_BlnkItmClass And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, Me, Me, txtitemcode, txtDesc, "X", "CompCount", 3, "ItemClass", "Description", 1, False, Toolbar1)
    End If
    If Mode = "A" Then
    txtitemcode = maxtranscode
    txtDesc.SetFocus
    End If
   
End Sub

Public Sub SaveValues()
'On Error GoTo LocalErr
Dim ls_sql As String

   

gc_dbcon.BeginTrans
     Select Case Mode
           Case "A"
           
             txtitemcode = maxtranscode
             Me.Refresh
           
              ls_sql = "INSERT into IC_ItemCategory(compcode,CatCode,Description,EmpDiscper,glpcode,glccode,glscode, ProfitPer) VALUES ('" & Gs_compcode & "','" & txtitemcode.Text & "','" & RepApp(txtDesc.Text) & "'," & Val(txtempDisc.Text) & ",'" & Trim(txtPurchaseACT.Text) & "','" & Trim(txtConsumptionAct.Text) & "','" & Trim(txtSaleAct.Text) & "'," & Val(txtempDisc.Text) & ")"
              gc_dbcon.Execute ls_sql
             
           Case "E"
              ls_sql = "UPDATE IC_ItemCategory SET Description= '" & RepApp(txtDesc.Text) & "', EmpDiscper=" & Val(txtempDisc.Text) & ",glpcode = '" & Trim(txtPurchaseACT) & "',glccode = '" & Trim(txtConsumptionAct) & "',glscode = '" & Trim(txtSaleAct) & "',profitper =" & Val(txtgprofit.Text) & "  WHERE  compcode = '" & Gs_compcode & "' and CatCode= '" & txtitemcode.Text & "'"
              gc_dbcon.Execute ls_sql
           Case "D"
           
              If pr_dumy.State = 1 Then pr_dumy.Close
              
              pr_dumy.Open "Select * from  IC_ItemClass where compcode = '" & Gs_compcode & "' and deptcode = '" & txtitemcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
              If pr_dumy.EOF Then
                ls_sql = "DELETE FROM IC_ItemCategory WHERE CatCode = '" & txtitemcode.Text & "' and compcode = '" & Gs_compcode & "'"
                gc_dbcon.Execute ls_sql
              Else
                Call MsgBox("Record Exist in Category Setup", vbCritical)
              End If
              
     End Select
gc_dbcon.CommitTrans


If Mode = "A" Then
    txtitemcode = maxtranscode
    End If
Exit Sub

LocalErr:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
End Sub
Public Sub ClearVal()
     txtitemcode = ""
     txtDesc = ""
     txtItemCount = ""
End Sub

Private Sub SetVal()
     txtDesc = Trim(PR_ItemClass("Description") & "")
     txtexpirydays = Val(0 & PR_ItemClass("SaleExpiryDays"))
     txtempDisc = Val(0 & PR_ItemClass("EmpDiscper"))
     txtPurchaseACT = Trim(PR_ItemClass("Glpcode") & "")
     txtConsumptionAct = Trim(PR_ItemClass("GlCcode") & "")
     txtSaleAct = Trim(PR_ItemClass("GlScode") & "")
     txtgprofit = Val(PR_ItemClass("ProfitPer"))
     If txtPurchaseACT <> "" Then Call txtPurchaseACT_KeyDown(vbKeyReturn, vbKeyShift)
     If txtConsumptionAct <> "" Then Call txtConsumptionAct_KeyDown(vbKeyReturn, vbKeyShift)
     If txtSaleAct <> "" Then Call txtSaleAct_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Public Function ChkInputs() As Boolean
    If Trim(txtitemcode.Text) = "" Then
       Call MsgBox("Enter Category Code !!!", vbCritical)
        txtitemcode.SetFocus
       ChkInputs = False
    ElseIf Trim(txtDesc.Text) = "" Then
       Call MsgBox("Enter Category Description !!!", vbCritical)
        txtDesc.SetFocus
        ChkInputs = False
    
    Else
        ChkInputs = True
    End If
End Function

Public Sub SetFrmEnv(ls_mode As String)
    txtDesc.Enabled = IIf(ls_mode <> "D", True, False)
End Sub


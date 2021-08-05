VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemSubCategory 
   Caption         =   "Sub Categories Setup"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   Icon            =   "frmItemSubCategory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1785
      Left            =   30
      TabIndex        =   1
      Top             =   570
      Width           =   7605
      Begin VB.CommandButton Command7 
         Height          =   315
         Left            =   2205
         Picture         =   "frmItemSubCategory.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   195
         Width           =   315
      End
      Begin VB.TextBox txtcatdesc 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   195
         Width           =   4935
      End
      Begin VB.TextBox txtcatcode 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   11
         Tag             =   "SKIPN"
         Top             =   210
         Width           =   525
      End
      Begin VB.TextBox txtclasscode 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   9
         Tag             =   "SKIPN"
         Top             =   585
         Width           =   525
      End
      Begin VB.TextBox txtClassDesc 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   585
         Width           =   4935
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   2205
         Picture         =   "frmItemSubCategory.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   585
         Width           =   315
      End
      Begin VB.TextBox txtItemCode 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   6
         Tag             =   "SKIPN"
         Top             =   960
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
         Left            =   1650
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1320
         Width           =   5850
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   315
         Left            =   2175
         Picture         =   "frmItemSubCategory.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   960
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Dept. Code :"
         Height          =   210
         Left            =   690
         TabIndex        =   14
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Category Code :"
         Height          =   210
         Left            =   480
         TabIndex        =   10
         Top             =   615
         Width           =   1170
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
         Left            =   720
         TabIndex        =   4
         Top             =   1335
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sub Category Code :"
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
         Left            =   135
         TabIndex        =   2
         Top             =   960
         Width           =   1500
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7680
      _ExtentX        =   13547
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
               Picture         =   "frmItemSubCategory.frx":0760
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemSubCategory.frx":0BB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemSubCategory.frx":1008
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemSubCategory.frx":145C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemSubCategory.frx":18B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemSubCategory.frx":1D04
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemSubCategory.frx":2458
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmItemSubCategory"
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
    Gs_SQL = "Select PackCode, Description from IC_ItemPacking "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  subcode = '" & txtclasscode & "' and deptcode = '" & txtcatcode & "'"
    MyLookupOLDB.Caption = "Sub Categories"
    MyLookupOLDB.Show 1
    If txtitemcode <> "" Then Call txtItemcode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub


Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtclasscode
    Set PO_DESC = txtClassDesc
    Gs_SQL = "Select ClassCode,   Description from IC_ItemClass "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' and deptcode = '" & txtcatcode & "'"
    MyLookupOLDB.Caption = "Categories"
    MyLookupOLDB.Show 1
    
    If txtclasscode <> "" Then Call txtclassCode_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command7_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtcatcode
    Set PO_DESC = txtcatdesc
    Gs_SQL = "Select CatCode,   Description from IC_ItemCategory "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Departments"
    MyLookupOLDB.Show 1
    
    If txtcatcode <> "" Then Call txtcatcode_KeyDown(vbKeyReturn, vbKeyShift)

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
pr_dumy.Open "select max(PackCode) as transcode from IC_ItemPacking where compcode = '" & Gs_compcode & "' and subcode = '" & txtclasscode & "' and deptcode = '" & txtcatcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 3)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 3)
End If
pr_dumy.Close
End Function

Private Sub txtcatcode_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtcatcode) <> "" And KeyCode = vbKeyReturn Then
        txtcatcode = DoPad(txtcatcode, 3)
        If pr_dumy.State = 1 Then pr_dumy.Close
        pr_dumy.Open "Select * from IC_ItemCategory where Catcode = '" & txtcatcode & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Department code not found !!!", vbCritical)
            txtcatcode = ""
            txtcatdesc = ""
            txtcatcode.SetFocus
        Else
            txtcatdesc = pr_dumy("Description")
            If txtclasscode.Enabled Then txtclasscode.SetFocus
            
        End If
        pr_dumy.Close
        
ElseIf Trim(txtcatcode) = "" And KeyCode = vbKeyReturn Then
        txtcatcode = ""
        txtcatdesc = ""
        Command7_Click
End If
End Sub


Private Sub txtclassCode_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtclasscode) <> "" And KeyCode = vbKeyReturn Then
        txtclasscode = DoPad(txtclasscode, 3)
       If pr_dumy.State = 1 Then pr_dumy.Close
        pr_dumy.Open "Select * from Ic_ItemClass where Classcode = '" & txtclasscode & "' and compcode = '" & Gs_compcode & "' and deptcode = '" & txtcatcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Category Code not found !!!", vbCritical)
            txtclasscode = ""
            txtClassDesc = ""
            txtclasscode.SetFocus
        Else
            txtClassDesc = pr_dumy("Description")
           
             If Mode = "A" Then
              txtitemcode = maxtranscode
             txtDesc.SetFocus
             Else
             txtitemcode.SetFocus
             End If
        End If
        pr_dumy.Close
        
ElseIf Trim(txtclasscode) = "" And KeyCode = vbKeyReturn Then
        txtclasscode = ""
        txtClassDesc = ""
        Command2_Click
End If
End Sub

Private Sub txtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtitemcode <> "" Then
        txtitemcode.Text = DoPad(UCase(txtitemcode.Text), txtitemcode.MaxLength)
        PR_ItemClass.Open "Select * from  IC_ItemPacking where compcode = '" & Gs_compcode & "' and PackCode = '" & txtitemcode & "' and  subcode = '" & txtclasscode & "' and deptcode = '" & txtcatcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        Select Case Mode
            Case "A"
                If Not PR_ItemClasPR_ItemClasss.EOF Then
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
  
    txtcatcode.SetFocus
   
End Sub

Public Sub SaveValues()
'On Error GoTo LocalErr
Dim ls_sql As String



gc_dbcon.BeginTrans
     Select Case Mode
           Case "A"
              txtitemcode = maxtranscode
              Me.Refresh
              ls_sql = "INSERT into IC_ItemPacking(compcode,Deptcode,subcode,PackCode,Description) VALUES ('" & Gs_compcode & "','" & txtcatcode & "','" & txtclasscode & "' ,'" & txtitemcode.Text & "','" & RepApp(txtDesc.Text) & "')"
              gc_dbcon.Execute ls_sql

             
           Case "E"
              ls_sql = "UPDATE IC_ItemPacking SET Subcode = '" & txtclasscode & "',  Description= '" & RepApp(txtDesc.Text) & "' WHERE subcode = '" & txtclasscode & "'  and  compcode = '" & Gs_compcode & "' and packcode= '" & txtitemcode.Text & "' and deptcode = '" & txtcatcode & "'"
              gc_dbcon.Execute ls_sql
           Case "D"
           
           
            If pr_dumy.State = 1 Then pr_dumy.Close
              
              pr_dumy.Open "Select * from  IC_Item where compcode = '" & Gs_compcode & "' and PackCode = '" & txtitemcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
              If pr_dumy.EOF Then
               
               ls_sql = "DELETE FROM IC_ItemPacking WHERE subcode = '" & txtclasscode & "'  and PackCode = '" & txtitemcode.Text & "' and compcode = '" & Gs_compcode & "' and deptcode = '" & txtcatcode & "'"
              gc_dbcon.Execute ls_sql
              Else
                Call MsgBox("Record Exist in Item Setup", vbCritical)
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
     txtDesc = PR_ItemClass("Description")
  
End Sub
Public Function ChkInputs() As Boolean
    If Trim(txtitemcode.Text) = "" Then
       Call MsgBox("Enter Item Class Code !!!", vbCritical)
        txtitemcode.SetFocus
       ChkInputs = False
    ElseIf Trim(txtDesc.Text) = "" Then
       Call MsgBox("Enter Item Class Description !!!", vbCritical)
        txtDesc.SetFocus
        ChkInputs = False
    Else
        ChkInputs = True
    End If
End Function

Public Sub SetFrmEnv(ls_mode As String)
    txtDesc.Enabled = IIf(ls_mode <> "D", True, False)
End Sub

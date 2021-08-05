VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemClass 
   Caption         =   "Categories Setup"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   Icon            =   "frmItemClass.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1395
      Left            =   30
      TabIndex        =   1
      Top             =   570
      Width           =   7605
      Begin VB.CommandButton Command7 
         Height          =   315
         Left            =   1875
         Picture         =   "frmItemClass.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtcatdesc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2235
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   210
         Width           =   5295
      End
      Begin VB.TextBox txtcatcode 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   1335
         MaxLength       =   3
         TabIndex        =   7
         Tag             =   "SKIPN"
         Top             =   225
         Width           =   510
      End
      Begin VB.TextBox txtItemCode 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   1335
         MaxLength       =   3
         TabIndex        =   6
         Tag             =   "SKIPN"
         Top             =   615
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
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   975
         Width           =   6165
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   315
         Left            =   1860
         Picture         =   "frmItemClass.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   615
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Dept. Code :"
         Height          =   210
         Left            =   375
         TabIndex        =   10
         Top             =   255
         Width           =   885
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
         Left            =   405
         TabIndex        =   4
         Top             =   990
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Category Code :"
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
         Top             =   615
         Width           =   1170
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
               Picture         =   "frmItemClass.frx":05EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemClass.frx":0A42
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemClass.frx":0E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemClass.frx":12EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemClass.frx":173E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemClass.frx":1B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemClass.frx":22E6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmItemClass"
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
    Gs_SQL = "Select ClassCode, Description from IC_ItemClass "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and deptcode ='" & txtcatcode & "'"
    MyLookupOLDB.Caption = "Category"
    MyLookupOLDB.Show 1
    If txtitemcode <> "" Then Call txtItemcode_KeyDown(vbKeyReturn, vbKeyShift)
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
pr_dumy.Open "select max(ClassCode) as transcode from IC_ItemClass where compcode = '" & Gs_compcode & "' and deptcode = '" & txtcatcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
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
            If Mode = "A" Then
            txtitemcode = maxtranscode
            txtDesc.SetFocus
            Else
            txtitemcode.SetFocus
            End If
        End If
        pr_dumy.Close
        
ElseIf Trim(txtcatcode) = "" And KeyCode = vbKeyReturn Then
        txtcatcode = ""
        txtcatdesc = ""
        Command7_Click
End If
End Sub

Private Sub txtDesc_LostFocus()
If txtDesc <> "" Then txtDesc = UCase(txtDesc)
End Sub

Private Sub txtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtitemcode <> "" Then
        txtitemcode.Text = DoPad(UCase(txtitemcode.Text), txtitemcode.MaxLength)
        PR_ItemClass.Open "Select * from  IC_ItemClass where compcode = '" & Gs_compcode & "' and classcode = '" & txtitemcode & "' and deptcode = '" & txtcatcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
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
    cmdLookup_Click
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
              
              ls_sql = "INSERT into IC_ItemClass(compcode,deptcode,ClassCode,Description) VALUES ('" & Gs_compcode & "','" & txtcatcode.Text & "','" & txtitemcode.Text & "','" & RepApp(txtDesc.Text) & "')"
              gc_dbcon.Execute ls_sql

             
           Case "E"
              ls_sql = "UPDATE IC_ItemClass SET deptcode = '" & txtcatcode & "', Description= '" & RepApp(txtDesc.Text) & "' WHERE  Deptcode = '" & txtcatcode & "' and  compcode = '" & Gs_compcode & "' and Classcode= '" & txtitemcode.Text & "'"
              gc_dbcon.Execute ls_sql
           Case "D"
           
             If pr_dumy.State = 1 Then pr_dumy.Close
              
              pr_dumy.Open "Select * from  IC_ItemPacking where compcode = '" & Gs_compcode & "' and subcode = '" & txtitemcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
              If pr_dumy.EOF Then
               
               ls_sql = "DELETE FROM IC_ItemClass WHERE  Deptcode = '" & txtcatcode & "'  and ClassCode = '" & txtitemcode.Text & "' and compcode = '" & Gs_compcode & "'"
               gc_dbcon.Execute ls_sql
            
              Else
                Call MsgBox("Record Exist in Sub Category Setup", vbCritical)
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

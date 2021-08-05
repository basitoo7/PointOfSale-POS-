VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemMaterial 
   Caption         =   "Item Material"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   Icon            =   "frmItemMaterial.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   0
      TabIndex        =   1
      Top             =   570
      Width           =   4935
      Begin VB.TextBox txtMTCode 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   1320
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
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   600
         Width           =   3495
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   315
         Left            =   1845
         Picture         =   "frmItemMaterial.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   240
         Width           =   315
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
         Left            =   330
         TabIndex        =   4
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Material Code :"
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
         Left            =   165
         TabIndex        =   2
         Top             =   240
         Width           =   1065
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4950
      _ExtentX        =   8731
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
               Picture         =   "frmItemMaterial.frx":047C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemMaterial.frx":08D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemMaterial.frx":0D24
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemMaterial.frx":1178
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemMaterial.frx":15CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemMaterial.frx":1A20
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemMaterial.frx":2174
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmItemMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PO_CODE As Object
Public PO_DESC As Object
Dim Mode As String
Dim lb_found As Boolean
Dim PR_ItemClass As New Recordset

Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtMTCode
    Set PO_DESC = txtDesc
    Gs_SQL = "Select MTCode, Description from IC_ItemMaterial"
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Item Class"
    MyLookupOLDB.Show 1
    If txtMTCode <> "" Then Call txtMTCode_KeyDown(vbKeyReturn, vbKeyShift)
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
PR_Dumy.Open "select max(MTCode) as transcode from IC_ItemMaterial where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not PR_Dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & PR_Dumy("transcode")) + 1)), 3)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 3)
End If
PR_Dumy.Close
End Function

Private Sub txtMTCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtMTCode <> "" Then
        txtMTCode.Text = DoPad(UCase(txtMTCode.Text), txtMTCode.MaxLength)
        PR_ItemClass.Open "Select * from  IC_ItemMaterial where compcode = '" & Gs_compcode & "' and MTCode = '" & txtMTCode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        Select Case Mode
            Case "A"
                If Not PR_ItemClass.EOF Then
                   Call MsgBox(Gs_RecFdMsg, vbCritical)
                  Cancel = True
                   Call ClearVal
                   txtMTCode.SetFocus
                Else
                   txtDesc.SetFocus
                End If
            Case Else
                If PR_ItemClass.EOF Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   Call ClearVal
                   txtMTCode.SetFocus
                Else
                   Call SetVal
                   txtDesc.SetFocus
                  
                End If
            End Select
  PR_ItemClass.Close
ElseIf KeyCode = vbKeyReturn And txtMTCode = "" Then
    txtMTCode = ""
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
      cmdLookup.Enabled = False
      txtMTCode.Locked = True
    Else
      txtMTCode.Locked = False
      cmdLookup.Enabled = True
    End If
    If PB_BlnkItmClass And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, Me, Me, txtMTCode, txtDesc, "X", "CompCount", 3, "ItemClass", "Description", 1, False, Toolbar1)
    End If
    If Mode = "A" Then
    txtMTCode = maxtranscode
    txtDesc.SetFocus
    End If
   
End Sub

Public Sub SaveValues()
'On Error GoTo LocalErr
Dim ls_sql As String



gc_dbcon.BeginTrans
     Select Case Mode
           Case "A"
              ls_sql = "INSERT into IC_ItemMaterial(compcode,MTCode,Description) VALUES ('" & Gs_compcode & "','" & txtMTCode.Text & "','" & RepApp(txtDesc.Text) & "')"
              gc_dbcon.Execute ls_sql
             
           Case "E"
              ls_sql = "UPDATE IC_ItemMaterial SET Description= '" & RepApp(txtDesc.Text) & "' WHERE  compcode = '" & Gs_compcode & "' and MTCode= '" & txtMTCode.Text & "'"
              gc_dbcon.Execute ls_sql
           Case "D"
              ls_sql = "DELETE FROM IC_ItemMaterial WHERE MTCode = '" & txtMTCode.Text & "' and compcode = '" & Gs_compcode & "'"
              gc_dbcon.Execute ls_sql
           
     End Select
gc_dbcon.CommitTrans


If Mode = "A" Then
    txtMTCode = maxtranscode
    End If
Exit Sub

LocalErr:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
End Sub
Public Sub ClearVal()
     txtMTCode = ""
     txtDesc = ""
     txtItemCount = ""
End Sub

Private Sub SetVal()
     txtDesc = PR_ItemClass("Description")
End Sub
Public Function ChkInputs() As Boolean
    If Trim(txtMTCode.Text) = "" Then
       Call MsgBox("Enter Item Pack Code !!!", vbCritical)
        txtMTCode.SetFocus
       ChkInputs = False
    ElseIf Trim(txtDesc.Text) = "" Then
       Call MsgBox("Enter Item Pack Description !!!", vbCritical)
        txtDesc.SetFocus
        ChkInputs = False
    Else
        ChkInputs = True
    End If
End Function

Public Sub SetFrmEnv(ls_mode As String)
    txtDesc.Enabled = IIf(ls_mode <> "D", True, False)
End Sub


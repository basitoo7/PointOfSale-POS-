VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSMModules 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Module Information"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   Icon            =   "frmSMModules.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameGroupInformation 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   960
      Left            =   60
      TabIndex        =   0
      Top             =   570
      Width           =   4815
      Begin VB.TextBox txtModuleDesc 
         Height          =   315
         Left            =   990
         MaxLength       =   50
         TabIndex        =   3
         Top             =   540
         Width           =   3735
      End
      Begin VB.TextBox txtModuleCode 
         BackColor       =   &H00FFFF80&
         Height          =   315
         Left            =   990
         MaxLength       =   3
         TabIndex        =   2
         Tag             =   "SKIPN"
         Top             =   180
         Width           =   525
      End
      Begin VB.CommandButton cmdModuleLookup 
         Height          =   315
         Left            =   1605
         Picture         =   "frmSMModules.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   165
         Width           =   315
      End
      Begin VB.Label lblGroupDesc 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description :"
         Height          =   195
         Left            =   75
         TabIndex        =   5
         Top             =   585
         Width           =   885
      End
      Begin VB.Label lblModuleCode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code :"
         Height          =   195
         Left            =   510
         TabIndex        =   4
         Top             =   210
         Width           =   465
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   1005
      ButtonWidth     =   1376
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
            Caption         =   "&Refresh"
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
               Picture         =   "frmSMModules.frx":047C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMModules.frx":08D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMModules.frx":0D24
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMModules.frx":1178
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMModules.frx":15CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMModules.frx":1A20
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMModules.frx":2174
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSMModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_Blnk As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Modules As New Recordset

Private Sub cmdModuleLookup_Click()

    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtModuleCode
    Set PO_DESC = txtModuleDesc
    
    Gs_SQL = "Select ModuleCode as Code,ModuleDesc as Description from Sys_Modules "
    Gs_FindFld = "ModuleDesc"
    Gs_OrderBy = "order by ModuleDesc"
    MyLookupOLDB.Caption = "Modules - " & App.ProductName
    MyLookupOLDB.Show 1
    txtModuleCode.SetFocus
    
    If Len(txtModuleCode) > 0 Then
        txtModuleCode_Validate False
        SendKeys vbTab
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF11 Then Mode = DentMode(Mode, 4, PR_Modules, Me, txtModuleCode, txtModuleDesc, "X", "CompCount", 3, "Acct_sub0", "Acct_Desc", 1, False, Toolbar1)
    
    If KeyCode = vbKeyReturn Then
    
        SendKeys vbTab
        
    End If
    
End Sub

Private Sub Form_Load()
  
    SetToolBar(1) = True
    SetToolBar(2) = True
    SetToolBar(3) = True
    SetToolBar(4) = True
    
    Toolbar1.Buttons(1).Enabled = SetToolBar(1)
    Toolbar1.Buttons(2).Enabled = SetToolBar(2)
    Toolbar1.Buttons(3).Enabled = SetToolBar(3)
    Toolbar1.Buttons(5).Enabled = SetToolBar(4)
    
    PR_Modules.Open "Select ModuleCode, ModuleDesc from Sys_Modules order by ModuleCode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
     
    PB_Blnk = IIf(PR_Modules.EOF, True, False)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    PR_Modules.Close
    
End Sub

Private Sub txtModuleCode_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF12 Then
        Call cmdModuleLookup_Click
    ElseIf KeyCode = vbKeyReturn And Mode = "D" Then
        Call txtModuleCode_Validate(True)
    End If
    
End Sub

Private Sub txtModuleCode_Validate(Cancel As Boolean)
    
    Dim lb_found As Boolean
    
    If Trim(txtModuleCode) <> "" Then
    
        txtModuleCode = UCase(Trim(txtModuleCode))
        
        lb_found = MySeek(txtModuleCode.Text, "ModuleCode", PR_Modules)
                      
        Select Case Mode
        
            Case "A"
            
                If lb_found Then
                
                    Call SetErr(Gs_RecFdMsg, vbCritical)
                    'Cancel = True
                    txtModuleCode = ""
                    txtModuleDesc = ""
                    txtModuleCode.SetFocus
                    
                Else
                
                    txtModuleDesc.SetFocus
                    
                End If
                
            Case Else
            
                If Not lb_found Then
                
                    Call SetErr(Gs_RecNFMsg, vbCritical)
                    'Cancel = True
                    txtModuleCode = ""
                    txtModuleDesc = ""
                    txtModuleCode.SetFocus
                Else
                
                    txtModuleDesc = Trim("" & PR_Modules("ModuleDesc"))
                    
                End If
                
            End Select
            
        Else
        
            txtModuleCode = ""
            txtModuleDesc = ""
        
        End If
            
End Sub
  
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If PB_Blnk And Range(Button.Index, 2, 3) Then
    
        MsgBox "Data not found :", vbCritical
        Mode = ""
        'Cancel = True
        
    Else
        
        Mode = DentMode(Mode, Button.Index, PR_Modules, Me, txtModuleCode, txtModuleDesc, "X", "CompCount", 3, "Acct_sub0", "Acct_Desc", 1, False, Toolbar1)
        
        If Mode = "A" Then
             Toolbar1.Buttons(1).Enabled = False
             cmdModuleLookup.Enabled = False
         Else
             cmdModuleLookup.Enabled = True
        End If
        
    End If
End Sub

Public Sub SaveValues()
On Error GoTo LocalErr
Dim cntsql As New ADODB.Command

PB_Blnk = False
cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText
gc_dbcon.BeginTrans
     Select Case Mode
           Case "A"
              cntsql.CommandText = "INSERT INTO Sys_Modules(ModuleCODE, ModuleDESC, ADDDATETIME) VALUES ('" & RepApp(UCase(Trim(txtModuleCode))) & "','" & UCase(RepApp(Trim(txtModuleDesc))) & "','" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "')"
              cntsql.Execute
              txtModuleCode = ""
           Case "E"
              cntsql.CommandText = "UPDATE Sys_Modules Set ModuleDESC = '" & UCase(RepApp(Trim(txtModuleDesc.Text))) & "', MODIFYDATETIME = '" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' WHERE ModuleCODE = '" & UCase(RepApp(Trim(txtModuleCode))) & "'"
              cntsql.Execute
           Case "D"
              cntsql.CommandText = "DELETE FROM Sys_Modules WHERE ModuleCODE = '" & UCase(RepApp(Trim(txtModuleCode))) & "'"
              cntsql.Execute
     End Select
gc_dbcon.CommitTrans
PR_Modules.Requery

Exit Sub

LocalErr:
gc_dbcon.RollbackTrans
MsgBox Err.Description
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Public Function ChkInputs() As Boolean
    If Len(Trim(txtModuleCode.Text)) > 0 And Len(RTrim(txtModuleDesc)) > 0 Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub FrmRefresh()
    PR_Modules.Requery
End Sub

Private Sub txtModuleDesc_Validate(Cancel As Boolean)
    txtModuleDesc.Text = UCase(txtModuleDesc.Text)
End Sub



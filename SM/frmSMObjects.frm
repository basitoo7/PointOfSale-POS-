VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSMObjects 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Objects"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   Icon            =   "frmSMObjects.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4965
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
      Height          =   1305
      Left            =   60
      TabIndex        =   3
      Top             =   570
      Width           =   4875
      Begin VB.TextBox txtObjectDesc 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   2
         Top             =   885
         Width           =   3705
      End
      Begin VB.CommandButton cmdObjectLookup 
         Height          =   315
         Left            =   2190
         Picture         =   "frmSMObjects.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   495
         Width           =   315
      End
      Begin VB.TextBox txtObjectCode 
         BackColor       =   &H00FFFF80&
         Height          =   315
         Left            =   1065
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "SKIP"
         Top             =   525
         Width           =   1125
      End
      Begin VB.CommandButton cmdModuleLookup 
         Height          =   315
         Left            =   1605
         Picture         =   "frmSMObjects.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   150
         Width           =   315
      End
      Begin VB.TextBox txtModuleCode 
         BackColor       =   &H00FFFF80&
         Height          =   315
         Left            =   1065
         MaxLength       =   3
         TabIndex        =   0
         Tag             =   "SKIPN"
         Top             =   150
         Width           =   525
      End
      Begin VB.TextBox txtModuleDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1935
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   150
         Width           =   2850
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description :"
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   915
         Width           =   885
      End
      Begin VB.Label lblObjectCode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code :"
         Height          =   195
         Left            =   555
         TabIndex        =   9
         Top             =   555
         Width           =   465
      End
      Begin VB.Label lblModule 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Module :"
         Height          =   195
         Left            =   405
         TabIndex        =   6
         Top             =   180
         Width           =   615
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4965
      _ExtentX        =   8758
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
               Picture         =   "frmSMObjects.frx":05EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMObjects.frx":0A42
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMObjects.frx":0E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMObjects.frx":12EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMObjects.frx":173E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMObjects.frx":1B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMObjects.frx":22E6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSMObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_Blnk As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Modules As New Recordset
Dim PR_Objects As New Recordset

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
        'SendKeys vbTab
    End If

End Sub

Private Sub cmdObjectLookup_Click()

    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtObjectCode
    Set PO_DESC = txtObjectDesc
    
    Gs_SQL = "Select ObjectCode as Code,ObjectDesc as Description from Sys_Objects "
    Gs_FindFld = "ObjectDesc"
    Gs_OtherPara = " Where ModuleCode = '" & Trim(txtModuleCode) & "' "
    Gs_OrderBy = "order by ObjectDesc"
    MyLookupOLDB.Caption = "Objects - " & App.ProductName
    MyLookupOLDB.Show 1
    txtObjectCode.SetFocus
    
    If Len(txtObjectCode) > 0 Then
        txtObjectCode_Validate False
        SendKeys vbTab
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF11 Then Mode = DentMode(Mode, 4, PR_Objects, Me, txtObjectCode, txtObjectDesc, "X", "CompCount", 3, "Acct_sub0", "Acct_Desc", 1, False, Toolbar1)
    
    If KeyCode = vbKeyReturn Then
    
        'SendKeys vbTab
        
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

    PR_Modules.Open "Select ModuleCode, ModuleDesc from Sys_Modules order by ModuleCode", gc_dbcon, adOpenDynamic, adLockReadOnly, 1
    PR_Objects.Open "Select ModuleCode, ObjectCode, ObjectDesc from Sys_Objects order by ObjectCode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
    
    'PR_Objects.Filter = " ModuleCode = '" & txtModuleCode & "'"
    
    PB_Blnk = IIf(PR_Objects.EOF, True, False)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    PR_Objects.Close
    PR_Modules.Close
    
End Sub

Private Sub txtModuleCode_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF12 Then
        
        Call cmdModuleLookup_Click
        
    End If
    
End Sub

Private Sub txtModuleCode_Validate(Cancel As Boolean)
    
    Dim lb_found As Boolean
    
    If Trim(txtModuleCode) <> "" Then
    
        txtModuleCode = UCase(Trim(txtModuleCode))
        
        lb_found = MySeek(txtModuleCode.Text, "ModuleCode", PR_Modules)
                      
        If Not lb_found Then
        
            Call SetErr(Gs_RecNFMsg, vbCritical)
            'Cancel = True
            txtModuleCode = ""
            txtModuleDesc = ""
            txtModuleCode.SetFocus
        Else
        
            txtModuleDesc = Trim("" & PR_Modules("ModuleDesc"))
            PR_Objects.Filter = adFilterNone
            If Mode <> "A" Then
                PR_Objects.Filter = " ModuleCode = '" & txtModuleCode & "'"
                lb_found = MySeek(txtObjectCode.Text, "ObjectCode", PR_Objects)
                If Not lb_found Then
                    txtObjectCode = ""
                    txtObjectDesc = ""
                End If
            End If
            
        End If
                
            
    Else
        
        txtModuleCode = ""
        txtModuleDesc = ""
        PR_Objects.Filter = adFilterNone
        
    End If
            
End Sub

Private Sub txtObjectCode_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF12 Then
        
        Call cmdObjectLookup_Click
        
    End If

End Sub

Private Sub txtObjectCode_Validate(Cancel As Boolean)
    
    Dim lb_found As Boolean
    
    If Trim(txtObjectCode) <> "" Then
    
        txtObjectCode = UCase(Trim(txtObjectCode))
        
        lb_found = MySeek(txtObjectCode.Text, "ObjectCode", PR_Objects)
                      
        Select Case Mode
        
            Case "A"
            
                If lb_found Then
                
                    Call SetErr(Gs_RecFdMsg, vbCritical)
                    'Cancel = True
                    txtObjectCode = ""
                    txtObjectDesc = ""
                    txtObjectCode.SetFocus
                    
                Else
                
                    txtObjectDesc.SetFocus
                    
                End If
                
            Case Else
            
                If Not lb_found Then
                
                    Call SetErr(Gs_RecNFMsg, vbCritical)
                    'Cancel = True
                    txtObjectCode = ""
                    txtObjectDesc = ""
                    txtObjectCode.SetFocus
                Else
                
                    txtObjectDesc = Trim("" & PR_Objects("ObjectDesc"))
                    
                End If
                
            End Select
            
        Else
        
            txtObjectCode = ""
            txtObjectDesc = ""
        
        End If
            
End Sub
  
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If PB_Blnk And Range(Button.Index, 2, 3) Then
    
        MsgBox "Data not found :", vbCritical
        Mode = ""
        'Cancel = True
        
    Else
        
        Mode = DentMode(Mode, Button.Index, PR_Objects, Me, txtObjectCode, txtObjectDesc, "X", "CompCount", 3, "Acct_sub0", "Acct_Desc", 1, False, Toolbar1)
        
        If Mode = "A" Then
             Toolbar1.Buttons(1).Enabled = False
             cmdObjectLookup.Enabled = False
             PR_Objects.Filter = adFilterNone
         Else
             cmdObjectLookup.Enabled = True
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
              cntsql.CommandText = "INSERT INTO Sys_Objects(ModuleCode, ObjectCode, ObjectDesc, ADDDATETIME) VALUES ('" & RepApp(UCase(Trim(txtModuleCode))) & "','" & UCase(RepApp(Trim(txtObjectCode))) & "','" & UCase(RepApp(Trim(txtObjectDesc))) & "','" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "')"
              cntsql.Execute
           Case "E"
              cntsql.CommandText = "UPDATE Sys_Objects Set ObjectDESC = '" & UCase(RepApp(Trim(txtObjectDesc.Text))) & "', MODIFYDATETIME = '" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "' WHERE ObjectCODE = '" & UCase(RepApp(Trim(txtObjectCode))) & "'"
              cntsql.Execute
           Case "D"
              cntsql.CommandText = "DELETE FROM Sys_Objects WHERE ObjectCODE = '" & UCase(RepApp(Trim(txtObjectCode))) & "'"
              cntsql.Execute
     End Select
gc_dbcon.CommitTrans
PR_Objects.Requery

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
    PR_Objects.Requery
    PR_Modules.Requery
End Sub

Private Sub txtModuleDesc_Validate(Cancel As Boolean)
    txtModuleDesc.Text = UCase(txtModuleDesc.Text)
End Sub

Private Sub txtObjectDesc_Validate(Cancel As Boolean)
    txtObjectDesc.Text = UCase(txtObjectDesc.Text)
End Sub

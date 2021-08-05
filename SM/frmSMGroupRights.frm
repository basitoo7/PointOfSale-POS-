VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSMGroupRights 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Group Rights"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   Icon            =   "frmSMGroupRights.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameGroup 
      Height          =   930
      Left            =   30
      TabIndex        =   6
      Top             =   570
      Width           =   6030
      Begin VB.TextBox txtGroupDesc 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1005
         MaxLength       =   50
         TabIndex        =   1
         Top             =   525
         Width           =   3840
      End
      Begin VB.TextBox txtGroupCode 
         BackColor       =   &H00FFFF80&
         Height          =   315
         Left            =   1005
         MaxLength       =   4
         TabIndex        =   0
         Tag             =   "SKIPN"
         Top             =   150
         Width           =   630
      End
      Begin VB.CommandButton cmdGroupLookup 
         Height          =   315
         Left            =   1665
         Picture         =   "frmSMGroupRights.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   135
         Width           =   315
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description :"
         Height          =   195
         Left            =   75
         TabIndex        =   18
         Top             =   555
         Width           =   885
      End
      Begin VB.Label lblGroup 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group :"
         Height          =   195
         Left            =   435
         TabIndex        =   8
         Top             =   180
         Width           =   525
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   1058
      ButtonWidth     =   1482
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
               Picture         =   "frmSMGroupRights.frx":047C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMGroupRights.frx":08D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMGroupRights.frx":0D24
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMGroupRights.frx":1178
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMGroupRights.frx":15CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMGroupRights.frx":1A20
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMGroupRights.frx":2174
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frameObjects 
      Height          =   3360
      Left            =   45
      TabIndex        =   10
      Top             =   1440
      Width           =   6015
      Begin VB.CommandButton cmdAddAll 
         Caption         =   "A&dd All"
         Height          =   330
         Left            =   4890
         TabIndex        =   3
         Top             =   195
         Width           =   1020
      End
      Begin VB.CommandButton cmdAddtoGrid 
         Caption         =   "&Add to Grid"
         Height          =   330
         Left            =   4905
         TabIndex        =   5
         Top             =   600
         Width           =   1020
      End
      Begin VB.TextBox txtModuleDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1860
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   210
         Width           =   2985
      End
      Begin VB.TextBox txtModuleCode 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   975
         MaxLength       =   3
         TabIndex        =   2
         Top             =   225
         Width           =   525
      End
      Begin VB.CommandButton cmdModuleLookup 
         Height          =   315
         Left            =   1515
         Picture         =   "frmSMGroupRights.frx":25C8
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtObjectCode 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   975
         MaxLength       =   10
         TabIndex        =   4
         Top             =   615
         Width           =   1125
      End
      Begin VB.CommandButton cmdObjectLookup 
         Height          =   315
         Left            =   2115
         Picture         =   "frmSMGroupRights.frx":273A
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   615
         Width           =   315
      End
      Begin VB.TextBox txtObjectDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2460
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   630
         Width           =   2385
      End
      Begin MSFlexGridLib.MSFlexGrid grdGroup 
         Height          =   2310
         Left            =   45
         TabIndex        =   11
         Top             =   1020
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   4075
         _Version        =   393216
         Rows            =   1
      End
      Begin VB.Label lblModule 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Module :"
         Height          =   195
         Left            =   330
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Object :"
         Height          =   195
         Left            =   390
         TabIndex        =   14
         Top             =   630
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmSMGroupRights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_Blnk As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Groups As New Recordset
Dim PR_Objects As New Recordset
Dim PR_Modules As New Recordset
Dim PR_UserRights As New Recordset

Dim PI_CurRow    As Integer
Dim PI_SrNo      As Integer
Dim PS_RowClicked As String
Dim lb_found As Boolean


Private Sub cmdAddAll_Click()
    If Trim(txtModuleCode) <> "" Then
        GoTop PR_Objects
        With grdGroup
            
        Do While Not PR_Objects.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                .TextMatrix(.Row, 1) = Trim("" & PR_Objects("ObjectCode"))
                .TextMatrix(.Row, 2) = Trim("" & PR_Objects("ObjectDesc"))
                .TextMatrix(.Row, 3) = Trim("" & PR_Objects("ModuleCode"))
                .Rows = .Rows + 1
             PR_Objects.MoveNext
           If PR_Objects.EOF Then Exit Do
        Loop
        If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
    End If
End Sub

Private Sub cmdAddtoGrid_Click()
    If Trim(txtModuleCode) <> "" And Trim(txtObjectCode) <> "" Then
     AddToGrid
    End If
End Sub

Private Sub AddToGrid()
Dim ln_cnt As Integer
         If txtGroupCode.Text <> "" Then
                    If PS_RowClicked = "" Then
                        If PI_SrNo = 0 Then
                            PI_SrNo = 1
                        Else
                            PI_SrNo = PI_SrNo + 1
                         End If
                     End If
        
                        With grdGroup
                            If PS_RowClicked = "" Then
                                    If Not PI_SrNo = 1 Then .Rows = .Rows + 1
                                        .Row = .Rows - 1
                                    Else
                                        .Row = PI_CurRow
                                    End If
                                    
                                    If PS_RowClicked = "" Then
                                        .TextMatrix(.Row, 0) = PI_SrNo
                                    Else
                                        .TextMatrix(.Row, 0) = PI_CurRow
                                    End If
                                       .TextMatrix(.Row, 1) = txtObjectCode
                                       .TextMatrix(.Row, 2) = txtObjectDesc
                                       .TextMatrix(.Row, 3) = txtModuleCode
                           
                        End With
                         txtObjectCode = ""
                         txtObjectDesc = ""
                         txtObjectCode.SetFocus
        
                   
        Else
            Call MsgBox("Enter Object Code !!!", vbCritical)
            txtObjectCode.SetFocus
        End If
      

End Sub

Private Sub cmdGroupLookup_Click()

    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtGroupCode
    Set PO_DESC = txtGroupDesc
    
    Gs_SQL = "Select GroupCode as Code,GroupDesc as Description from Sys_Groups "
    Gs_FindFld = "GroupDesc"
    Gs_OrderBy = "order by GroupDesc"
    MyLookupOLDB.Caption = "Groups - " & App.ProductName
    MyLookupOLDB.Show 1
    txtGroupCode.SetFocus
    
    If Len(txtGroupCode) > 0 Then
        txtGroupCode_Validate False
        SendKeys vbTab
    End If

End Sub

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

Private Sub cmdObjectLookup_Click()

    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtObjectCode
    Set PO_DESC = txtObjectDesc
    
    Gs_SQL = "Select ObjectCode as Code,ObjectDesc as Description from Sys_Objects "
    Gs_FindFld = "ObjectDesc"
    Gs_OtherPara = " Where ModuleCode = '" & RepApp(Trim(txtModuleCode)) & "' "
    Gs_OrderBy = "order by ObjectDesc"
    MyLookupOLDB.Caption = "Objects - " & App.ProductName
    MyLookupOLDB.Show 1
  
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF11 Then Mode = DentMode(Mode, 4, PR_Groups, Me, txtGroupCode, txtGroupDesc, "X", "CompCount", 3, "Acct_sub0", "Acct_Desc", 1, False, Toolbar1)
    
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
    
    PR_Groups.Open "Select GroupCode, GroupDesc from Sys_Groups order by GroupCode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
    PR_Modules.Open "Select ModuleCode, ModuleDesc from Sys_Modules order by ModuleCode", gc_dbcon, adOpenDynamic, adLockReadOnly, 1
    PR_Objects.Open "Select ModuleCode, ObjectCode, ObjectDesc from Sys_Objects order by ObjectCode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
    PR_UserRights.Open "Select * from Sys_UserRights Where GroupFlag = 1 Order By GroupOrUserCode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
    
    PR_Objects.Filter = " ModuleCode = '" & txtModuleCode & "'"
     
    PB_Blnk = IIf(PR_Groups.EOF, True, False)
    Call InitializeGrid
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    PR_UserRights.Close
    PR_Groups.Close
    PR_Modules.Close
    PR_Objects.Close

End Sub

Private Sub grdGroup_DblClick()

    With grdGroup
    
        If .Row = .Rows - 1 Then Exit Sub
        txtObjectCode = .TextMatrix(.Row, 1)
        txtObjectDesc = .TextMatrix(.Row, 2)
        txtModuleCode = .TextMatrix(.Row, 3)
        txtModuleCode_Validate True
    
    End With

End Sub

Private Sub grdGroup_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
    
        With grdGroup
            
            If (.Row = 1 Or .Row = 0) And .Rows = 2 Then Exit Sub
            If .Row + 1 = .Rows Then Exit Sub
            .RemoveItem .Row
            
        End With
    
    End If

End Sub

Private Sub txtGroupCode_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF12 Then
        Call cmdGroupLookup_Click
    ElseIf KeyCode = vbKeyReturn And Mode = "D" Then
        Call txtGroupCode_Validate(True)
    End If
    
End Sub

Private Sub txtGroupCode_Validate(Cancel As Boolean)
    
    Dim lb_found As Boolean
    
    If Trim(txtGroupCode) <> "" Then
    
        txtGroupCode = Val(txtGroupCode)
        
        lb_found = MySeek(txtGroupCode.Text, "GroupCode", PR_Groups)
                      
        Select Case Mode
        
            Case "A"
            
                If lb_found Then
                
                    Call SetErr(Gs_RecFdMsg, vbCritical)
                    'Cancel = True
                    txtGroupCode = ""
                    txtGroupDesc = ""
                    txtGroupCode.SetFocus
                    
                Else
                
                    txtGroupDesc.SetFocus
                    
                End If
                
            Case Else
            
                If Not lb_found Then
                
                    Call SetErr(Gs_RecNFMsg, vbCritical)
                    'Cancel = True
                    txtGroupCode = ""
                    txtGroupDesc = ""
                    txtGroupCode.SetFocus
                Else
                
                    
                    Call InitializeGrid
                    txtGroupDesc = Trim("" & PR_Groups("GroupDesc"))
                    
                    PR_Objects.Filter = adFilterNone
                    PR_UserRights.Filter = " GroupOrUserCode = " & txtGroupCode & ""
                    PR_UserRights.Requery
                    With grdGroup
                    Do While Not PR_UserRights.EOF
                            .Row = .Rows - 1
                            .TextMatrix(.Row, 0) = .Row
                            .TextMatrix(.Row, 1) = PR_UserRights("ObjectCode")
                             If MySeek(Trim(PR_UserRights("ObjectCode")), "ObjectCode", PR_Objects) Then
                             .TextMatrix(.Row, 2) = PR_Objects("ObjectDesc")
                             End If
                            .TextMatrix(.Row, 3) = PR_UserRights("ModuleCode")
                            .Rows = .Rows + 1
                        
                           If PR_UserRights.EOF Then Exit Do
                           PR_UserRights.MoveNext
                    
                    Loop
                    If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
                    End With
                    PR_UserRights.Filter = adFilterNone
                    PR_Objects.Filter = " ModuleCode = '" & txtModuleCode & "'"
                    
                End If
                
            End Select
            
        Else
        
            txtGroupCode = ""
            txtGroupDesc = ""
            Call InitializeGrid
        
        End If
            
End Sub
  
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If PB_Blnk And Range(Button.Index, 2, 3) Then
    
        MsgBox "Data not found :", vbCritical
        Mode = ""
        'Cancel = True
        
    Else
        
        Mode = DentMode(Mode, Button.Index, PR_Groups, Me, txtGroupCode, txtGroupDesc, "X", "CompCount", 3, "Acct_sub0", "Acct_Desc", 1, False, Toolbar1)
        
        If Mode = "A" Then
             Toolbar1.Buttons(1).Enabled = False
             txtGroupCode = SerMaxNo(gc_dbcon, "Sys_Groups", "GroupCode")
             txtGroupCode.Locked = True
             cmdGroupLookup.Enabled = False
         Else
             txtGroupCode.Locked = False
             cmdGroupLookup.Enabled = True
        End If
        
        If Range(Button.Index, 1, 3) Then Call InitializeGrid
        
    End If
End Sub

Private Sub txtGroupDesc_Validate(Cancel As Boolean)
    txtGroupDesc.Text = UCase(txtGroupDesc.Text)
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
            PR_Objects.Filter = " ModuleCode = '" & RepApp(txtModuleCode) & "'"
            lb_found = MySeek(txtObjectCode.Text, "ObjectCode", PR_Objects)
            If Not lb_found Then
                txtObjectCode = ""
                txtObjectDesc = ""
            End If
            
        End If
                
            
    Else
        
        txtModuleCode = ""
        txtModuleDesc = ""
        PR_Objects.Filter = adFilterNone
        
    End If

End Sub


Public Sub InitializeGrid()

    With grdGroup
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Object Code |<Object Desc |<Module Code "
        .ColWidth(1) = 1400
        .ColWidth(2) = 2600
        .ColWidth(3) = 1100
        .Redraw = True
    End With
    
End Sub

Private Sub txtObjectCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Trim(txtObjectCode) <> "" Then
 lb_found = MySeek(txtObjectCode.Text, "ObjectCode", PR_Objects)
 If Not lb_found Then
            Call SetErr(Gs_RecNFMsg, vbCritical)
            'Cancel = True
            txtObjectCode = ""
            txtObjectDesc = ""
            txtObjectCode.SetFocus
 Else
  txtObjectDesc = Trim("" & PR_Objects("ObjectDesc"))
  cmdAddtoGrid_Click
  txtObjectCode.SetFocus
 End If
End If


End Sub



Public Sub FrmRefresh()

    PR_Groups.Requery
    PR_Modules.Requery
    PR_Objects.Requery

End Sub

Public Sub SaveValues()
On Error GoTo LocalErr
Dim cntsql As New ADODB.Command
Dim ln_Index As Integer
Dim ln_GroupCode As Integer

PB_Blnk = False
cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText
gc_dbcon.BeginTrans

    Select Case Mode
    
        Case "A"
            ln_GroupCode = SerMaxNo(gc_dbcon, "Sys_Groups", "GroupCode")
        
            cntsql.CommandText = "INSERT INTO Sys_Groups(GROUPCODE, GROUPDESC, ADDDATETIME) VALUES (" & Val(ln_GroupCode) & ",'" & UCase(RepApp(Trim(txtGroupDesc))) & "','" & Format(Now, "YYYY/MM/DD hh:mm:ss AMPM") & "')"
            cntsql.Execute
            
            For ln_Index = 1 To grdGroup.Rows - 2
                gc_dbcon.Execute "INSERT into Sys_UserRights(GroupOrUserCode, ObjectCode, GroupFlag, ModuleCode, AddDateTime) VALUES(" & ln_GroupCode & ",'" & RepApp(Trim(grdGroup.TextMatrix(ln_Index, 1))) & "',1,'" & RepApp(Trim(grdGroup.TextMatrix(ln_Index, 3))) & "','" & Format(Now, "YYYY/MM/DD hh:mm:ss AMPM") & "')"
            Next
            
            txtGroupCode = SerMaxNo(gc_dbcon, "Sys_Groups", "GroupCode")
        Case "E"
            ln_GroupCode = Val(txtGroupCode)
        
            cntsql.CommandText = "UPDATE Sys_Groups Set GROUPDESC = '" & UCase(RepApp(Trim(txtGroupDesc.Text))) & "', MODIFYDATETIME = '" & Format(Now, "YYYY/MM/DD HH:mm:ss") & "' WHERE GROUPCODE = " & Val(txtGroupCode) & ""
            cntsql.Execute
            
            cntsql.CommandText = "DELETE Sys_UserRights Where GroupOrUserCode = " & ln_GroupCode & ""
            cntsql.Execute
            
            For ln_Index = 1 To grdGroup.Rows - 2
                gc_dbcon.Execute "INSERT into Sys_UserRights(GroupOrUserCode, ObjectCode, GroupFlag, ModuleCode, AddDateTime) VALUES(" & ln_GroupCode & ",'" & RepApp(Trim(grdGroup.TextMatrix(ln_Index, 1))) & "',1,'" & RepApp(Trim(grdGroup.TextMatrix(ln_Index, 3))) & "','" & Format(Now, "YYYY/MM/DD HH:mm:ss") & "')"
            Next
            
        Case "D"
            ln_GroupCode = Val(txtGroupCode)
            
            cntsql.CommandText = "DELETE Sys_UserRights Where GroupOrUserCode = " & ln_GroupCode & ""
            cntsql.Execute
            
            cntsql.CommandText = "DELETE FROM Sys_Groups WHERE GROUPCODE = " & Val(txtGroupCode) & ""
            cntsql.Execute
            
    End Select
    
gc_dbcon.CommitTrans
PR_Groups.Requery
Call InitializeGrid

Exit Sub

LocalErr:
gc_dbcon.RollbackTrans
MsgBox Err.Description
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0

End Sub

Public Function ChkInputs() As Boolean
    If Trim(txtGroupCode) <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function


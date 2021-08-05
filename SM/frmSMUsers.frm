VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSMUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Maintenance"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "frmSMUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   5490
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameUserInformation 
      Caption         =   "User Information"
      ForeColor       =   &H00400000&
      Height          =   1995
      Left            =   60
      TabIndex        =   6
      Top             =   570
      Width           =   5370
      Begin VB.CheckBox chkcosmeticsystem 
         Caption         =   "Cosmetics System"
         Height          =   315
         Left            =   3390
         TabIndex        =   27
         Top             =   1575
         Width           =   1875
      End
      Begin VB.CheckBox chkmedsystem 
         Caption         =   "Medicine System"
         Height          =   315
         Left            =   3390
         TabIndex        =   26
         Top             =   1260
         Width           =   1875
      End
      Begin VB.CheckBox chkposaccess 
         Caption         =   "POS Access only"
         Height          =   315
         Left            =   3375
         TabIndex        =   25
         Top             =   960
         Width           =   1875
      End
      Begin VB.CommandButton cmdUserLookup 
         Height          =   315
         Left            =   3210
         Picture         =   "frmSMUsers.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   195
         Width           =   315
      End
      Begin VB.TextBox txtConfirmPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1755
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1425
         Width           =   1485
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1755
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1035
         Width           =   1485
      End
      Begin VB.TextBox txtFullName 
         Height          =   315
         Left            =   1755
         MaxLength       =   50
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   615
         Width           =   3480
      End
      Begin VB.TextBox txtUsername 
         BackColor       =   &H00FFFF80&
         Height          =   315
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   210
         Width           =   1485
      End
      Begin VB.Label lblConfirmPassword 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password :"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1470
         Width           =   1350
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         Height          =   195
         Left            =   825
         TabIndex        =   9
         Top             =   1050
         Width           =   780
      End
      Begin VB.Label lblFullName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name :"
         Height          =   195
         Left            =   810
         TabIndex        =   8
         Top             =   645
         Width           =   795
      End
      Begin VB.Label lblUsername 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username/User Id :"
         Height          =   195
         Left            =   210
         TabIndex        =   7
         Top             =   255
         Width           =   1395
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5490
      _ExtentX        =   9684
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
               Picture         =   "frmSMUsers.frx":047C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMUsers.frx":08D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMUsers.frx":0D24
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMUsers.frx":1178
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMUsers.frx":15CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMUsers.frx":1A20
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSMUsers.frx":2174
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3645
      Left            =   60
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2685
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   6429
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   4194304
      TabCaption(0)   =   "Member Of &Group (1)"
      TabPicture(0)   =   "frmSMUsers.frx":25C8
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frameGroup"
      Tab(0).Control(1)=   "grdGroup"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Individual &Rights (2)"
      TabPicture(1)   =   "frmSMUsers.frx":25E4
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "grdRight"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtsiteid"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.ComboBox txtsiteid 
         Height          =   315
         ItemData        =   "frmSMUsers.frx":2600
         Left            =   2370
         List            =   "frmSMUsers.frx":260A
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   3060
         Width           =   1785
      End
      Begin VB.Frame frameGroup 
         Height          =   570
         Left            =   -74880
         TabIndex        =   19
         Top             =   360
         Width           =   5115
         Begin VB.TextBox txtGroupCode 
            Height          =   300
            Left            =   615
            MaxLength       =   3
            TabIndex        =   24
            Top             =   165
            Width           =   675
         End
         Begin VB.TextBox txtGroupDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   150
            Width           =   3375
         End
         Begin VB.CommandButton cmdGroupLookup 
            Height          =   315
            Left            =   1320
            Picture         =   "frmSMUsers.frx":2620
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   150
            Width           =   315
         End
         Begin VB.Label lblGroup 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Group :"
            Height          =   195
            Left            =   75
            TabIndex        =   22
            Top             =   195
            Width           =   525
         End
      End
      Begin VB.Frame Frame1 
         Height          =   570
         Left            =   120
         TabIndex        =   13
         Top             =   375
         Width           =   5115
         Begin VB.TextBox txtObjectCode 
            Height          =   315
            Left            =   555
            MaxLength       =   10
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   165
            Width           =   1275
         End
         Begin VB.CommandButton cmdObjectLookup 
            Height          =   315
            Left            =   1875
            Picture         =   "frmSMUsers.frx":2792
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   165
            Width           =   315
         End
         Begin VB.TextBox txtObjectDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   300
            Left            =   2220
            MaxLength       =   50
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   165
            Width           =   2040
         End
         Begin VB.ComboBox cmbStatus 
            Height          =   315
            ItemData        =   "frmSMUsers.frx":2904
            Left            =   4290
            List            =   "frmSMUsers.frx":290E
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   150
            Width           =   750
         End
         Begin VB.Label lblRight 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Right :"
            Height          =   195
            Left            =   60
            TabIndex        =   17
            Top             =   195
            Width           =   465
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdGroup 
         Height          =   1995
         Left            =   -74895
         TabIndex        =   18
         Top             =   945
         Width           =   5130
         _ExtentX        =   9049
         _ExtentY        =   3519
         _Version        =   393216
         Rows            =   1
      End
      Begin MSFlexGridLib.MSFlexGrid grdRight 
         Height          =   1980
         Left            =   105
         TabIndex        =   23
         Top             =   960
         Width           =   5130
         _ExtentX        =   9049
         _ExtentY        =   3493
         _Version        =   393216
         Rows            =   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Site ID For Purchase  :"
         Height          =   195
         Left            =   165
         TabIndex        =   28
         Top             =   3090
         Width           =   2160
      End
   End
End
Attribute VB_Name = "frmSMUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_Blnk As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Users As New Recordset
Dim PR_Groups As New Recordset
Dim PR_Objects As New Recordset
Dim PR_UserGroups As New Recordset
Dim PR_UserRights As New Recordset

Dim PI_CurRow    As Integer
Dim PI_SrNo      As Integer
Dim PS_RowClicked As String

Dim PI_CurRow1    As Integer
Dim PI_SrNo1      As Integer
Dim PS_RowClicked1 As String

Dim ln_systemaccess As Integer

Dim lb_found As Boolean
Public Sub InitializeGrid1()
   With grdRight
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Object Code |<Object Desc |<Module |<Status "
        .ColWidth(1) = 1200
        .ColWidth(2) = 2600
        .ColWidth(3) = 600
        .ColWidth(4) = 600
        .Redraw = True
    End With
    PI_SrNo1 = 0
    PI_CurRow1 = 0
    PS_RowClicked1 = ""
End Sub
Public Sub InitializeGrid()
    With grdGroup
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Group Code |<Group Desc "
        .ColWidth(1) = 1000
        .ColWidth(2) = 3300
        .Redraw = True
    End With
    PI_SrNo = 0
    PI_CurRow = 0
    PS_RowClicked = ""

End Sub

Private Sub chkposaccess_Click()
If chkposaccess.Value = 1 Then
chkcosmeticsystem.Value = 0
chkmedsystem.Value = 0
End If
End Sub
Private Sub chkmedsystem_Click()
If chkmedsystem.Value = 1 Then
chkcosmeticsystem.Value = 0
chkposaccess.Value = 0
End If
End Sub
Private Sub chkcosmeticsystem_Click()
If chkcosmeticsystem.Value = 1 Then
chkposaccess.Value = 0
chkmedsystem.Value = 0
End If
End Sub

Private Sub cmbStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Trim(txtObjectCode) <> "" Then
     AddToGrid1
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
End Sub

Private Sub cmdObjectLookup_Click()

    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtObjectCode
    Set PO_DESC = txtObjectDesc
    
    Gs_OtherPara = ""
    Gs_ExtraPara = ""
    
    Gs_SQL = "Select ObjectCode as Code,ObjectDesc as Description from Sys_Objects "
    Gs_FindFld = "ObjectDesc"
    Gs_OrderBy = "order by ObjectDesc"
    MyLookupOLDB.Caption = "Objects - " & App.ProductName
    MyLookupOLDB.Show 1
    
   ' If txtObjectCode <> "" Then Call txtObjectCode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub cmdUserLookup_Click()

    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtUsername
    Set PO_DESC = txtFullName
    
    Gs_SQL = "Select UserId as Code,UserName as Description from SyUsers "
    Gs_FindFld = "UserName"
    Gs_OrderBy = "order by UserName"
    MyLookupOLDB.Caption = "Users - " & App.ProductName
    MyLookupOLDB.Show 1
    If txtUsername <> "" Then Call txtusername_KeyDown(vbKeyReturn, vbKeyShift)
    
   

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF11 Then
 '  '     Mode = DentMode(Mode, 4, PR_Users, Me, txtUsername, txtFullName, "X", "CompCount", 3, "Acct_sub0", "Acct_Desc", 1, False, Toolbar1)
 '  End If
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

    PR_Users.Open "Select * from SyUsers", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
    PR_Groups.Open "Select * from Sys_Groups", gc_dbcon, adOpenStatic, adLockReadOnly, 1
    PR_Objects.Open "Select * from Sys_Objects", gc_dbcon, adOpenStatic, adLockReadOnly, 1
    PR_UserRights.Open "Select * from Sys_UserRights Where GroupFlag = 0", gc_dbcon, adOpenStatic, adLockReadOnly, 1
    PR_UserGroups.Open "Select * from Sys_UserGroups", gc_dbcon, adOpenStatic, adLockReadOnly, 1
    
    Call InitializeGrid
    Call InitializeGrid1
    cmbStatus.ListIndex = 0
    TxtSiteID.ListIndex = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

    PR_Objects.Close
    PR_Groups.Close
    PR_Users.Close
    PR_UserRights.Close
    PR_UserGroups.Close

End Sub

Private Sub grdGroup_DblClick()
    If grdGroup.Row = grdGroup.Rows - 1 Then Exit Sub
    With grdGroup
        txtGroupCode = .TextMatrix(.Row, 1)
    End With
End Sub

Private Sub grdGroup_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If (grdGroup.Row = 1 Or grdGroup.Row = 0) And grdGroup.Rows = 2 Then Exit Sub
        If grdGroup.Row + 1 = grdGroup.Rows Then Exit Sub
        grdGroup.RemoveItem grdGroup.Row
    End If
End Sub

Private Sub grdRight_DblClick()
    If grdRight.Row = grdRight.Rows - 1 Then Exit Sub
    With grdRight
        txtObjectCode = .TextMatrix(.Row, 1)
        cmbStatus.ListIndex = .TextMatrix(.Row, 4)
    End With
End Sub

Private Sub grdRight_KeyDown(KeyCode As Integer, Shift As Integer)
    
   If KeyCode = vbKeyDelete Then 'Delete Key Pressed
    With grdRight
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            
            .RemoveItem .Row
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                InitializeGrid1
            End If
     End With
   End If
    
   End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If PB_Blnk And Range(Button.Index, 2, 3) Then
    
        MsgBox "Data not found :", vbCritical
        Mode = ""
        'Cancel = True
        
    Else
        
        Mode = DentMode(Mode, Button.Index, PR_Users, Me, txtUsername, txtFullName, "X", "CompCount", 3, "Acct_sub0", "Acct_Desc", 1, False, Toolbar1)
        
       ' If Mode = "A" Then
        '     cmdUserLookup.Enabled = False
             Call InitializeGrid
             Call InitializeGrid1
        'Else
        '     cmdUserLookup.Enabled = True
        'End If
        
    End If

End Sub

Private Sub txtConfirmPassword_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtObjectCode.SetFocus
End Sub

Private Sub txtFullName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtPassword.SetFocus
End Sub

Private Sub txtFullName_LostFocus()
If txtFullName <> "" Then
txtFullName = StrConv(txtFullName, vbProperCase)
End If
End Sub

Private Sub txtGroupCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn And Trim(txtGroupCode) <> "" Then
       Dim lb_found As Boolean
       lb_found = MySeek(txtGroupCode.Text, "GroupCode", PR_Groups)
       If Not lb_found Then
         Call MsgBox("Group code not found!!!", vbCritical)
         txtGroupCode = ""
         txtGroupDesc = ""
         txtGroupCode.SetFocus
        Else
            txtGroupDesc = Trim("" & PR_Groups("GroupDesc"))
            AddToGrid
            txtGroupCode.SetFocus
        End If
        
    Else
        txtGroupCode = ""
        txtGroupDesc = ""
    End If
End Sub
Private Sub txtObjectCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtObjectCode <> "" Then

lb_found = MySeek(txtObjectCode.Text, "ObjectCode", PR_Objects)
    If Not lb_found Then
        Call MsgBox("Object code not found!!!", vbCritical)
        txtObjectCode.SetFocus
    Else
          txtObjectDesc.Text = Trim(PR_Objects("ObjectDesc") & "")
         Call AddToGrid1
      
         txtObjectCode.SetFocus
    End If
End If
End Sub

Private Sub txtpassword_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtConfirmPassword.SetFocus
End Sub


Private Sub txtusername_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtUsername <> "" Then
   
    'txtUsername = Trim(txtUsername)
    lb_found = MySeek(txtUsername.Text, "Userid", PR_Users)
        Select Case Mode
            Case "A"
                If lb_found Then
                    Call MsgBox("User ID already Exist !!!", vbCritical)
                    txtUsername = ""
                    txtFullName = ""
                    txtUsername.SetFocus
                Else
                    txtFullName.SetFocus
                End If
                
            Case Else
                If Not lb_found Then
                    Call MsgBox("Record not found !!! ", vbCritical)
                    txtUsername = ""
                    txtFullName = ""
                    txtUsername.SetFocus
                Else
                    
                    Call InitializeGrid
                    Call InitializeGrid1
                    Call SetVal
                    If txtFullName.Enabled Then txtFullName.SetFocus
                End If
            End Select
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
                                       .TextMatrix(.Row, 1) = txtGroupCode
                                       .TextMatrix(.Row, 2) = txtGroupDesc
                           
                        End With
                        txtGroupCode = ""
                        txtGroupDesc = ""
                        txtGroupCode.SetFocus
                   
        Else
            Call MsgBox("Enter Group Code !!!", vbCritical)
            txtGroupCode.SetFocus
        End If
      

End Sub
Private Sub AddToGrid1()
Dim ln_cnt As Integer
         If txtObjectCode.Text <> "" Then
                    If PS_RowClicked1 = "" Then
                        If PI_SrNo1 = 0 Then
                            PI_SrNo1 = 1
                        Else
                            PI_SrNo1 = PI_SrNo1 + 1
                         End If
                     End If
        
                        With grdRight
                            If PS_RowClicked1 = "" Then
                                    If Not PI_SrNo1 = 1 Then .Rows = .Rows + 1
                                        .Row = .Rows - 1
                                    Else
                                        .Row = PI_CurRow1
                                    End If
                                    
                                    If PS_RowClicked1 = "" Then
                                        .TextMatrix(.Row, 0) = PI_SrNo1
                                    Else
                                        .TextMatrix(.Row, 0) = PI_CurRow1
                                    End If
                                       .TextMatrix(.Row, 1) = txtObjectCode
                                       .TextMatrix(.Row, 2) = txtObjectDesc
                                       .TextMatrix(.Row, 3) = PR_Objects("ModuleCode")
                                       .TextMatrix(.Row, 4) = cmbStatus.ListIndex
                           
                        End With
                       
                        txtObjectCode = ""
                        txtObjectDesc = ""
                        SSTab1.Tab = 1
                        txtObjectCode.SetFocus
                         
                   
        Else
            Call MsgBox("Enter Object Code !!!", vbCritical)
            txtObjectCode.SetFocus
        End If
      

End Sub


Public Sub SetVal()
    On Error GoTo LocalErr
    txtFullName = Trim("" & PR_Users("Username"))
    
    If Val(0 & PR_Users("Posaccess")) = 1 Then
    chkposaccess.Value = 1
    End If
    
    If Val(0 & PR_Users("Posaccess")) = 2 Then
    chkmedsystem.Value = 1
    End If
    
    If Val(0 & PR_Users("Posaccess")) = 3 Then
    chkcosmeticsystem.Value = 1
    End If
    
    txtPassword = LTrim(RTrim(DeCode(PR_Users("Password"), 50))) & ""
    txtConfirmPassword = LTrim(RTrim(DeCode(PR_Users("Password"), 50))) & ""
    LoadTrans
    LoadTrans1
    TxtSiteID.Text = Trim(PR_Users("Siteid"))
Exit Sub
LocalErr:
    
End Sub
Private Sub LoadTrans()
InitializeGrid
PR_UserGroups.Filter = "UserCode = " & PR_Users("UserCode")
        

If Not PR_UserGroups.EOF Then
        With grdGroup
            Do While Not PR_UserGroups.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = PR_UserGroups("GroupCode")
                 If MySeek(Trim(PR_UserGroups("GroupCode")), "GroupCode", PR_Groups) Then
                    .TextMatrix(.Row, 2) = PR_Groups("GroupDesc")
                 End If
                .Rows = .Rows + 1
                PR_UserGroups.MoveNext
                If PR_UserGroups.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
        
        
End If
PR_UserGroups.Filter = adFilterNone
    
Exit Sub
LocalErr:
Call MsgBox(Err.Description)
End Sub

Private Sub LoadTrans1()
InitializeGrid1

PR_UserRights.Filter = "grouporusercode = " & PR_Users("UserCode")

If Not PR_UserRights.EOF Then
        With grdRight
            Do While Not PR_UserRights.EOF
                .Row = .Rows - 1
               
                 PI_SrNo1 = .Row
                 .TextMatrix(.Row, 0) = .Row
                .TextMatrix(.Row, 1) = PR_UserRights("ObjectCode")
                If MySeek(Trim(PR_UserRights("ObjectCode")), "ObjectCode", PR_Objects) Then
                .TextMatrix(.Row, 2) = PR_Objects("ObjectDesc")
                End If
                .TextMatrix(.Row, 3) = PR_UserRights("ModuleCode")
                .TextMatrix(.Row, 4) = PR_UserRights("Status")
                .Rows = .Rows + 1
                PR_UserRights.MoveNext
                If PR_UserRights.EOF Then Exit Do
             Loop
               If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
        
        
End If
PR_UserRights.Filter = adFilterNone
    
Exit Sub
LocalErr:
Call MsgBox(Err.Description)
End Sub


Public Sub SaveValues()
'On Error GoTo LocalErr
Dim cntsql As New ADODB.Command
Dim ln_usercode As Integer
Dim ln_cnt As Integer
ln_systemaccess = 0

If chkposaccess.Value = 1 Then
ln_systemaccess = 1
End If


If chkmedsystem.Value = 1 Then
ln_systemaccess = 2
End If


If chkcosmeticsystem.Value = 1 Then
ln_systemaccess = 3
End If


PB_Blnk = False
cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText
gc_dbcon.BeginTrans
Ls_Password = EnCode(txtPassword.Text, 50)
    Select Case Mode
        Case "A"
            ln_usercode = SerMaxNo(gc_dbcon, "SyUsers", "UserCode")
            gc_dbcon.Execute "INSERT into SyUsers(compcode,usercode,userid,UserName,password,POSAccess,Siteid) VALUES ('" & Gs_compcode & "'," & ln_usercode & ",'" & txtUsername.Text & "','" & txtFullName.Text & "','" & Ls_Password & "'," & ln_systemaccess & ",'" & TxtSiteID.Text & "')"
            
            With grdGroup
                For ln_cnt = 1 To .Rows - 1
                    If Trim(.TextMatrix(1, 1)) <> "" Then
                        gc_dbcon.Execute "INSERT INTO SYS_USERGROUPS(USERCODE, GROUPCODE, ADDDATETIME) VALUES(" & ln_usercode & "," & Val(.TextMatrix(ln_cnt, 1)) & ",'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "')"
                    End If
                Next
            End With
            
            With grdRight
                For ln_cnt = 1 To .Rows - 1
                    If Trim(.TextMatrix(1, 1)) <> "" Then
                        gc_dbcon.Execute "INSERT INTO SYS_USERRIGHTS(GROUPORUSERCODE, OBJECTCODE, GROUPFLAG, MODULECODE, STATUS, ADDDATETIME) VALUES(" & ln_usercode & ",'" & .TextMatrix(ln_cnt, 1) & "',0,'" & .TextMatrix(ln_cnt, 3) & "'," & Val(.TextMatrix(ln_cnt, 4)) & ",'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "')"
                    End If
                Next
            End With
         
            Case "E"
            ln_usercode = PR_Users("UserCode")
            gc_dbcon.Execute "UPDATE syusers SET userid = '" & txtUsername & "', password ='" & Ls_Password & "',UserName ='" & txtFullName & "',Posaccess =" & ln_systemaccess & ",Siteid = '" & TxtSiteID.Text & "' WHERE  compcode = '" & Gs_compcode & "' and usercode= " & Trim(ln_usercode) & ""
            
            gc_dbcon.Execute "Delete from Sys_UserGroups WHERE Usercode = " & PR_Users("UserCode")
            gc_dbcon.Execute "Delete from Sys_UserRights WHERE GroupOrUserCode = " & PR_Users("UserCode") & " "
           
            With grdGroup
                For ln_cnt = 1 To .Rows - 1
                   If Trim(.TextMatrix(1, 1)) <> "" Then
                    gc_dbcon.Execute "INSERT INTO SYS_USERGROUPS(USERCODE, GROUPCODE, ADDDATETIME) VALUES(" & ln_usercode & "," & Val(.TextMatrix(ln_cnt, 1)) & ",'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "')"
                   End If
                Next
            End With
            With grdRight
               For ln_cnt = 1 To .Rows - 1
                If Trim(.TextMatrix(1, 1)) <> "" Then
                    gc_dbcon.Execute "INSERT INTO SYS_USERRIGHTS(GROUPORUSERCODE, OBJECTCODE, GROUPFLAG, MODULECODE, STATUS, ADDDATETIME) VALUES(" & ln_usercode & ",'" & .TextMatrix(ln_cnt, 1) & "',0,'" & .TextMatrix(ln_cnt, 3) & "'," & Val(.TextMatrix(ln_cnt, 4)) & ",'" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "')"
                End If
               Next
            End With

           Case "D"
            ln_usercode = PR_Users("UserCode")

            gc_dbcon.Execute "Delete from Sys_UserGroups WHERE Usercode = " & PR_Users("UserCode")
            gc_dbcon.Execute "Delete from Sys_UserRights WHERE GROUPORUSERCODE = " & PR_Users("UserCode")
            gc_dbcon.Execute "DELETE FROM SyUsers WHERE Userid = '" & txtUsername & "'"

           End Select
gc_dbcon.CommitTrans
PR_Users.Requery
PR_UserGroups.Requery
PR_UserRights.Requery
Call InitializeGrid

Exit Sub

LocalErr:
gc_dbcon.RollbackTrans
MsgBox Err.Description
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0

End Sub

Public Function ChkInputs() As Boolean

    If Trim(txtUsername) = "" Then
        MsgBox "Enter Username!!!", vbInformation, App.ProductName
        txtUsername.SetFocus
        ChkInputs = False
        Exit Function
    End If
    
    If Trim(txtFullName) = "" Then
        MsgBox "Enter User Full Name!!!", vbInformation, App.ProductName
        txtFullName.SetFocus
        ChkInputs = False
        Exit Function
    End If
    
    If Trim(txtPassword) = "" Then
        MsgBox "Enter Password!!!", vbInformation, App.ProductName
        txtPassword.SetFocus
        ChkInputs = False
        Exit Function
    End If
    
    If Trim(txtConfirmPassword) = "" Then
        MsgBox "Enter Confirm Password!!!", vbInformation, App.ProductName
        txtConfirmPassword.SetFocus
        ChkInputs = False
        Exit Function
    End If
    
    If Trim(txtPassword) <> Trim(txtConfirmPassword) Then
        MsgBox "Passwords Do Not Match!!!", vbInformation, App.ProductName
        txtPassword.SetFocus
        ChkInputs = False
        Exit Function
    End If
    
    ChkInputs = True

End Function

Private Sub txtUsername_LostFocus()
If txtUsername <> "" Then
Call txtusername_KeyDown(vbKeyReturn, vbKeyShift)
End If
End Sub

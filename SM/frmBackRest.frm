VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBackRest 
   Caption         =   "Backup Database"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   15
      TabIndex        =   0
      Top             =   -60
      Width           =   8025
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   4110
         MaxLength       =   50
         TabIndex        =   25
         Top             =   2925
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   7065
         Picture         =   "frmBackRest.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1425
         Width           =   315
      End
      Begin VB.TextBox cmbDatabaseName 
         Height          =   315
         Left            =   4560
         TabIndex        =   21
         Text            =   "Datatransfer"
         Top             =   675
         Width           =   3375
      End
      Begin VB.CommandButton cmdBackup 
         Caption         =   "Backup"
         Height          =   390
         Left            =   5805
         TabIndex        =   17
         Top             =   3210
         Width           =   1215
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "Restore"
         Height          =   390
         Left            =   5580
         TabIndex        =   16
         Top             =   3210
         Width           =   1215
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   390
         Left            =   765
         TabIndex        =   15
         Top             =   3180
         Width           =   1215
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "Disconnect"
         Height          =   390
         Left            =   2025
         TabIndex        =   14
         Top             =   3180
         Width           =   1215
      End
      Begin VB.TextBox txtStatus 
         Height          =   2070
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   13
         Top             =   3630
         Width           =   7815
      End
      Begin VB.Frame frmAuthorization 
         Caption         =   "Authorization:"
         Height          =   1815
         Left            =   120
         TabIndex        =   6
         Top             =   1140
         Width           =   3855
         Begin VB.TextBox txtUserName 
            Height          =   288
            Left            =   1320
            TabIndex        =   10
            Text            =   "sa"
            Top             =   1080
            Width           =   2292
         End
         Begin VB.TextBox txtPassword 
            Height          =   288
            IMEMode         =   3  'DISABLE
            Left            =   1320
            PasswordChar    =   "*"
            TabIndex        =   9
            Top             =   1440
            Width           =   2292
         End
         Begin VB.OptionButton optWinNTAuth 
            Caption         =   "Use Windows NT authentication"
            Height          =   252
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   2892
         End
         Begin VB.OptionButton optSSAuth 
            Caption         =   "Use SQL Server authentication"
            Height          =   252
            Left            =   240
            TabIndex        =   7
            Top             =   720
            Value           =   -1  'True
            Width           =   3252
         End
         Begin VB.Label lblUserName 
            Caption         =   "Login name:"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblPassword 
            Caption         =   "Password:"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   1440
            Width           =   975
         End
      End
      Begin VB.Frame frmConnectionInfo 
         Caption         =   "Connection:"
         Height          =   855
         Left            =   120
         TabIndex        =   3
         Top             =   180
         Width           =   3855
         Begin VB.TextBox txtServerName 
            Height          =   288
            Left            =   1320
            TabIndex        =   4
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lblServer 
            Caption         =   "Server:"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.TextBox txtDataFileName 
         Height          =   288
         Left            =   4560
         TabIndex        =   2
         Top             =   2340
         Width           =   3375
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Height          =   405
         Left            =   6960
         TabIndex        =   1
         Top             =   1800
         Width           =   975
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3870
         Top             =   180
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Data File Name:"
      End
      Begin MSMask.MaskEdBox txtbranchcode 
         Height          =   315
         Left            =   6525
         TabIndex        =   23
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         ToolTipText     =   "Default Currency"
         Top             =   1440
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin VB.Label Label18 
         Caption         =   "Branch # :"
         Height          =   255
         Left            =   5730
         TabIndex        =   24
         Top             =   1470
         Width           =   765
      End
      Begin VB.Label lblQueryResults 
         Caption         =   "Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3915
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Database To Backup/Restore:"
         Height          =   255
         Left            =   4560
         TabIndex        =   19
         Top             =   420
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Backup/Restore File Name:"
         Height          =   255
         Left            =   4560
         TabIndex        =   18
         Top             =   1860
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmBackRest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Backup Restore with Events Sample Application
' Microsoft Copyright 2000
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Dim gSQLServer As SQLDMO.SQLServer

Dim WithEvents oBackupEvent As SQLDMO.Backup
Attribute oBackupEvent.VB_VarHelpID = -1
Dim WithEvents oRestoreEvent As SQLDMO.Restore
Attribute oRestoreEvent.VB_VarHelpID = -1

Dim gbConnected As Boolean
Dim gDatabaseName As String
Dim gBkupRstrFileName As String
Dim gBkupRstrFilePath As String
Dim PR_Branch As New Recordset
Public PO_CODE As Object
Public PO_DESC As Object



Const gTitle = "Server Connection"
Private Sub Command4_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbranchcode
    Set PO_DESC = Text1
    GoTop PR_Branch
    MyLookup.Caption = "Branches"
    MyLookup.FillGrid PR_Branch, "BranchCode", "BranchDesc", txtbranchcode.MaxLength
    MyLookup.Show 1
    
    If Len(txtbranchcode) > 0 Then txtBranchCode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Form_Load()
On Error GoTo LocalErr
Dim ls_servername As New Connection
Dim ls_rsetname As New Recordset
If ls_servername.State <> 1 Then ls_servername.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Gs_AppPath & " ;Persist Security Info=False"
    ls_rsetname.Open "select * from servernametab", ls_servername, adOpenStatic, adLockReadOnly, adCmdText
    If Not ls_rsetname.EOF Then
        txtServerName = Trim(ls_rsetname("servername") + "")
        'cmbDatabaseName = Trim(ls_rsetname("DBFname") + "")
        ls_rsetname.Close
        ls_servername.Close
    End If
    Set gSQLServer = Nothing
    'optWinNTAuth.Value = True
    gbConnected = False
    'WinNTAuthOptionsOn
    buttonsConnectClosed
    PR_Branch.Open "Select * from SysBranch Where Compcode = '" & Gs_compcode & "' order by Branchcode", gc_dbcon, adOpenStatic, adLockReadOnly, 1
    txtbranchcode = Gs_BranchCode
    Exit Sub
LocalErr:

MsgBox Err.Description
    Call SetErr("Processing error call system administrator", vbCritical)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gbConnected = True Then
        Call gSQLServer.Disconnect
    End If
    If Not gSQLServer Is Nothing Then
        Set gSQLServer = Nothing
    End If
    PR_Branch.Close
End Sub



Private Sub cmdConnect_Click()
    Dim ServerName As String
    Dim UserName As String
    Dim Password As String

    On Error GoTo ErrHandler:

    If gSQLServer Is Nothing Then
        Set gSQLServer = New SQLDMO.SQLServer
    End If
    
    ' Put text box values into connection variables.
    ServerName = txtServerName.Text
    UserName = txtUserName.Text
    Password = txtPassword.Text
     
    ' Set the login timeout.
    gSQLServer.LoginTimeout = 15
    
    ' Decision code for login authorization type: WinNT or SQL Server.
    If optWinNTAuth.Value = True Then
        gSQLServer.LoginSecure = True
    End If
    
    ' Change mousepointer while trying to connect.
    Screen.MousePointer = vbHourglass
    
    gSQLServer.Connect ServerName, UserName, Password
    
    gbConnected = True
    
    ' List all of the database names.
    'FillDatabaseList
    
    ' Change mousepointer back to the default after connect.
    Screen.MousePointer = vbDefault
    
    ' Notify user that connection was successful.
    'MsgBox "Connection to server successful.", vbOKOnly, gTitle
    
    buttonsConnectOpen
    
    ' Clear up the status text in the "result field".
    txtStatus.Text = ""
    
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Description
    
    ' Change mousepointer back if it's hourglass.
    If Screen.MousePointer = vbHourglass Then
        Screen.MousePointer = vbDefault
    End If
    
End Sub

Private Sub cmdDisconnect_Click()
    On Error GoTo ErrHandler:
    
    Dim Msg As String
    Dim Response As String

    ' Disconnect from a connected server.
    If gbConnected = True Then
        Msg = "Disconnect from Server?"
        Response = MsgBox(Msg, vbOKCancel, gTitle)
        If Response = vbOK Then
            Call gSQLServer.Disconnect
            Set gSQLServer = Nothing
            cmbDatabaseName.Clear
            txtDataFileName.Text = ""
            txtStatus.Text = ""
            gbConnected = False
            buttonsConnectClosed
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error " & Err.Description
    Resume Next
End Sub

Private Sub cmdBackup_Click()
    On Error GoTo ErrHandler:
    Call backuppath
    Call takebackup
    Call cmdConnect_Click
    Dim oBackup As SQLDMO.Backup
    
    gDatabaseName = cmbDatabaseName.Text
    Set oBackup = New SQLDMO.Backup
    Set oBackupEvent = oBackup ' enable events
    
    oBackup.Database = gDatabaseName
    gBkupRstrFileName = txtDataFileName.Text
    oBackup.Files = gBkupRstrFileName
    
    ' Delete the datafile to allow the application to create a brand new file.
    ' This will prevent attaching the new backup data to the old data if there
    ' is any.
    If Len(Dir(gBkupRstrFileName)) > 0 Then
        Kill (gBkupRstrFileName)
    End If
    
    ' Change mousepointer while trying to connect.
    Screen.MousePointer = vbHourglass
    
    ' Backup the database.
    oBackup.SQLBackup gSQLServer
    
    ' Change mousepointer back to the default after connect.
    Screen.MousePointer = vbDefault
   
    Set oBackupEvent = Nothing ' disable events
    Set oBackup = Nothing
    Call MsgBox("Backup Successfully", vbInformation)
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Description
    Resume Next
End Sub
Private Sub takebackup()
On Error Resume Next
Dim ls_sql As String
    ls_sql = "drop table datatransfer.dbo.gl_ref "
    gc_dbcon.Execute ls_sql
    ls_sql = "drop table datatransfer.dbo.gl_TRANS "
    gc_dbcon.Execute ls_sql
    ls_sql = "drop table datatransfer.dbo.iC_TRANS "
    gc_dbcon.Execute ls_sql
    ls_sql = "drop table datatransfer.dbo.iC_Item "
    gc_dbcon.Execute ls_sql
    
    
    ls_sql = "drop table datatransfer.dbo.gl_detail "
    gc_dbcon.Execute ls_sql
    
    'ls_sql = "select *  INTO datatransfer.dbo.gl_detail from ecounts.dbo.gl_detail where compcode = '" & Gs_compcode & "' and branchcode = '" & Gs_BranchCode & "' "
    'gc_dbcon.Execute ls_sql
    
    
    ls_sql = "select *  INTO datatransfer.dbo.gl_ref from ecounts.dbo.gl_ref where compcode = '" & Gs_compcode & "' and branchcode = '" & Gs_BranchCode & "' "
    gc_dbcon.Execute ls_sql
    
    
    ls_sql = "select *  INTO datatransfer.dbo.gl_trans from ecounts.dbo.gl_trans where compcode = '" & Gs_compcode & "' and branchcode = '" & Gs_BranchCode & "' "
    gc_dbcon.Execute ls_sql
    
    ls_sql = "select *  INTO datatransfer.dbo.ic_trans from ecounts.dbo.ic_trans where compcode = '" & Gs_compcode & "' and branchcode = '" & Gs_BranchCode & "' "
    gc_dbcon.Execute ls_sql
    
    ls_sql = "select *  INTO datatransfer.dbo.ic_item from ecounts.dbo.ic_item where compcode = '" & Gs_compcode & "' "
    gc_dbcon.Execute ls_sql
    
End Sub
Private Sub restorebackup()
On Error Resume Next
Dim ls_sql As String
    ls_sql = "delete from gl_ref where compcode = '" & Gs_compcode & "' and branchcode = '" & txtbranchcode & "' "
    gc_dbcon.Execute ls_sql
    
    
    'ls_sql = "delete from gl_detail where compcode = '" & Gs_compcode & "' and branchcode = '" & txtBranchCode & "' "
    'gc_dbcon.Execute ls_sql
    
    ls_sql = "delete from gl_trans where compcode = '" & Gs_compcode & "' and branchcode = '" & txtbranchcode & "' "
    gc_dbcon.Execute ls_sql
    
    ls_sql = "delete from ic_trans where compcode = '" & Gs_compcode & "' and branchcode = '" & txtbranchcode & "' "
    gc_dbcon.Execute ls_sql
        
    ls_sql = "delete from ic_item where compcode = '" & Gs_compcode & "' and ltrim(rtrim(locationcode)) = '" & Right(txtbranchcode, 2) & "' "
    gc_dbcon.Execute ls_sql

    ls_sql = "insert into ecounts.dbo.gl_detail  select * from datatransfer.dbo.gl_detail where compcode = '" & Gs_compcode & "' and branchcode = '" & txtbranchcode & "' "
    gc_dbcon.Execute ls_sql
    
    ls_sql = "insert into ecounts.dbo.gl_ref  select * from datatransfer.dbo.gl_ref where compcode = '" & Gs_compcode & "' and branchcode = '" & txtbranchcode & "' "
    gc_dbcon.Execute ls_sql
    ls_sql = "insert into ecounts.dbo.gl_trans select * from datatransfer.dbo.gl_trans  where compcode = '" & Gs_compcode & "' and branchcode = '" & txtbranchcode & "' "
    gc_dbcon.Execute ls_sql
    ls_sql = "insert into ecounts.dbo.ic_trans select * from datatransfer.dbo.ic_trans where compcode = '" & Gs_compcode & "' and branchcode = '" & txtbranchcode & "' "
    gc_dbcon.Execute ls_sql
    
    ls_sql = "insert into ecounts.dbo.ic_item select * from datatransfer.dbo.ic_item where compcode = '" & Gs_compcode & "' and ltrim(rtrim(locationcode)) = '" & Right(txtbranchcode, 2) & "' "
    gc_dbcon.Execute ls_sql
    
End Sub


Private Sub cmdRestore_Click()
    On Error GoTo ErrHandler:
    
    Call backuppath
    Call cmdConnect_Click
    
    Dim oRestore As SQLDMO.Restore
    
    Dim Msg As String
    Dim Response As String

'    Msg = "You must choose the right database name according to the data file name selected. Do you want to continue?"
'    Response = MsgBox(Msg, vbYesNo, gTitle)
'    If Response = vbNo Then
'        Exit Sub
'    End If
        
    gDatabaseName = cmbDatabaseName.Text
    Set oRestore = New SQLDMO.Restore
    Set oRestoreEvent = oRestore        ' enable events
    
    oRestore.Database = gDatabaseName
    gBkupRstrFileName = txtDataFileName.Text
    oRestore.Files = gBkupRstrFileName
    
    ' Change mousepointer while trying to connect.
    Screen.MousePointer = vbHourglass
    
    ' Restore the database.
    oRestore.SQLRestore gSQLServer
    
    ' Change mousepointer back to the default after connect.
    Screen.MousePointer = vbDefault
   
    Set oRestoreEvent = Nothing         ' disable events
    Set oRestore = Nothing
    
    Call restorebackup
    Call MsgBox("Restore Successfully", vbInformation)
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Description
    Resume Next
End Sub

Private Sub cmdBrowse_Click()
    On Error GoTo ErrHandler:
    
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "All Files (*.*)|*.*|Backup Files (*.bak)|*.bak"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.InitDir = gBkupRstrFilePath
    CommonDialog1.DefaultExt = "bak"
    CommonDialog1.DialogTitle = "Data File Name:"
    CommonDialog1.Action = 1
    txtDataFileName.Text = CommonDialog1.FileName
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub
Private Sub backuppath()
   Dim pr_dump As New Recordset
    pr_dump.Open " Select * from sysbackuppath ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
    If Not pr_dump.EOF Then
            txtDataFileName = pr_dump("Backuppath") & txtbranchcode
    End If
    pr_dump.Close
End Sub


' VB will create the right prototypes for you, if you select the oBackupEvent in
' the drop down listbox of your editor
Private Sub oBackupEvent_Complete(ByVal Message As String)
    PrintStat "oBackupEvent_Complete -- " & Message
End Sub

Private Sub oBackupEvent_NextMedia(ByVal Message As String)
    PrintStat "oBackupEvent_NextMedia -- " & Message
End Sub

Private Sub oBackupEvent_PercentComplete(ByVal Message As String, ByVal Percent As Long)
    PrintStat "oBackupEvent_PercentComplete -- " & Message & " " & Percent
End Sub

Private Sub oRestoreEvent_Complete(ByVal Message As String)
    PrintStat "oRestoreEvent_Complete -- " & Message
End Sub

Private Sub oRestoreEvent_NextMedia(ByVal Message As String)
    PrintStat "oRestoreEvent_NextMedia -- " & Message
End Sub

Private Sub oRestoreEvent_PercentComplete(ByVal Message As String, ByVal Percent As Long)
    PrintStat "oRestoreEvent_PercentComplete -- " & Message & " " & Percent
End Sub

Private Sub PrintStat(ByRef Message As String)
    txtStatus.Text = txtStatus.Text + Message + vbCrLf
End Sub



Private Sub optSSAuth_Click()
    If optSSAuth.Value = True Then
        SSAuthOptionsOn
    End If
End Sub

Private Sub optWinNTAuth_Click()
    optWinNTAuth.Value = True
    WinNTAuthOptionsOn
    txtUserName.Text = ""
    txtPassword.Text = ""
End Sub

Private Sub buttonsConnectClosed()
    cmdConnect.Default = True
    
    cmdConnect.Enabled = True
    cmdBackup.Enabled = True
    cmdRestore.Enabled = True
    cmdDisconnect.Enabled = False
    
    cmdBrowse.Enabled = False
    cmbDatabaseName.Enabled = False
    txtDataFileName.Enabled = False
    
    ' Enable the Authorization stuff.
    optWinNTAuth.Enabled = True
    optSSAuth.Enabled = True
    txtServerName.Enabled = True
    lblServer.Enabled = True
    If optWinNTAuth = True Then
        WinNTAuthOptionsOn
    Else
        SSAuthOptionsOn
    End If
End Sub

Private Sub buttonsConnectOpen()
    cmdConnect.Enabled = False
    cmdBackup.Enabled = True
    cmdRestore.Enabled = True
    cmdDisconnect.Enabled = True
    
    cmdBrowse.Enabled = True
    cmbDatabaseName.Enabled = True
    txtDataFileName.Enabled = True
    
    ' Disable the Authorization stuff.
    optWinNTAuth.Enabled = False
    optSSAuth.Enabled = False
    txtServerName.Enabled = False
    lblServer.Enabled = False
    lblUserName.Enabled = False
    lblPassword.Enabled = False
    txtUserName.Enabled = False
    txtPassword.Enabled = False
End Sub

Private Sub WinNTAuthOptionsOn()
    lblUserName.Enabled = False
    lblPassword.Enabled = False
    txtUserName.Enabled = False
    txtPassword.Enabled = False
End Sub

Private Sub SSAuthOptionsOn()
    lblUserName.Enabled = True
    lblPassword.Enabled = True
    txtUserName.Enabled = True
    txtPassword.Enabled = True
End Sub




'Private Sub FillDatabaseList()
'    cmbDatabaseName.Clear
'
'    ' Enumerate all of the databases and add the names to the list box.
'    Dim oDB As SQLDMO.Database
'    For Each oDB In gSQLServer.Databases
'        If oDB.SystemObject = False Then
'            cmbDatabaseName.AddItem oDB.Name
'        End If
'    Next oDB
'
'    ' Take care of the assignment of gBkupRstrFilePath.
'    Dim MyPos As Integer
'    gBkupRstrFilePath = CurDir
'    MyPos = InStr(1, CurDir, "DevTools", 1)
'    If MyPos > 0 Then
'        gBkupRstrFilePath = Left(gBkupRstrFilePath, MyPos - 1)
'        If Len(Dir(gBkupRstrFilePath + "backup", vbDirectory)) Then
'            gBkupRstrFilePath = gBkupRstrFilePath + "backup\"
'        Else
'            gBkupRstrFilePath = "c:\"
'        End If
'    Else
'        gBkupRstrFilePath = "c:\"
'    End If
'
'    ' Select the first database in the list.
'    If cmbDatabaseName.ListCount > 0 Then
'        cmbDatabaseName.ListIndex = 0
'        ' Assign the default backup/restore file name.
'        If Len(cmbDatabaseName.Text) > 0 Then
'            txtDataFileName.Text = gBkupRstrFilePath + cmbDatabaseName.Text + ".bak"
'        End If
'    End If
'
'End Sub


Private Sub cmbDatabaseName_Click()
    ' Assign the default backup/restore file name.
    If Len(cmbDatabaseName.Text) > 0 Then
        txtDataFileName.Text = gBkupRstrFilePath + cmbDatabaseName.Text + ".bak"
    End If
End Sub

Private Sub txtBranchCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If Lastkey(KeyCode) And txtbranchcode.Text <> "" Then
         txtbranchcode = DoPad(txtbranchcode, 3)
         lb_found = MySeek(txtbranchcode.Text, "BranchCode", PR_Branch)
        
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtbranchcode.SetFocus
         Else
             Call backuppath
         End If
 ElseIf KeyCode = vbKeyF12 Then
     Call Command4_Click
 End If
End Sub


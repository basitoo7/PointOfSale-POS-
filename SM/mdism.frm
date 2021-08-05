VERSION 5.00
Begin VB.Form mdism 
   BackColor       =   &H00808000&
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   5715
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "mdism.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5715
   WindowState     =   2  'Maximized
   Begin VB.Menu Sm_Mtn 
      Caption         =   "&Maintain"
      Begin VB.Menu Sm_Comp 
         Caption         =   "Companies"
      End
      Begin VB.Menu FCM_Branches 
         Caption         =   "Company Branches"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu System_Group 
         Caption         =   "System Group"
      End
      Begin VB.Menu Group_Right 
         Caption         =   "Group Rights"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu System_Modules 
         Caption         =   "System Modules"
      End
      Begin VB.Menu KimSm_Proc 
         Caption         =   "System Objects"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu KimSm_Users 
         Caption         =   "User Maintenance"
      End
   End
   Begin VB.Menu Sm_Trns 
      Caption         =   "&Transactions"
      Begin VB.Menu KimSm_YearClose 
         Caption         =   "Year Closing"
      End
      Begin VB.Menu DbBackup 
         Caption         =   "DataBase Backup"
      End
   End
   Begin VB.Menu Sm_Rpts 
      Caption         =   "&Utilities"
      Begin VB.Menu Sm_Cal 
         Caption         =   "Calculator"
      End
      Begin VB.Menu Sm_Calendar 
         Caption         =   "Calendar"
      End
   End
   Begin VB.Menu sys_reports 
      Caption         =   "Reports"
      Begin VB.Menu Users_Log 
         Caption         =   "Users Log"
      End
   End
   Begin VB.Menu Sm_Retrn 
      Caption         =   "&Return"
   End
End
Attribute VB_Name = "mdism"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Opt As String

   ' Dim SqlApp As New SQLDMO.Application
   ' Dim ITsServer As New SQLServer
  '  Dim ITsServer2 As New SQLServer2
    
    
  '  Dim ServerList As String
   ' Dim fsobj, f


Private Sub DBBackup_Click()
    'On Error GoTo Master
    'ITsServer.Connect Replace(ServerName, "'", ""), SUserID, SPassword
    
    'ITsServer.Connect Replace("ecountsserver", "'", ""), "sa", ""
      
   '   Gs_DBDataSource

  '  ITsServer.Connect Replace("AMTC1\SQLSERVER2005", "'", ""), "sa", ""
    
 
   ' ServerList = SqlApp.ListAvailableSQLServers(0)

   ' ITsServer.Connect Replace(ServerList, "'", ""), "sa", ""

 '   Dim ITSBackup As New SQLDMO.Backup
    
  '  Set fsobj = CreateObject("Scripting.FileSystemObject")
 '   If fsobj.FolderExists("d:\RahatNesPakDb_Backup") = False Then
'        If fsobj.DriveExists("D:") = True Then
'            FileSystem.MkDir "D:\RahatNesPakDb_Backup"
 '       Else
  '          FileSystem.MkDir "C:\RahatNesPakDb_Backup"
 '       End If
'    Else
'    End If
    
    
   ' ITsServer.Databases = "ecounts"
    
  '  ITSBackup.Action = SQLDMOBackup_Database
 '   ITSBackup.Database = "Ecounts"
    
'    If fsobj.DriveExists("D:") = True Then
'        ITSBackup.Files = "D:\RahatNesPakDb_Backup\RahatNesPakDb_" & Format(Date, "yyyy") & "_" & Format(Date, "MM") & "_" & Format(Date, "dd") & "_" & Format(Time, "HH") & Format(Time, "MM") & Format(Time, "SS") & ".Bak"
'    Else
'        ITSBackup.Files = "C:\RahatNesPakDb_Backup\RahatNesPakDb_" & Format(Date, "yyyy") & "_" & Format(Date, "MM") & "_" & Format(Date, "dd") & "_" & Format(Time, "HH") & Format(Time, "MM") & Format(Time, "SS") & ".Bak"
'    End If
'    Screen.MousePointer = vbHourglass
'    ITSBackup.SQLBackup ITsServer
    'ITSBackup.SQLBackup
'    Screen.MousePointer = vbDefault
'    MsgBox "BackUp Operation Completed Successfully . ", vbOKOnly
    ', Gs_MsgTitle
'    Exit Sub
'Master:
'    Screen.MousePointer = vbDefault
 '   MsgBox "Request Terminated " & vbCrLf & "Error Occured While BackingUp DataBase. " & vbCrLf & "Invalid File Name May Cause Errors." & vbCrLf & Err.Description, vbExclamation
    ', Gs_MsgTitle
'    Err.Clear
End Sub

Private Sub FCM_Branches_Click()
   frmBranches.Show
End Sub

'Private Sub FCM_Branches_Click()
'   frmBranches.Show
'End Sub

Private Sub Form_Load()
    Sm_Comp.Enabled = chkRights1("SYMGR00001")
    FCM_Branches.Enabled = chkRights1("SYMGR00002")
    System_Group.Enabled = chkRights1("SYMGR00003")
    Group_Right.Enabled = chkRights1("SYMGR00004")
    System_Modules.Enabled = chkRights1("SYMGR00005")
    KimSm_Proc.Enabled = chkRights1("SYMGR00006")
    KimSm_Users.Enabled = chkRights1("SYMGR00007")
    KimSm_YearClose.Enabled = chkRights1("SYMGR00008")

End Sub

Private Sub Group_Right_Click()
frmSMGroupRights.Show
End Sub

Private Sub KimSm_Proc_Click()
  frmSMObjects.Show
End Sub



Private Sub KimSm_Users_Click()
  frmSMUsers.Show
End Sub

Private Sub KimSm_YearClose_Click()
    Me.Tag = 1
    Para_Rs.Filter = adFilterNone
    frmGroupCompanies.Show 1
    Module1.YearClose
    ParaCntr_Rs.Close
End Sub


Private Sub Sm_Cal_Click()
On Error GoTo LocalErr
Dim X
X = Shell("Calc.exe", 1)
Exit Sub
LocalErr:
Call SetErr("Calculator exe file not found call system administrator", vbCritical)
End Sub

Private Sub Sm_Calendar_Click()
On Error GoTo LocalErr:
frmcalandar.Show
Exit Sub
LocalErr:
Call SetErr("Processing error call system administrator", vbCritical)
End Sub

Private Sub Sm_Comp_Click()
   Frmcompstp.Show
End Sub

Private Sub Sm_Retrn_Click()
    MDIForm1.Toolbar1.Visible = False
    MDIForm1.Toolbar2.Top = 0
    MDIForm1.Toolbar2.Visible = True
    If ParaCntr_Rs.State = 1 Then ParaCntr_Rs.Close
    Unload Me
End Sub

Private Sub Sm_SysInst_Click()
    frmRegister.Show
End Sub

Private Sub System_Group_Click()
frmSMGroups.Show
End Sub

Private Sub System_Modules_Click()
frmSMModules.Show
End Sub




Private Sub Users_Log_Click()
frmfromtodate.Show
End Sub

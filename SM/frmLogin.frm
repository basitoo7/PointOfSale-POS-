VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4335
   ClientLeft      =   2880
   ClientTop       =   3240
   ClientWidth     =   5445
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2561.262
   ScaleMode       =   0  'User
   ScaleWidth      =   5112.56
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3660
      Left            =   255
      Picture         =   "frmLogin.frx":030A
      ScaleHeight     =   3660
      ScaleWidth      =   4890
      TabIndex        =   5
      Top             =   285
      Width           =   4890
      Begin VB.TextBox txtUserName 
         BackColor       =   &H00FFFF00&
         DataSource      =   "CUser"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1665
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1005
         Width           =   1950
      End
      Begin VB.TextBox txtPassword 
         DataSource      =   "lc_Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1665
         MaxLength       =   35
         PasswordChar    =   "*"
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1455
         Width           =   1950
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ok"
         Height          =   330
         Left            =   1665
         TabIndex        =   7
         Top             =   2100
         Width           =   915
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   330
         Left            =   2715
         TabIndex        =   6
         Top             =   2100
         Width           =   915
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "User Id :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   0
         Left            =   750
         TabIndex        =   9
         Top             =   1050
         Width           =   960
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   1
         Left            =   540
         TabIndex        =   8
         Top             =   1500
         Width           =   1080
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6870
      Top             =   2595
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "09:23:34"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4305
      TabIndex        =   4
      Top             =   4020
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label lblLogin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login Time :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3090
      TabIndex        =   3
      Top             =   4020
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label lblGoodMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Demo Version"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2730
      TabIndex        =   2
      Top             =   3345
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim startdate As String
Dim enddate   As Date

Dim Pr_Pass As Recordset
Dim SyPassword As String
Dim ls_CompName As String
Dim Pr_Auser As New Recordset
Dim pr_csetting As New Recordset


Private Sub Command1_Click()
Call txtpassword_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
Set Pr_Pass = New Recordset
   
  Pr_Pass.Open "Select SyUsers.* from syUsers order by 1", gc_dbcon, adOpenStatic, adLockReadOnly

'commint below line at final stages
'SendKeys "admin{enter}admin{ENTER}"

End Sub



Private Sub Timer1_Timer()
lblTime = Time
End Sub

Private Sub txtpassword_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ls_CompName1 As String
   If KeyCode = vbKeyReturn Then
     If Trim(txtPassword) <> "" Then
      If UCase(txtPassword.Text) = SyPassword Then
         Gc_UserId = txtUsername.Text
         ls_CompName = String(50, "0")
         Call GetComputerName(ls_CompName, 50)
         ls_CompName1 = Replace(ls_CompName, 0, "")
         ls_CompName1 = Left(ls_CompName1, Len(ls_CompName1) - 1)
         Gs_ComputerName = ls_CompName1
         
        pr_csetting.Open "select * from Sys_ComSetting where compname = '" & Gs_ComputerName & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If Not pr_csetting.EOF Then
        gn_comportset = Val(0 & pr_csetting("comsettingno"))
        End If
        pr_csetting.Close
        
'If gn_comportset = 0 Then gn_comportset = 1
         
          gc_dbcon.Execute "Update SyUsers  set activestatus = 0 where userid = '" & Gc_UserId & "' and activestatus < 0"
             
         ' gc_dbcon.Execute "Insert into Sys_passlog(CompCode, Userid, Remarks, Uoption, Adddate) Values ('001','" & Gc_UserId & "','" & "LogIn :" + Trim(ls_CompName1) & "',1,'" & Format(Gd_SysDate, "YYYY/MM/DD HH:MM:SS") & "')"
          
          
            GR_SMURights.Open "Select * from Sys_UserRights where grouporusercode = " & Gn_UserCode & "", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
          Unload Me
      
  If ln_posaccess = 1 Or ln_posaccess = 2 Or ln_posaccess = 3 Then
      
   'this area for the cashier
   
            frmGroupCompanies.SetValues ("001")
            MDIForm1.StatusBar1.Panels(2).Text = Gc_UserId
            MDIForm1.StatusBar1.Panels(6).Text = Gs_CompName
            MDIForm1.Toolbar2.Visible = False
            MDIForm1.Toolbar5.Top = 0
            MDIForm1.Toolbar5.Visible = True
      
      If ln_posaccess = 1 Or ln_posaccess = 2 Or ln_posaccess = 3 Then
              
            
          Pr_Auser.Open "select * from syusers where activestatus  >= 2 and userid = '" & Gc_UserId & "'", gc_dbcon, adOpenStatic, adLockReadOnly, adCmdText
         If Not Pr_Auser.EOF Then
            MsgBox ("User Already Active At " & ls_CompName1 & " Computer")
            End
           gc_dbcon.Execute "Update SyUsers  set activestatus = (activestatus + 1),compname = '" & Trim(ls_CompName1) & "', logintime = '" & Time & "',logouttime = 'Still Active' where userid = '" & Gc_UserId & "'"
          Else
            gc_dbcon.Execute "Update SyUsers  set activestatus = (activestatus + 1) ,compname = '" & Trim(ls_CompName1) & "', logintime = '" & Time & "',logouttime = 'Still Active' where userid = '" & Gc_UserId & "'"
          End If
          Pr_Auser.Close
           
            MDIForm1.Toolbar5.Buttons(1).Enabled = False
            MDIForm1.Toolbar5.Buttons(4).Enabled = False
            MDIForm1.Toolbar5.Buttons(5).Enabled = False
            MDIForm1.Toolbar5.Buttons(6).Enabled = False
            MDIForm1.Toolbar5.Buttons(6).Enabled = False
            MDIForm1.Toolbar5.Buttons(7).Enabled = False
            MDIForm1.KimFile_Lgin.Item(12).Enabled = False
            MDIForm1.KimFile_LgOt.Item(13).Enabled = False
            MDIForm1.FCM_Import.Enabled = False
            MDIForm1.Fcm_Export.Enabled = False
            MDIForm1.KimFile_PwdChg.Item(14).Enabled = False
            MDIForm1.KimFile_Exit.Item(15).Enabled = False
            MDIForm1.Change_Server.Enabled = False
            MDIForm1.KimFile_Open.Enabled = False
            MDIForm1.FCM_Auser.Enabled = False
        
           ' MDIForm1.MaximizeBox = False
           ' MDIForm1.MinimizeBox = False
           
         End If

            MDIForm1.Show
            MDIForm1.WindowState = 1
            frmSO_Posform.Show
            
         Else
       
         
         
            MDIForm1.Show
         End If
   Else
         Call SetErr("Incorrect Password.", vbCritical)
         txtPassword.Text = ""
      txtPassword.SetFocus
  End If
   
  Else
       Call SetErr("Please Enter Password.", vbCritical)
       txtPassword.SetFocus
   End If
End If
End Sub

Private Sub txtusername_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ls_Pathname As String
Dim ls_File As New FileSystemObject


   If Lastkey(KeyCode) Then
      Dim lb_found As Boolean
      lb_found = MySeek(UCase(txtUsername), "userid", Pr_Pass)
      If lb_found Then
         SyPassword = UCase(LTrim(RTrim(DeCode(Pr_Pass.Fields("Password").Value, 50))))
         Gn_UserCode = Pr_Pass("UserCode")
         Gs_Siteid = Trim(Pr_Pass("Siteid") & "")
         ln_posaccess = Val(0 & Pr_Pass("Posaccess"))
         Gn_posaccess = ln_posaccess
         'ln_changeprinter = Val(0 & Pr_Pass("PrinterType"))
         Gc_UserName = Trim(StrConv(Pr_Pass.Fields("UserName"), 3)) & ""
         
         
        ' If ls_File.FileExists(App.Path & Gs_PictPath & "\" & Trim(txtUsername) & ".bmp") Then
        '   Picture1.Picture = LoadPicture(App.Path & Gs_PictPath & "\" & Trim(txtUsername) & ".bmp")
        ' Else
        '  Picture1.Picture = LoadPicture()
        ' End If
         txtPassword.SetFocus
        ' Call Logintime
      Else
         Call SetErr("Invalid UserId.", vbCritical)
         txtUsername.Text = ""
         txtUsername.SetFocus
      End If
   End If
   
End Sub
Private Sub Logintime()
If Hour(Time) >= 5 And Hour(Time) <= 11 Then
    lblGoodMsg = "Good Morning " & StrConv(Gc_UserName, 3)
ElseIf Hour(Time) >= 12 And Hour(Time) <= 16 Then
    lblGoodMsg = "Good Afternoon " & StrConv(Gc_UserName, 3)
ElseIf Hour(Time) >= 16 And Hour(Time) <= 19 Then
    lblGoodMsg = "Good Evening " & StrConv(Gc_UserName, 3)
 Else
    lblGoodMsg = "Good Night " & StrConv(Gc_UserName, 3)
End If
 lblTime = Time
 lblTime.Visible = True
 lblGoodMsg.Visible = True
 lblLogin.Visible = True
End Sub

Private Sub txtUsername_LostFocus()
If txtUsername <> "" Then
    txtUsername = LCase(txtUsername)
    Call txtusername_KeyDown(vbKeyReturn, vbKeyShift)
End If
End Sub

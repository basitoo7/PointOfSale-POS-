VERSION 5.00
Begin VB.Form frmChgPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   2565
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3900
   Icon            =   "frmChgPass.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515.488
   ScaleMode       =   0  'User
   ScaleWidth      =   3661.889
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   2085
      TabIndex        =   10
      Top             =   2115
      Width           =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   330
      Left            =   2940
      TabIndex        =   9
      Top             =   2115
      Width           =   840
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   3840
      Begin VB.TextBox txtoldpassword 
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
         Left            =   1275
         MaxLength       =   35
         PasswordChar    =   "*"
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   660
         Width           =   2325
      End
      Begin VB.TextBox TxtConfirm 
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
         Left            =   1275
         MaxLength       =   35
         PasswordChar    =   "*"
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2325
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
         Left            =   1275
         MaxLength       =   35
         PasswordChar    =   "*"
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1140
         Width           =   2325
      End
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
         Left            =   1275
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Password :"
         Height          =   270
         Index           =   3
         Left            =   435
         TabIndex        =   8
         Top             =   675
         Width           =   960
      End
      Begin VB.Label lblLabels 
         Caption         =   "Confirm  :"
         Height          =   270
         Index           =   2
         Left            =   540
         TabIndex        =   6
         Top             =   1575
         Width           =   720
      End
      Begin VB.Label lblLabels 
         Caption         =   "New Password :"
         Height          =   270
         Index           =   1
         Left            =   60
         TabIndex        =   5
         Top             =   1155
         Width           =   1170
      End
      Begin VB.Label lblLabels 
         Caption         =   "&User Name :"
         Height          =   270
         Index           =   0
         Left            =   330
         TabIndex        =   4
         Top             =   255
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmChgPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim startdate As String
Dim enddate   As Date
Dim cntsql As New ADODB.Command
Dim Ls_Newpassword As String
Dim Ls_Newpassword1 As String
     
Dim Pr_Pass As Recordset
Dim SyPassword As String

Private Sub Command1_Click()
Call txtConfirm_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
cntsql.ActiveConnection = gc_dbcon
cntsql.CommandText = adCmdText
Set Pr_Pass = New Recordset
   
  Pr_Pass.Open "Select SyUsers.* from syUsers order by 1", gc_dbcon, adOpenDynamic, adLockOptimistic
End Sub
Private Sub txtConfirm_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo LocalErr
  If Lastkey(KeyCode) Then
   If (UCase(Gc_UserId) = "ADMIN" Or UCase(Gc_UserId) = UCase(txtusername.Text)) Then
        If UCase(Gc_UserId) <> "VENDOR" And UCase(txtusername.Text) = "VENDOR" Then
          Call SetErr("Change Prohabited.", vbCritical)
          txtusername = ""
          txtusername.SetFocus
        Else
         'Ls_Newpassword1 = UCase(LTrim(RTrim(DeCode(txtoldpassword.Text, 50))))
       
         If UCase(txtoldpassword.Text) = UCase(Ls_Newpassword1) Then
   
            If txtpassword.Text = txtConfirm.Text And (Len(txtpassword.Text) <= txtpassword.MaxLength And Len(txtConfirm.Text) <= txtConfirm.MaxLength) Then
               Ls_Newpassword = EnCode(txtConfirm.Text, 50)
               
               cntsql.CommandText = "Update SyUsers Set Password = '" & Ls_Newpassword & "' Where UserId = '" & txtusername.Text & "'"
               cntsql.Execute
               Call SetErr("Your Password has been changed successfully.", vbInformation)
               txtusername = ""
               txtpassword = ""
               txtConfirm = ""
               txtoldpassword = ""
               txtusername.SetFocus
            Else
               Call SetErr("Unmached passwords.", vbCritical)
               txtpassword.SetFocus
            End If
            
          Else
          Call MsgBox("Old Password not Valid !!!", vbCritical)
          txtoldpassword.SetFocus
          End If
         End If
    Else
        Call SetErr("You cannot change password of other user.", vbCritical)
         txtusername = ""
         txtpassword = ""
         txtConfirm = ""
         txtusername.SetFocus
  End If
  End If
Exit Sub
LocalErr:
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub txtoldpassword_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtpassword.SetFocus
End Sub

Private Sub txtpassword_KeyDown(KeyCode As Integer, Shift As Integer)
   If Lastkey(KeyCode) Then txtConfirm.SetFocus
End Sub

Private Sub txtusername_KeyDown(KeyCode As Integer, Shift As Integer)
   If Lastkey(KeyCode) Then
      Dim lb_found As Boolean
      If (UCase(Gc_UserId) = "ADMIN" Or UCase(Gc_UserId) = UCase(txtusername.Text)) Then
      If UCase(Gc_UserId) <> "VENDOR" And UCase(txtusername.Text) = "VENDOR" Then
          Call SetErr("Change Prohabited.", vbCritical)
          txtusername = ""
          txtusername.SetFocus
      Else
      lb_found = MySeek(UCase(txtusername.Text), "userid", Pr_Pass)
      Ls_Newpassword1 = UCase(LTrim(RTrim(DeCode(Pr_Pass.Fields("Password").Value, 50))))
      If lb_found Then
         txtoldpassword.SetFocus
      Else
         Call SetErr("Invalid UserId.", vbCritical)
         txtusername.Text = ""
         txtusername.SetFocus
      End If
      End If
      Else
         Call SetErr("You cannot change password of other user.", vbCritical)
         txtusername = ""
         txtusername.SetFocus
      End If
   End If
   
End Sub



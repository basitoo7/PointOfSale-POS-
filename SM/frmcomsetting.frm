VERSION 5.00
Begin VB.Form frmcomportsetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Com Port Setting"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2745
   Icon            =   "frmcomsetting.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   1890
      TabIndex        =   3
      Top             =   840
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   360
      Left            =   1035
      TabIndex        =   2
      Top             =   840
      Width           =   780
   End
   Begin VB.TextBox txtcompport 
      Height          =   330
      Left            =   1215
      TabIndex        =   0
      Top             =   240
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   "Com Port :"
      Height          =   360
      Left            =   420
      TabIndex        =   1
      Top             =   255
      Width           =   1080
   End
End
Attribute VB_Name = "frmcomportsetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pr_csetting As New Recordset

Private Sub Command1_Click()
If pr_csetting.State = 1 Then pr_csetting.Close
pr_csetting.Open "select * from Sys_ComSetting where compname = '" & Gs_ComputerName & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_csetting.EOF Then
gc_dbcon.Execute "update Sys_ComSetting  set comsettingno = " & Val(txtcompport) & " where compname = '" & Gs_ComputerName & "'"
Else
gc_dbcon.Execute "insert into  Sys_ComSetting (compname,comsettingno) values ('" & Gs_ComputerName & "', " & Val(txtcompport) & ") "
End If
pr_csetting.Close
gn_comportset = Val(txtcompport)
Call MsgBox("Successfully Updated")
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
If pr_csetting.State = 1 Then pr_csetting.Close
pr_csetting.Open "select * from Sys_ComSetting", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_csetting.EOF Then
txtcompport = Val(0 & pr_csetting("comsettingno"))
End If
pr_csetting.Close
End Sub

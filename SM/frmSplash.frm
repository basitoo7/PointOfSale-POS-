VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5340
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3480
      Left            =   270
      Picture         =   "frmSplash.frx":000C
      ScaleHeight     =   3480
      ScaleWidth      =   4845
      TabIndex        =   2
      Top             =   345
      Width           =   4845
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   6420
      Top             =   1380
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Please Wait..."
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
      Height          =   240
      Left            =   255
      TabIndex        =   1
      Top             =   3930
      Width           =   4725
   End
   Begin VB.Label lblLicenseTo 
      BackStyle       =   0  'Transparent
      Caption         =   "Demo Versin"
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
      Height          =   255
      Left            =   255
      TabIndex        =   0
      Top             =   90
      Width           =   6165
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    'lblVersion.Caption = "Build " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
    If Gl_Demo Then
      lblLicenseTo = "Demo Version."
    Else
      lblLicenseTo = "This Product is Licensed To " & Gs_RegisterTo
    End If
   ' LoadPicture1
End Sub

Private Sub Timer1_Timer()
        Unload Me
        frmLogin.Show
End Sub
Private Sub LoadPicture1()
On Error Resume Next
Dim ln_pic As Integer
Dim ls_pic As String
ln_pic = CInt(Int((Val(13) * Rnd()) + 1))

ls_pic = Trim(str(ln_pic)) & ".jpg"
Picture1.Picture = LoadPicture(App.Path & "\pictures\" & ls_pic)

End Sub

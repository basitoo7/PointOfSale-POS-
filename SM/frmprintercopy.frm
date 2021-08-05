VERSION 5.00
Begin VB.Form frmSoPrinterCopy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Printer Copy"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmprintercopy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3675
      TabIndex        =   1
      Top             =   885
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   315
      Left            =   2700
      TabIndex        =   0
      Top             =   885
      Width           =   945
   End
   Begin VB.TextBox txtnoofcopy 
      Height          =   285
      Left            =   1290
      TabIndex        =   2
      Text            =   "1"
      Top             =   225
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "No of Copies:"
      Height          =   225
      Left            =   195
      TabIndex        =   3
      Top             =   225
      Width           =   1515
   End
End
Attribute VB_Name = "frmSoPrinterCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmSO_PosformCredit.ln_printerCopy = Val(txtnoofcopy)
Unload Me
End Sub

Private Sub Command2_Click()
frmSO_PosformCredit.ln_printerCopy = 0
Unload Me
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmcalandar 
   Caption         =   "System Calandar"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3045
   Icon            =   "frmcalandar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   3045
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2760
      Left            =   30
      TabIndex        =   0
      Top             =   -75
      Width           =   2970
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   135
         TabIndex        =   1
         Top             =   255
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   320995329
         CurrentDate     =   37589
      End
   End
End
Attribute VB_Name = "frmcalandar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

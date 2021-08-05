VERSION 5.00
Begin VB.Form frmzapdata 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Truncate Data"
   ClientHeight    =   4095
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   2985
   Icon            =   "frmzapdata.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2419.464
   ScaleMode       =   0  'User
   ScaleWidth      =   2802.753
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H80000008&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   270
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Save"
      Top             =   3630
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Discard 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1530
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Save License info."
      Top             =   3630
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2955
      Begin VB.CheckBox Check9 
         Alignment       =   1  'Right Justify
         Caption         =   "Payroll Transaction Files :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   9
         Top             =   3060
         Width           =   2475
      End
      Begin VB.CheckBox Check8 
         Alignment       =   1  'Right Justify
         Caption         =   "Payroll Setup Files :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   8
         Top             =   2790
         Width           =   2475
      End
      Begin VB.CheckBox Check7 
         Alignment       =   1  'Right Justify
         Caption         =   "AR Transaction Files :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   7
         Top             =   2340
         Width           =   2475
      End
      Begin VB.CheckBox Check6 
         Alignment       =   1  'Right Justify
         Caption         =   "AR. Setup Files :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   6
         Top             =   2070
         Width           =   2475
      End
      Begin VB.CheckBox Check5 
         Alignment       =   1  'Right Justify
         Caption         =   "Inventory Transaction Files :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   5
         Top             =   1620
         Width           =   2475
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Caption         =   "Inventory Setup Files :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   4
         Top             =   1320
         Width           =   2475
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Caption         =   "GL Transaction Files :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   3
         Top             =   900
         Width           =   2475
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "GL Setup Files :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Top             =   630
         Width           =   2475
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "System Manager Files :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   2475
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   2940
         Y1              =   2700
         Y2              =   2700
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   2940
         Y1              =   1980
         Y2              =   1980
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   2940
         Y1              =   1230
         Y2              =   1230
      End
      Begin VB.Line Line1 
         X1              =   30
         X2              =   2970
         Y1              =   570
         Y2              =   570
      End
   End
End
Attribute VB_Name = "frmzapdata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_Discard_Click()
  Unload Me
End Sub

Private Sub Cmd_Save_Click()
Dim choice As Integer

    choice = SetErr("Are you sure.  ", vbYesNo)
           If choice = vbYes Then
           
             If Check1.Value = 1 Then
                gc_dbcon.Execute ("truncate table Syscomp")
                gc_dbcon.Execute ("truncate table SysFins")
                gc_dbcon.Execute ("truncate table SysTax")
             End If
             
             If Check2.Value = 1 Then
                gc_dbcon.Execute ("truncate table gl_sub0")
                gc_dbcon.Execute ("truncate table gl_sub1")
                gc_dbcon.Execute ("truncate table gl_sub2")
                gc_dbcon.Execute ("truncate table gl_sub3")
                gc_dbcon.Execute ("truncate table gl_sub4")
                gc_dbcon.Execute ("truncate table gl_sub5")
                gc_dbcon.Execute ("truncate table gl_sub6")
                gc_dbcon.Execute ("truncate table gl_sub7")
                gc_dbcon.Execute ("truncate table gl_sub8")
                gc_dbcon.Execute ("truncate table gl_sub9")
                gc_dbcon.Execute ("truncate table gl_Detail")
                
                Dim Pr_VType As New Recordset
                Pr_VType.Open "Select * From GLVchrtype", gc_dbcon, adOpenDynamic, adLockOptimistic
                If Not Pr_VType.EOF Then
                   Do While Not Pr_VType.EOF
                      Pr_VType.Fields("VchrCount") = 0
                      Pr_VType.Fields("VchrMonth1") = 0
                      Pr_VType.Fields("VchrMonth2") = 0
                      Pr_VType.Fields("VchrMonth3") = 0
                      Pr_VType.Fields("VchrMonth4") = 0
                      Pr_VType.Fields("VchrMonth5") = 0
                      Pr_VType.Fields("VchrMonth6") = 0
                      Pr_VType.Fields("VchrMonth7") = 0
                      Pr_VType.Fields("VchrMonth8") = 0
                      Pr_VType.Fields("VchrMonth9") = 0
                      Pr_VType.Fields("VchrMonth10") = 0
                      Pr_VType.Fields("VchrMonth11") = 0
                      Pr_VType.Fields("VchrMonth12") = 0
                      
                      Pr_VType.Update
                      Pr_VType.MoveNext
                      If Pr_VType.EOF Then Exit Do
                   Loop
                   Pr_VType.Close
                End If
             End If

             If Check3.Value = 1 Then
                gc_dbcon.Execute ("truncate table gl_Ref")
                gc_dbcon.Execute ("truncate table gl_Trans")
                gc_dbcon.Execute ("truncate table gl_budget")
                gc_dbcon.Execute ("truncate table gl_RptRouting")
                
             End If

             If Check4.Value = 1 Then
                gc_dbcon.Execute ("truncate table IC_Item")
                gc_dbcon.Execute ("truncate table IC_ItemClass")
                gc_dbcon.Execute ("truncate table IC_Job")
                gc_dbcon.Execute ("truncate table IC_Locations")
                gc_dbcon.Execute ("truncate table IC_Supplier")
             End If

             If Check5.Value = 1 Then
                gc_dbcon.Execute ("truncate table IC_Trans")
             End If

             If Check6.Value = 1 Then
                gc_dbcon.Execute ("truncate table AR_Areas")
                gc_dbcon.Execute ("truncate table AR_Customer")
                gc_dbcon.Execute ("truncate table Ar_SaleTypes")
             End If

             If Check7.Value = 1 Then
                gc_dbcon.Execute ("truncate table AR_Master")
                gc_dbcon.Execute ("truncate table Ar_Receipts")
             End If

             If Check8.Value = 1 Then
                gc_dbcon.Execute ("truncate table HRM_Areas")
                gc_dbcon.Execute ("truncate table HRM_Bank")
                gc_dbcon.Execute ("truncate table HRM_CmpAllow")
                gc_dbcon.Execute ("truncate table HRM_CmpDeds")
                gc_dbcon.Execute ("truncate table HRM_Lvs")
                gc_dbcon.Execute ("truncate table HRM_EmpAdj")
                gc_dbcon.Execute ("truncate table HRM_EmpAdv")
                gc_dbcon.Execute ("truncate table HRM_EmpAllow")
                gc_dbcon.Execute ("truncate table HRM_EmpDedus")
                gc_dbcon.Execute ("truncate table HRM_EmpLvs")
                gc_dbcon.Execute ("truncate table HRM_Grades")
             End If

             If Check9.Value = 1 Then
                gc_dbcon.Execute ("truncate table HRM_SalAllow")
                gc_dbcon.Execute ("truncate table HRM_SalDeduc")
             End If
           End If
End Sub


VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRegister 
   Caption         =   "Software License "
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   Icon            =   "frmregister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7095
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   1200
      TabIndex        =   28
      Top             =   2820
      Width           =   3975
      Begin VB.CommandButton Cmd_Discard 
         Caption         =   "Discard"
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
         Left            =   2040
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Save License info."
         Top             =   180
         Width           =   1575
      End
      Begin VB.CommandButton Cmd_Save 
         Caption         =   "Save"
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
         Left            =   300
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Save License info."
         Top             =   180
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1935
      Left            =   4740
      TabIndex        =   20
      Top             =   780
      Width           =   2325
      Begin VB.CheckBox chkModules 
         DataSource      =   "gn_sysAm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   6
         Left            =   2010
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1620
         Width           =   255
      End
      Begin VB.CheckBox chkModules 
         DataSource      =   "gn_sysHRI"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   5
         Left            =   2010
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1380
         Width           =   255
      End
      Begin VB.CheckBox chkModules 
         DataSource      =   "gn_sysSo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   4
         Left            =   2010
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1140
         Width           =   255
      End
      Begin VB.CheckBox chkModules 
         DataSource      =   "gn_sysPo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   2010
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   900
         Width           =   255
      End
      Begin VB.CheckBox chkModules 
         DataSource      =   "gn_sysGL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   2010
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   660
         Width           =   255
      End
      Begin VB.CheckBox chkModules 
         DataSource      =   "gn_sysap"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   2010
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   420
         Width           =   255
      End
      Begin VB.CheckBox chkModules 
         DataSource      =   "gn_sysar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   2010
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   180
         Width           =   255
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "temp:"
         Height          =   255
         Left            =   390
         TabIndex        =   27
         Top             =   1620
         Width           =   1575
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Placements Management :"
         Height          =   255
         Left            =   60
         TabIndex        =   26
         Top             =   1380
         Width           =   1905
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "COIs :"
         Height          =   255
         Left            =   1110
         TabIndex        =   25
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Lease Management :"
         Height          =   255
         Left            =   60
         TabIndex        =   24
         Top             =   900
         Width           =   1905
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "General Ledger :"
         Height          =   255
         Left            =   60
         TabIndex        =   23
         Top             =   660
         Width           =   1905
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer Registration:"
         Height          =   255
         Left            =   60
         TabIndex        =   22
         Top             =   420
         Width           =   1905
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Inflow Management :"
         Height          =   255
         Left            =   60
         TabIndex        =   21
         Top             =   180
         Width           =   1905
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   60
      TabIndex        =   15
      Top             =   780
      Width           =   4695
      Begin MSMask.MaskEdBox txtGroupaddr1 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Enter Group Address 1"
         Top             =   180
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "c"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtgroupaddr2 
         Height          =   315
         Left            =   840
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Enter 2nd Address"
         Top             =   540
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "c"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtgroupcity 
         Height          =   315
         Left            =   840
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   """""-abc"
         ToolTipText     =   "Enter City"
         Top             =   900
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "c"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "City :"
         Height          =   255
         Left            =   420
         TabIndex        =   19
         Top             =   900
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Address :"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   180
         Width           =   675
      End
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
      Height          =   675
      Left            =   60
      TabIndex        =   14
      Top             =   60
      Width           =   6975
      Begin MSMask.MaskEdBox txtgroupname 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Enter Group Name"
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   35
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "c"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtgroupcode 
         Height          =   315
         Left            =   720
         TabIndex        =   0
         TabStop         =   0   'False
         Tag             =   "SKIP"
         ToolTipText     =   "Group Code"
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         MaxLength       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Name :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Code :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   240
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mode                As String
Dim PB_BlnkSysRegs      As Boolean
Dim PS_SysModules       As String
Dim Ls_InvBase As String
Dim PR_SysRegs          As Recordset

Private Sub Cmd_Discard_Click()
    ClearVal
    If PR_SysRegs.RecordCount > 0 Then
        PR_SysRegs.MoveFirst
        SetVal
        txtGroupCode.Enabled = True
    End If
    txtGroupCode.SetFocus
End Sub

Private Sub Cmd_Save_Click()
    If Mode = "A" And PR_SysRegs.RecordCount > 0 Then
        Call SetErr("Record already exists.", vbInformation)
    Else
        If ChkInputs = True Then SaveValues
        PR_SysRegs.Requery
        txtGroupCode.Enabled = False
        txtgroupname.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Set PR_SysRegs = New Recordset
    
    PR_SysRegs.Open "SELECT * FROM SysRegs", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_SysRegs.Close
End Sub

Public Sub SaveValues()
Dim cntsql As New ADODB.Command
PB_BlnkSysRegs = False

cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText

     Select Case Mode
           Case "A"
              cntsql.CommandText = "INSERT into SysRegs(GroupCode, GroupName, GroupAddr1, GroupAddr2, GroupCity, SysModules, AddDate, AddTime) VALUES ('" & txtGroupCode & "','" & txtgroupname & "','" & txtGroupaddr1 & "','" & txtgroupaddr2 & "','" & txtgroupcity & "','" & PS_SysModules & "'," & Date & ",'" & Time & "')"
              cntsql.Execute
           Case "E"
              cntsql.CommandText = "UPDATE SysRegs SET GroupCode = '" & txtGroupCode & "', GroupName = '" & txtgroupname & "',  GroupAddr1 = '" & txtGroupaddr1 & "', GroupAddr2 ='" & txtgroupaddr2 & "', GroupCity = '" & txtgroupcity & "', SysModules = '" & PS_SysModules & "' WHERE GroupCode = '" & txtGroupCode & "'"
              cntsql.Execute
           Case "D"
            cntsql.CommandText = "DELETE FROM SysRegs WHERE GroupCode = '" & txtGroupCode & "'"
            cntsql.Execute
           
           
     End Select
End Sub
Public Sub ClearVal()
    Dim ln_cnt As Integer
    
    txtGroupCode = ""
    txtgroupname = ""
    txtGroupaddr1 = ""
    txtgroupaddr2 = ""
    txtgroupcity = ""
    For ln_cnt = 0 To chkModules.UBound
        chkModules(ln_cnt) = False
    Next
End Sub

Private Sub SetVal()
    Dim ln_cnt As Integer
    
    txtGroupCode = PR_SysRegs("GroupCode")
    txtgroupname = Trim(PR_SysRegs("GroupName"))
    txtGroupaddr1 = Trim(PR_SysRegs("GroupAddr1") & "")
    txtgroupaddr2 = Trim(PR_SysRegs("GroupAddr2") & "")
    txtgroupcity = Trim(PR_SysRegs("GroupCity") & "")
    PS_SysModules = Trim(PR_SysRegs("SysModules") & "")
    
    For ln_cnt = 0 To chkModules.UBound
        chkModules(ln_cnt).Value = Val(Mid(PS_SysModules, ln_cnt + 1, 1))
    Next
End Sub
Public Function ChkInputs() As Boolean
    Dim ln_cnt As Integer
    
    PS_SysModules = ""
    For ln_cnt = 0 To chkModules.UBound
        PS_SysModules = PS_SysModules & chkModules(ln_cnt).Value
    Next
        
    If Len(txtGroupCode) > 0 And Len(txtgroupname) > 0 And Val(PS_SysModules) > 0 Then
       ChkInputs = True
    Else
       Call SetErr("Incomplete Data found", vbCritical)
       ChkInputs = False
    End If
End Function
Private Sub txtGroupaddr1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtgroupaddr2.SetFocus
End Sub

Private Sub txtgroupaddr2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtgroupcity.SetFocus
End Sub

Private Sub txtgroupcity_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then chkModules(0).SetFocus
End Sub

Private Sub txtgroupcode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtgroupcode_Validate False
        txtgroupname.SetFocus
    End If
End Sub

Private Sub txtgroupcode_Validate(Cancel As Boolean)
    If PR_SysRegs.RecordCount = 0 Then
        If MsgBox("Record does not exist. Want to create it?", vbQuestion + vbYesNo, "FSIB Financials") = vbYes Then
            txtgroupname.SetFocus
            Mode = "A"
        Else
            ClearVal
            txtGroupCode.SetFocus
        End If
    Else
        SetVal
        txtGroupCode.Enabled = False
        Mode = "E"
    End If
End Sub

Private Sub txtgroupname_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtGroupaddr1.SetFocus
End Sub

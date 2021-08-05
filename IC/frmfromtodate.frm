VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmfromtodate 
   Caption         =   "Users Log"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3600
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmfromtodate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1710
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "Description"
            TextSave        =   "Description"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   105833
            MinWidth        =   105833
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   60
      TabIndex        =   4
      Top             =   -45
      Width           =   3450
      Begin Crystal.CrystalReport crrpt 
         Left            =   135
         Top             =   225
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         DiscardSavedData=   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Frame Frame3 
         ForeColor       =   &H00000080&
         Height          =   1140
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   3435
         Begin MSComCtl2.DTPicker dtpfrom 
            Height          =   315
            Left            =   1680
            TabIndex        =   1
            Top             =   210
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   65077249
            CurrentDate     =   37293
         End
         Begin MSComCtl2.DTPicker DTPTo 
            Height          =   315
            Left            =   1680
            TabIndex        =   8
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   65077249
            CurrentDate     =   37293
         End
         Begin VB.Label txtselectiveaccount 
            Height          =   300
            Left            =   3855
            TabIndex        =   10
            Top             =   930
            Visible         =   0   'False
            Width           =   2520
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "To Date :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   915
            TabIndex        =   9
            Top             =   630
            Width           =   645
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "From Date :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   735
            TabIndex        =   6
            Top             =   240
            Width           =   825
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   2250
         TabIndex        =   3
         Top             =   1185
         Width           =   1035
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1140
         MaskColor       =   &H00000000&
         TabIndex        =   2
         Top             =   1185
         Width           =   1050
      End
      Begin VB.TextBox txtVchrDesc 
         Height          =   315
         Left            =   435
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmfromtodate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Pb_BlnkVchr As Boolean
Public PO_CODE As Object
Public PO_DESC As Object
Dim pr_dumy As New Recordset
Public codeid As String
Public Reporttype As String
Dim ls_sql As String





Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
On Error GoTo LocalErr

 
    With crrpt
        .SQLQuery = ""
        .ReportFileName = App.Path & Gs_ICRepoPath & "\Userlog.RPT"
        .WindowTitle = "" & Me.Caption & ""
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = 'Users Log'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .SelectionFormula = "{sys_passlog.addDate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {sys_passlog.addDate} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ")"

        .Connect = "DNS=Censoft;UID=Sa"
        
        .Action = 1
    End With
Exit Sub

LocalErr:

Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub


Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    DTPTo.SetFocus
End If
End Sub
Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    cmdGenerate.SetFocus
End If
End Sub

Private Sub Form_Load()
  
  dtpfrom = Date
  DTPTo = Date
  
End Sub


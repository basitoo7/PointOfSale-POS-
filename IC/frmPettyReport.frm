VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPettyreport 
   Caption         =   "Petty Cash"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4515
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmPettyReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   0
      TabIndex        =   3
      Top             =   -45
      Width           =   4410
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   405
         Left            =   2370
         TabIndex        =   5
         Top             =   1590
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
         Height          =   405
         Left            =   1275
         MaskColor       =   &H00000000&
         TabIndex        =   1
         Top             =   1590
         Width           =   1035
      End
      Begin Crystal.CrystalReport rptLedger 
         Left            =   3960
         Top             =   1725
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
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   1170
         TabIndex        =   0
         Top             =   285
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   37309
      End
      Begin MSComCtl2.DTPicker DTPto 
         Height          =   315
         Left            =   1170
         TabIndex        =   7
         Top             =   705
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   37309
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "To  Date :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   420
         TabIndex        =   6
         Top             =   750
         Width           =   690
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "From Date :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   285
         TabIndex        =   4
         Top             =   300
         Width           =   825
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2100
      Width           =   4515
      _ExtentX        =   7964
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
            Object.Width           =   10583
            MinWidth        =   10583
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
End
Attribute VB_Name = "frmPettyreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdGenerate_Click()
    
       With rptLedger
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "Pettycash.Rpt"
        .SelectionFormula = "{Pettycash.ValueDate} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {Pettycash.ValueDate} <= Date(" & DTPto.Year & "," & DTPto.Month & "," & DTPto.Day & ") "
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = 'Petty Cash'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPto & "'"
        .Action = 1
       End With
End Sub
Private Sub dtpfrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdGenerate.SetFocus
End Sub
Private Sub DTPto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdGenerate.SetFocus
End Sub
Private Sub Form_Load()
dtpfrom.Value = Date
DTPto.Value = Date
End Sub

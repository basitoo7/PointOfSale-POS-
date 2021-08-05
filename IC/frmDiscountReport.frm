VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDiscountReport 
   Caption         =   "Discount Report"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4200
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDiscountReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4200
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1725
      Width           =   4200
      _ExtentX        =   7408
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
      Height          =   1845
      Left            =   30
      TabIndex        =   5
      Top             =   -120
      Width           =   4125
      Begin VB.CheckBox ChkSummary 
         Caption         =   "Summary Only"
         Height          =   225
         Left            =   60
         TabIndex        =   10
         Top             =   1485
         Width           =   1425
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   2985
         TabIndex        =   4
         Top             =   1440
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
         Left            =   1905
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Top             =   1440
         Width           =   1035
      End
      Begin Crystal.CrystalReport crrpt 
         Left            =   -90
         Top             =   405
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
      Begin VB.TextBox txtVchrDesc 
         Height          =   315
         Left            =   3735
         MaxLength       =   50
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame Frame3 
         ForeColor       =   &H00000080&
         Height          =   1380
         Left            =   0
         TabIndex        =   6
         Top             =   30
         Width           =   4080
         Begin VB.ComboBox txtdiscperson 
            Height          =   330
            ItemData        =   "frmDiscountReport.frx":030A
            Left            =   1530
            List            =   "frmDiscountReport.frx":030C
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   945
            Width           =   2505
         End
         Begin MSComCtl2.DTPicker dtpfrom 
            Height          =   315
            Left            =   1530
            TabIndex        =   1
            Top             =   180
            Width           =   2085
            _ExtentX        =   3678
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
            CustomFormat    =   "dd-MM-yyyy HH:mm:ss"
            Format          =   54198273
            CurrentDate     =   37293
         End
         Begin MSComCtl2.DTPicker dtpto 
            Height          =   315
            Left            =   1530
            TabIndex        =   2
            Top             =   555
            Width           =   2100
            _ExtentX        =   3704
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
            CustomFormat    =   "dd-MM-yyyy HH:mm:ss"
            Format          =   54198273
            CurrentDate     =   37293
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Discount Person :"
            Height          =   210
            Left            =   225
            TabIndex        =   13
            Top             =   990
            Width           =   1980
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "To Date :"
            Height          =   210
            Left            =   855
            TabIndex        =   8
            Top             =   570
            Width           =   645
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "From Date :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   675
            TabIndex        =   7
            Top             =   195
            Width           =   825
         End
      End
   End
   Begin VB.ComboBox txtdiscperson1 
      Height          =   330
      ItemData        =   "frmDiscountReport.frx":030E
      Left            =   1560
      List            =   "frmDiscountReport.frx":0310
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   855
      Width           =   2505
   End
End
Attribute VB_Name = "frmDiscountReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pb_BlnkVchr As Boolean
Public PO_CODE As Object
Public PO_DESC As Object
Dim pr_dumy As New Recordset
Dim PR_Branch As New Recordset
Public codeid As String
Dim ls_sql As String
Dim ls_branchdesc As String

Private Sub Check1_Click()

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
'On Error GoTo LocalErr
MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
  With crrpt
        If ChkSummary.Value = 0 Then
            .ReportFileName = App.Path & Gs_ICRepoPath & "\DiscountReportDetail.RPT"
        Else
            .ReportFileName = App.Path & Gs_ICRepoPath & "\DiscountReport.RPT"
        End If
        
        .WindowTitle = Me.Caption
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = 'Discount Report'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .SQLQuery = "SELECT  SO_TransMaster.DiscAmount, SO_TransMaster.DiscBY, SO_AuthorityPerson.ACode, SO_AuthorityPerson.AName"
        .SQLQuery = .SQLQuery & " FROM  SO_TransMaster SO_TransMaster INNER JOIN   SO_AuthorityPerson SO_AuthorityPerson ON SO_TransMaster.DiscBY = SO_AuthorityPerson.ACode"
        .SQLQuery = .SQLQuery & " where  SO_TransMaster.Compcode = '" & Gs_compcode & "'"
        .SQLQuery = .SQLQuery & " and convert(varchar, SO_TransMaster.transdate,111) >= '" & Format(dtpfrom.Value, "YYYY/MM/DD") & "' "
        .SQLQuery = .SQLQuery & " and convert(varchar, SO_TransMaster.transdate,111) <= '" & Format(DTPTo.Value, "YYYY/MM/DD") & "' "
         If txtdiscperson1.Text <> "" Then
            .SQLQuery = .SQLQuery & " and SO_TransMaster.DiscBY = '" & txtdiscperson1.Text & "' "
         End If
        .SQLQuery = .SQLQuery & " ORDER BY SO_TransMaster.DiscBY"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1


    End With

   
MDIForm1.StatusBar1.Panels(7).Text = ""
Exit Sub

LocalErr:
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then DTPTo.SetFocus
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdGenerate.SetFocus
End Sub

Private Sub Form_Load()
  dtpfrom = Date
  DTPTo = Date
  LoadDiscountPersons
End Sub
Private Sub LoadDiscountPersons()
Dim pr_loadDiscperson As New Recordset
pr_loadDiscperson.Open "SELECT Acode ,Aname  from SO_AuthorityPerson where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_loadDiscperson.EOF Then
Do While Not pr_loadDiscperson.EOF
   txtdiscperson1.AddItem pr_loadDiscperson("Acode")
   txtdiscperson.AddItem pr_loadDiscperson("Aname")
pr_loadDiscperson.MoveNext
Loop
End If
pr_loadDiscperson.Close
End Sub

Private Sub txtdiscperson_Click()
txtdiscperson1.ListIndex = txtdiscperson.ListIndex
End Sub


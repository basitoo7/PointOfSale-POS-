VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmicreport11 
   Caption         =   "Expense Report"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmicreports11.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   7980
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1830
      Width           =   7980
      _ExtentX        =   14076
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
      Height          =   1935
      Left            =   30
      TabIndex        =   5
      Top             =   -120
      Width           =   7905
      Begin VB.CheckBox ChkSummary 
         Caption         =   "Summary"
         Height          =   225
         Left            =   135
         TabIndex        =   14
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   6810
         TabIndex        =   4
         Top             =   1515
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
         Left            =   5745
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Top             =   1515
         Width           =   1035
      End
      Begin Crystal.CrystalReport crrpt 
         Left            =   2895
         Top             =   270
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
         Left            =   2985
         MaxLength       =   50
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame Frame3 
         ForeColor       =   &H00000080&
         Height          =   1440
         Left            =   0
         TabIndex        =   6
         Top             =   30
         Width           =   7860
         Begin VB.TextBox txtexpensedesc 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   2490
            Locked          =   -1  'True
            MaxLength       =   64
            TabIndex        =   12
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   930
            Width           =   5310
         End
         Begin VB.TextBox txtexpensetype 
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "XXX"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   1530
            TabIndex        =   11
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   945
            Width           =   555
         End
         Begin VB.CommandButton Command6 
            Height          =   315
            Left            =   2145
            Picture         =   "frmicreports11.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   930
            Width           =   315
         End
         Begin MSComCtl2.DTPicker dtpfrom 
            Height          =   315
            Left            =   1530
            TabIndex        =   1
            Top             =   180
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
            Format          =   71696385
            CurrentDate     =   37293
         End
         Begin MSComCtl2.DTPicker dtpto 
            Height          =   315
            Left            =   1530
            TabIndex        =   2
            Top             =   540
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
            Format          =   71696385
            CurrentDate     =   37293
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Expense Code :"
            Height          =   255
            Left            =   150
            TabIndex        =   13
            Top             =   960
            Width           =   1350
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
End
Attribute VB_Name = "frmicreport11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pb_BlnkVchr As Boolean
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Dumy As New Recordset
Dim PR_Branch As New Recordset
Public codeid As String
Dim ls_sql As String
Dim ls_branchdesc As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub txtexpensetype_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtexpensetype) <> "" And KeyCode = vbKeyReturn Then
       txtexpensetype = DoPad(txtexpensetype, 3)
        PR_Dumy.Open "Select * from GL_ExenseType where Compcode  = '" & Gs_compcode & "' and Ecode = '" & txtexpensetype & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If PR_Dumy.EOF Then
            Call MsgBox("Expense Code not found !!!", vbCritical)
            txtexpensetype = ""
            txtexpensedesc = ""
            txtexpensetype.SetFocus
        Else
            txtexpensedesc = PR_Dumy("Description")
            
        End If
        PR_Dumy.Close
End If


End Sub
Private Sub Command6_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtexpensetype
    Set PO_DESC = txtexpensedesc
    
    Gs_SQL = "Select ECode,  Description from GL_ExenseType "
    Gs_FindFld = "Aname"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Expense Type"
    MyLookupOLDB.Show 1
    
    If txtexpensetype <> "" Then Call txtexpensetype_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub cmdGenerate_Click()
'On Error GoTo LocalErr
MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
    
 
    With crrpt
        If ChkSummary.Value = 1 Then
        .ReportFileName = App.Path & Gs_GlRepoPath & "\EmployeeExpenseSum.rpt"
        Else
        .ReportFileName = App.Path & Gs_GlRepoPath & "\EmployeeExpense.rpt"
        End If
        .WindowTitle = Me.Caption
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = 'Employee Expense Report'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & dtpto & "'"
        .SQLQuery = ""
        .RetrieveSQLQuery
        .SQLQuery = .SQLQuery & " where  GL_EmpExpense.Compcode = '" & Gs_compcode & "'"
        .SQLQuery = .SQLQuery & " and  GL_EmpExpense.transdate >= '" & Format(dtpfrom, "YYYY/MM/DD") & "'"
        .SQLQuery = .SQLQuery & " and  GL_EmpExpense.transdate <= '" & Format(dtpto, "YYYY/MM/DD") & "'"

        If txtexpensetype <> "" Then
        .SQLQuery = .SQLQuery & " and GL_EmpExpense.expcode = '" & txtexpensetype & "'"
        End If
        .Action = 1
    End With
   
   
MDIForm1.StatusBar1.Panels(7).Text = ""
Exit Sub

LocalErr:
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then dtpto.SetFocus
End Sub



Private Sub Form_Load()
 
  dtpfrom = Date
  dtpto = Date
 
End Sub


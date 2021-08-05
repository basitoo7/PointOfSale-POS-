VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmicreportReturn 
   Caption         =   "Sale Return Report"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmicreportsreturn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5610
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3180
      Width           =   5610
      _ExtentX        =   9895
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
      Height          =   3255
      Left            =   45
      TabIndex        =   5
      Top             =   -90
      Width           =   5520
      Begin VB.CheckBox ChkSummary 
         Caption         =   "Summary Only"
         Height          =   225
         Left            =   45
         TabIndex        =   10
         Top             =   2865
         Width           =   1425
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   4440
         TabIndex        =   4
         Top             =   2850
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
         Left            =   3360
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Top             =   2850
         Width           =   1035
      End
      Begin Crystal.CrystalReport crrpt 
         Left            =   135
         Top             =   195
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
         Height          =   2640
         Left            =   0
         TabIndex        =   6
         Top             =   30
         Width           =   5520
         Begin VB.CheckBox Chktime 
            Caption         =   "Sale With Time"
            Height          =   210
            Left            =   75
            TabIndex        =   19
            Top             =   930
            Width           =   1650
         End
         Begin VB.ComboBox txtcasher1 
            Height          =   330
            ItemData        =   "frmicreportsreturn.frx":030A
            Left            =   1530
            List            =   "frmicreportsreturn.frx":030C
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   2220
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
            Format          =   54657025
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
            Format          =   54657025
            CurrentDate     =   37293
         End
         Begin VB.ComboBox txtcasher 
            Height          =   330
            ItemData        =   "frmicreportsreturn.frx":030E
            Left            =   1530
            List            =   "frmicreportsreturn.frx":0310
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   2220
            Width           =   2505
         End
         Begin VB.Frame Frame2 
            Enabled         =   0   'False
            Height          =   1260
            Left            =   15
            TabIndex        =   14
            Top             =   915
            Width           =   5505
            Begin MSComCtl2.DTPicker DTPtimefrom 
               Height          =   315
               Left            =   1545
               TabIndex        =   15
               Top             =   345
               Width           =   1350
               _ExtentX        =   2381
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
               CustomFormat    =   "HH:mm:ss"
               Format          =   54657027
               CurrentDate     =   37293
            End
            Begin MSComCtl2.DTPicker DTPtimeto 
               Height          =   315
               Left            =   1545
               TabIndex        =   16
               Top             =   720
               Width           =   1365
               _ExtentX        =   2408
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
               CustomFormat    =   "HH:mm:ss"
               Format          =   54657027
               CurrentDate     =   37293
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "To Time :"
               Height          =   210
               Left            =   675
               TabIndex        =   18
               Top             =   735
               Width           =   825
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "From Time :"
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   675
               TabIndex        =   17
               Top             =   360
               Width           =   825
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Casher :"
            Height          =   210
            Left            =   870
            TabIndex        =   11
            Top             =   2250
            Width           =   615
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
Attribute VB_Name = "frmicreportReturn"
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

Private Sub Chktime_Click()
If Chktime.Value = 1 Then
Frame2.Enabled = True
    DTPtimefrom.SetFocus
Else
    Frame2.Enabled = False
End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
'On Error GoTo LocalErr
MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
  With crrpt
        If ChkSummary.Value = 1 Then
          .ReportFileName = App.Path & Gs_ICRepoPath & "\SaleReturnReportsummary.rpt"
        Else
            .ReportFileName = App.Path & Gs_ICRepoPath & "\SaleReturnReportAll.rpt"
        End If
          .SQLQuery = "SELECT IC_TransMaster.TransDate, IC_TransMaster.SaleStatus, IC_Clients.Description, IC_Trans.Quantity, IC_Trans.Amount, IC_Trans.DiscAmount,"
          .SQLQuery = .SQLQuery & " IC_Item.Description AS ItemDescription, IC_Item.CatCode, IC_ItemCategory.Description AS ItemCategory FROM SO_TransReturnMaster IC_TransMaster LEFT OUTER JOIN"
          .SQLQuery = .SQLQuery & " SO_TransReturn IC_Trans ON IC_TransMaster.Compcode = IC_Trans.Compcode AND IC_TransMaster.TransCode = IC_Trans.TransCode LEFT OUTER JOIN"
          .SQLQuery = .SQLQuery & " IC_Clients IC_Clients ON IC_TransMaster.Compcode = IC_Clients.Compcode AND  IC_TransMaster.AccountCode = IC_Clients.ClientCode LEFT OUTER JOIN"
          .SQLQuery = .SQLQuery & " IC_Item IC_Item ON IC_Trans.Compcode = IC_Item.Compcode AND IC_Trans.ItemCode = IC_Item.ItemCode LEFT OUTER JOIN    IC_ItemCategory IC_ItemCategory ON IC_Item.Compcode = IC_ItemCategory.Compcode AND IC_Item.CatCode = IC_ItemCategory.CatCode"
          .SQLQuery = .SQLQuery & " where IC_TransMaster.compcode = '" & Gs_compcode & "'  "
          .SQLQuery = .SQLQuery & " and convert(varchar,IC_TransMaster.transdate,111) >= '" & Format(dtpfrom.Value, "YYYY/MM/DD") & "' "
          .SQLQuery = .SQLQuery & " and convert(varchar,IC_TransMaster.transdate,111) <= '" & Format(DTPTo.Value, "YYYY/MM/DD") & "' "
          If Chktime.Value = 1 Then
            .SQLQuery = .SQLQuery & " and convert(varchar,Ic_TransMaster.transdate,108) >='" & Format(DTPtimefrom, "HH:mm:ss") & "' "
            .SQLQuery = .SQLQuery & " and convert(varchar,Ic_TransMaster.transdate,108) <='" & Format(DTPtimeto, "HH:mm:ss") & "' "
          End If
          
          .WindowTitle = Me.Caption
          .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
          .Formulas(1) = "Reportname = 'Sale Return Report'"
          .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
          
          
         If txtcasher.Text <> "" Then
         .SQLQuery = .SQLQuery & " and IC_TransMaster.usercode = " & txtcasher.Text & ""
         End If
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
If KeyCode = vbKeyReturn Then
 If Chktime.Value = 1 Then
    DTPtimefrom.SetFocus
 Else
    txtcasher1.SetFocus
 End If
End If
End Sub
Private Sub Form_Load()
  dtpfrom = Date
  DTPTo = Date
  LoadCasher
  DTPtimefrom = Time
  DTPtimeto = Time
  
End Sub
Private Sub LoadCasher()
Dim pr_loadcasher As New Recordset
pr_loadcasher.Open "SELECT ltrim(rtrim(UserCode)) as UserCode, ltrim(rtrim(UserName)) as UserName  from SyUsers where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_loadcasher.EOF Then
Do While Not pr_loadcasher.EOF
   txtcasher1.AddItem pr_loadcasher("Username")
   txtcasher.AddItem pr_loadcasher("UserCode")
pr_loadcasher.MoveNext
Loop
End If
pr_loadcasher.Close
End Sub

Private Sub txtcasher1_Click()
txtcasher.ListIndex = txtcasher1.ListIndex
End Sub

Private Sub txtcasher1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdGenerate.SetFocus
End Sub

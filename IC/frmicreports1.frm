VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmicreport1 
   Caption         =   "Sale Register Report"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmicreports1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5835
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2010
      Width           =   5835
      _ExtentX        =   10292
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
      Height          =   2130
      Left            =   30
      TabIndex        =   5
      Top             =   -120
      Width           =   5775
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   4680
         TabIndex        =   4
         Top             =   1725
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
         Left            =   3600
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Top             =   1725
         Width           =   1035
      End
      Begin Crystal.CrystalReport crrpt 
         Left            =   90
         Top             =   1770
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
         Left            =   1905
         MaxLength       =   50
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1725
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame Frame3 
         ForeColor       =   &H00000080&
         Height          =   1590
         Left            =   75
         TabIndex        =   6
         Top             =   90
         Width           =   5655
         Begin VB.OptionButton Option3 
            Caption         =   "Summary"
            Height          =   285
            Left            =   4335
            TabIndex        =   15
            Top             =   1215
            Width           =   1155
         End
         Begin VB.OptionButton Option2 
            Caption         =   "GRN Reference Summary"
            Height          =   285
            Left            =   1830
            TabIndex        =   14
            Top             =   1215
            Value           =   -1  'True
            Width           =   2340
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Detail With Items"
            Height          =   285
            Left            =   105
            TabIndex        =   13
            Top             =   1215
            Width           =   1500
         End
         Begin MSComCtl2.DTPicker dtpfrom 
            Height          =   315
            Left            =   1545
            TabIndex        =   2
            Top             =   780
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
            Format          =   62717953
            CurrentDate     =   37293
         End
         Begin VB.Frame Frame5 
            Height          =   585
            Left            =   30
            TabIndex        =   9
            Top             =   135
            Width           =   5580
            Begin VB.TextBox txtselectedcode 
               Height          =   315
               Left            =   1545
               MaxLength       =   6
               TabIndex        =   1
               Top             =   180
               Width           =   690
            End
            Begin VB.TextBox txtdesc 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Enabled         =   0   'False
               ForeColor       =   &H80000002&
               Height          =   315
               Left            =   2580
               MaxLength       =   50
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   180
               Width           =   2940
            End
            Begin VB.CommandButton Command5 
               Height          =   315
               Left            =   2235
               Picture         =   "frmicreports1.frx":030A
               Style           =   1  'Graphical
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   180
               Width           =   315
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Selected Cap :"
               Height          =   210
               Left            =   45
               TabIndex        =   12
               Top             =   225
               Width           =   1395
            End
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "As on Date :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   600
            TabIndex        =   7
            Top             =   810
            Width           =   900
         End
      End
   End
End
Attribute VB_Name = "frmicreport1"
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
'On Error GoTo LocalErr
MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
    With crrpt
        .WindowTitle = "" & Me.Caption & ""
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = '" & Me.Caption & "'"
        If Option2.Value = True Then
        .Formulas(1) = "ReportName = 'Vendor Payable (Summary)'"
        .ReportFileName = App.Path & Gs_ICRepoPath & "\AccounthistoryVendor.RPT"
        ElseIf Option3.Value = True Then
        .Formulas(1) = "ReportName = 'Vendor Payable (Summary)'"
        .ReportFileName = App.Path & Gs_ICRepoPath & "\AccountHistoryVendorSum.RPT"
        Else
        .Formulas(1) = "ReportName = 'Vendor Payable (Detail)'"
        .ReportFileName = App.Path & Gs_ICRepoPath & "\AccountHistoryVendorDetail.RPT"
        End If
        
        .Formulas(2) = "Period = '" & " As on date " & dtpfrom & "'"
        .Formulas(3) = "Currentdate = cdate(" & Format(dtpfrom, "yyyy,MM,dd") & ")"
        
        .SelectionFormula = "{IC_TransMaster.CompCode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " AND {IC_TransMaster.transdate} <= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ")"
        
        
        If Not Trim(txtselectedcode) = "" Then
        .SelectionFormula = .SelectionFormula & " AND {Ic_TransMaster.AccountCode} = '" & Trim(txtselectedcode) & "'"
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
Private Sub Command5_Click()

    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtselectedcode
    Set PO_DESC = txtdesc
    If Me.Caption = "Vendor Account History" Then
        Gs_SQL = "SELECT Suppliercode,Description  from IC_Supplier"
        Gs_FindFld = "Description"
        Gs_OrderBy = "Order by Description"
        Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
        MyLookupOLDB.Caption = "Suppliers"
    Else
        Gs_SQL = "SELECT clientcode,Description  from IC_clients"
        Gs_FindFld = "Description"
        Gs_OrderBy = "Order by Description"
        Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' "
        MyLookupOLDB.Caption = "Clients"
    End If
    
    MyLookupOLDB.Show 1
    SendKeys "{Tab}"
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub Form_Load()
  
  dtpfrom = Date
  
  
End Sub
Private Sub txtselectedcode_Validate(Cancel As Boolean)
If txtselectedcode <> "" Then
    txtselectedcode = DoPad(txtselectedcode, txtselectedcode.MaxLength)
    If Me.Caption = "Vendor Account History" Then
        ls_sql = "Select suppliercode,Description from Ic_Supplier where compcode = '" & Gs_compcode & "' and suppliercode = '" & txtselectedcode & "' "
    Else
        ls_sql = "Select clientcode,Description from Ic_clients where compcode = '" & Gs_compcode & "' and clientcode = '" & txtselectedcode & "' "
    End If
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("code not found", vbCritical)
                'Cancel = True
            Else
                txtdesc = pr_dumy("description")
            End If
         pr_dumy.Close

    
End If
End Sub

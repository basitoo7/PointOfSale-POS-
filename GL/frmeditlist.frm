VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmeditlist 
   Caption         =   "Edit List"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5085
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmeditlist.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5085
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   30
      TabIndex        =   1
      Top             =   -90
      Width           =   5025
      Begin VB.Frame Frame7 
         Caption         =   "Voucher Type"
         ForeColor       =   &H00000080&
         Height          =   990
         Left            =   90
         TabIndex        =   13
         Top             =   1440
         Width           =   4890
         Begin VB.TextBox txtvoucherno 
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "XXX"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   585
            Width           =   1140
         End
         Begin VB.TextBox txtvchrdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   2520
            MaxLength       =   50
            TabIndex        =   17
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   225
            Width           =   2295
         End
         Begin VB.TextBox txtVchrType 
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "XXX"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   225
            Width           =   615
         End
         Begin VB.CommandButton Command3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2190
            Picture         =   "frmeditlist.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   225
            Width           =   315
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Voucher # :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   660
            TabIndex        =   19
            Top             =   615
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Voucher Type :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   375
            TabIndex        =   16
            Top             =   270
            Width           =   1125
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Period"
         ForeColor       =   &H00000080&
         Height          =   1305
         Left            =   75
         TabIndex        =   4
         Top             =   150
         Width           =   4890
         Begin VB.CommandButton Command5 
            Height          =   315
            Left            =   2190
            Picture         =   "frmeditlist.frx":047C
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   195
            Width           =   315
         End
         Begin VB.TextBox txtbranchname 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   2535
            MaxLength       =   50
            TabIndex        =   5
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   195
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker dtpfrom 
            Height          =   315
            Left            =   1560
            TabIndex        =   7
            Top             =   555
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
            Format          =   54657025
            CurrentDate     =   37293
         End
         Begin MSComCtl2.DTPicker dtpto 
            Height          =   315
            Left            =   1560
            TabIndex        =   8
            Top             =   915
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
            Format          =   54657025
            CurrentDate     =   37293
         End
         Begin MSMask.MaskEdBox txtbranchcode 
            Height          =   315
            Left            =   1560
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Default Currency"
            Top             =   195
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "To Date :"
            Height          =   210
            Left            =   870
            TabIndex        =   12
            Top             =   960
            Width           =   645
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "From Date :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   690
            TabIndex        =   11
            Top             =   600
            Width           =   825
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Branch Code :"
            Height          =   210
            Left            =   480
            TabIndex        =   10
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   405
         Left            =   3930
         TabIndex        =   3
         Top             =   2445
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
         Left            =   2850
         MaskColor       =   &H00000000&
         TabIndex        =   2
         Top             =   2445
         Width           =   1035
      End
      Begin Crystal.CrystalReport rptEditList 
         Left            =   360
         Top             =   2460
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
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2775
      Width           =   5085
      _ExtentX        =   8969
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
End
Attribute VB_Name = "frmeditlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Pb_BlnkVchr As Boolean
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_VchType As Recordset
Dim PR_Branch As New Recordset


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
On Error GoTo LocalErr
Dim ls_Branch   As String
Dim ls_VchrDesc As String
'Dim ls_BranchName As String
MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
If MySeek(txtbranchcode, "BranchCode", PR_Branch) Then txtbranchname = PR_Branch("BranchDesc")
Select Case Me.Caption
Case "Edit List"
    
     ls_Branch = IIf(LTrim(RTrim(txtbranchcode)) = "", "", " And {Gl_Trans.BranchCode} = '" & txtbranchcode & "'")
    With rptEditList
        .ReportFileName = App.Path & Gs_GlRepoPath & "\EditList.RPT"
        .WindowTitle = "Edit List"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = 'Edit List'"
        .Formulas(2) = "BranchName = '" & txtbranchcode + "-" + txtbranchname & "'"
        .SelectionFormula = "{Gl_Trans.CompCode} = '" & Gs_compcode & "'" & ls_Branch
        .SelectionFormula = .SelectionFormula & " AND {Gl_Trans.Value_Date} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {Gl_Trans.Value_Date} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ")"
        If Not Trim(txtvchrType) = "" Then
            .SelectionFormula = .SelectionFormula & " AND {Gl_Trans.VchrType} = '" & Trim(txtvchrType) & "'"
        End If
        
        If Not Trim(txtvoucherno) = "" Then
            .SelectionFormula = .SelectionFormula & " AND {Gl_Trans.voucher_no} = '" & Trim(txtvoucherno) & "'"
        End If
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
    End With
Case Else
        If MySeek(Gs_BranchCode, "Branchcode", PR_Branch) Then txtbranchname = PR_Branch("BranchDesc")
        If MySeek(txtvchrType.Text, "VchrType", PR_VchType) Then txtVchrDesc = PR_VchType.Fields("Vchrdescrip")
         ls_Branch = IIf(LTrim(RTrim(txtbranchcode)) = "", "", " And {Gl_Trans.BranchCode} = '" & txtbranchcode & "'")
        With rptEditList
        .SelectionFormula = ""
        If Me.Caption = "Unposted Vouchers" Then
        .ReportFileName = App.Path & Gs_GlRepoPath & "\unpostVchr_Print.rpt"
        Else
        .ReportFileName = App.Path & Gs_GlRepoPath & "\Vchr_Print.rpt"
        End If
        .WindowTitle = "Print Voucher"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = '" & Trim(txtVchrDesc) & "'"
        .Formulas(2) = "Sig1 = '" & Gc_UserName & "'"
        .Formulas(3) = "Sig2 = '" & Gs_Sign2 & "'"
        .Formulas(4) = "Sig3 = '" & Gs_Sign3 & "'"
        .Formulas(5) = "branchName = '" & txtbranchcode + "-" + txtbranchname & "'"
        
        .SelectionFormula = "{Gl_Trans.CompCode} = '" & Gs_compcode & "'" & ls_Branch
        .SelectionFormula = .SelectionFormula & " AND {Gl_Trans.Value_Date} >= Date(" & dtpfrom.Year & "," & dtpfrom.Month & "," & dtpfrom.Day & ") AND {Gl_Trans.Value_Date} <= Date(" & DTPTo.Year & "," & DTPTo.Month & "," & DTPTo.Day & ")"
        If Not Trim(txtvchrType) = "" Then
            .SelectionFormula = .SelectionFormula & " AND {Gl_Trans.VchrType} = '" & Trim(txtvchrType) & "'"
        End If
        If Not Trim(txtvoucherno) = "" Then
            .SelectionFormula = .SelectionFormula & " AND {Gl_Trans.voucher_no} = '" & Trim(txtvoucherno) & "'"
        End If
        
        If Me.Caption = "Unposted Vouchers" Then
           .SelectionFormula = .SelectionFormula & " AND isnull({Gl_ref.pflag})"
        End If
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
    End With

End Select

MDIForm1.StatusBar1.Panels(7).Text = ""
Exit Sub

LocalErr:
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub Command3_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtvchrType
    Set PO_DESC = txtVchrDesc
    PR_VchType.Filter = "BranchCode = '" & txtbranchcode & "'"
    GoTop PR_VchType
    MyLookup.Caption = "Voucher Types"
    MyLookup.FillGrid PR_VchType, "VchrType", "VchrDescrip", 5
    MyLookup.Show 1
    PR_VchType.Filter = adFilterNone
    If Len(txtvchrType) > 0 Then txtVchrType_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Command5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbranchcode
    Set PO_DESC = txtbranchname
    
    GoTop PR_Branch
    MyLookup.Caption = "Company Branches"
    MyLookup.FillGrid PR_Branch, "BranchCode", "BranchDesc", txtbranchcode.MaxLength
    MyLookup.Show 1

    If Len(txtbranchcode) > 0 Then txtBranchCode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then DTPTo.SetFocus
If KeyCode = vbKeyPageUp Then txtbranchcode.SetFocus
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtvchrType.SetFocus
If KeyCode = vbKeyPageUp Then dtpfrom.SetFocus
End Sub

Private Sub Form_Load()
  
  Set PR_VchType = New Recordset
  PR_VchType.Open "Select *,BranchCode+VchrType as Findfld from GlVchrType where CompCode ='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_Branch.Open "Select * From SysBranch where compcode = '" & Gs_compcode & "' Order By BranchCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
  Pb_BlnkVchr = IIf(PR_VchType.EOF, True, False)
  dtpfrom = Date
  DTPTo = Date
  txtbranchcode = Gs_BranchCode
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_VchType.Close
    PR_Branch.Close
End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text2_Validate(Cancel As Boolean)

End Sub

Private Sub txtBranchCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) And txtbranchcode <> "" Then
     txtbranchcode = DoPad(txtbranchcode, txtbranchcode.MaxLength)
     
     If Not MySeek(txtbranchcode, "Branchcode", PR_Branch) Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtbranchcode.SetFocus
     Else
        txtbranchname = PR_Branch("BranchDesc")
        If dtpfrom.Enabled Then
            dtpfrom.SetFocus
        Else
          DTPTo.SetFocus
        End If
     End If
  ElseIf KeyCode = vbKeyF12 Then
     Command5_Click
  End If
End Sub
Private Sub txtVchrType_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 
 If Lastkey(KeyCode) And txtvchrType.Text <> "" Then
         txtvchrType = UCase(txtvchrType)
         lb_found = MySeek(Trim(txtbranchcode) + Trim(txtvchrType.Text), "Findfld", PR_VchType)
   
         If Not lb_found Then
            Call SetErr(Gs_RecNFMsg, vbCritical)
            txtvchrType.SetFocus
         Else
            StatusBar1.Panels(2) = PR_VchType.Fields("VchrType")
            cmdGenerate.SetFocus
         End If
  ElseIf Lastkey(KeyCode) And txtvchrType.Text = "" Then
         cmdGenerate.SetFocus
 ElseIf KeyCode = vbKeyF12 Then
      Command3_Click
 ElseIf KeyCode = vbKeyPageUp Then
        DTPTo.SetFocus
 End If
End Sub

Private Sub txtvoucherno_Validate(Cancel As Boolean)
txtvoucherno = DoPad(txtvoucherno, txtvoucherno.MaxLength)
End Sub

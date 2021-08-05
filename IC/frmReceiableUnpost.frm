VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReceiableUnpost 
   Caption         =   "Unpost Client Receipt"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4080
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
   Icon            =   "frmReceiableUnpost.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4080
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   30
      TabIndex        =   1
      Top             =   -90
      Width           =   4005
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "&Unpost"
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
         Left            =   1755
         MaskColor       =   &H00000000&
         TabIndex        =   11
         Top             =   1785
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   2865
         TabIndex        =   10
         Top             =   1785
         Width           =   1035
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1065
         Left            =   90
         TabIndex        =   5
         Top             =   135
         Width           =   3855
         Begin MSComCtl2.DTPicker dtpto 
            Height          =   315
            Left            =   1965
            TabIndex        =   6
            Top             =   645
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   16384001
            CurrentDate     =   37309
         End
         Begin MSComCtl2.DTPicker dtpfrom 
            Height          =   315
            Left            =   1965
            TabIndex        =   7
            Top             =   210
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   16384001
            CurrentDate     =   37309
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "From Date :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1020
            TabIndex        =   9
            Top             =   225
            Width           =   825
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "To Date :"
            Height          =   210
            Left            =   1200
            TabIndex        =   8
            Top             =   690
            Width           =   645
         End
      End
      Begin VB.TextBox txtAcctNarration 
         Height          =   315
         Left            =   3630
         MaxLength       =   50
         TabIndex        =   4
         Top             =   165
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Frame Frame7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   600
         Left            =   90
         TabIndex        =   2
         Top             =   1155
         Width           =   3855
         Begin VB.TextBox txtLocCode 
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
            Left            =   1950
            MaxLength       =   10
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   180
            Width           =   1470
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   3420
            Picture         =   "frmReceiableUnpost.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   180
            Width           =   315
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   -240
            MaxLength       =   50
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   435
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Reference # :"
            Height          =   210
            Left            =   870
            TabIndex        =   14
            Top             =   210
            Width           =   990
         End
      End
      Begin Crystal.CrystalReport rptLedger 
         Left            =   -15
         Top             =   2400
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
      Top             =   2190
      Width           =   4080
      _ExtentX        =   7197
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
Attribute VB_Name = "frmReceiableUnpost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PR_Item As New Recordset
Public PO_DESC As Object
Public PO_CODE As Object

Dim pr_dumy As New Recordset
Dim ls_sql As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdGenerate_Click()
On Error GoTo LocalErr
If txtLocCode <> "" Then

Dim pr_dumy1 As New Recordset
Dim ld_date As Date
Dim ls_vchrtype As String
Dim ls_voucherno As String
Dim ln_glstatus As Integer

pr_dumy1.Open "Select * from Ic_receiable where compcode = '" & Gs_compcode & "' and Transcode = '" & txtLocCode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy1.EOF Then
    ld_date = pr_dumy1("TransDate")
    ls_voucherno = Trim(pr_dumy1("VoucherNo") & "")
    ln_glstatus = Val(0 & pr_dumy1("glstatus"))
    
    If ls_voucherno = "" And ln_glstatus = 0 Then
        Call MsgBox("Voucher not posted", vbCritical)
        pr_dumy1.Close
        Exit Sub
    End If
    
    ls_vchrtype = Trim(pr_dumy1("Vchrtype") & "")
    
    ls_sql = "DELETE FROM Gl_Trans WHERE CompCode = '" & Gs_compcode & "' And BranchCode = '" & Gs_BranchCode & "' AND Voucher_No = '" & ls_voucherno & "' AND VchrType = '" & ls_vchrtype & "' and month(value_date) = " & Month(ld_date) & " and year(value_date) = " & Year(ld_date) & ""
    gc_dbcon.Execute ls_sql
    ls_sql = "DELETE FROM Gl_Ref WHERE CompCode = '" & Gs_compcode & "' And BranchCode = '" & Gs_BranchCode & "' AND Voucher_No = '" & ls_voucherno & "' AND VchrType = '" & ls_vchrtype & "' and month(value_date) = " & Month(ld_date) & " and year(value_date) = " & Year(ld_date) & ""
    gc_dbcon.Execute ls_sql

    ls_sql = "Update Ic_receiable set Glstatus = 0 where compcode = '" & Gs_compcode & "' and Transcode = '" & txtLocCode & "'"
    gc_dbcon.Execute ls_sql
    
    Call MsgBox("Receipts Reference Unpost Successfully", vbInformation)
    Call MsgBox("General Ledger Voucher Removed Successfully" + ls_vchrtype + "-" + ls_voucherno, vbInformation)
    
    txtLocCode = ""
Else
    Call MsgBox("Receipts Reference not found", vbInformation)
    txtLocCode = ""
End If

pr_dumy1.Close

Else
    Call MsgBox("Please Enter Receipts Reference", vbInformation)
    txtLocCode.SetFocus

End If

Exit Sub
LocalErr:
Call SetErr(Err.Description, vbCritical)
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub Command2_Click()
 Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtLocCode
    Set PO_DESC = Text1
    
        Gs_SQL = "Select Transcode, TransDate from IC_Receiable "
        Gs_FindFld = "TransCode"
        Gs_OrderBy = "Order by TransCode"
        Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and  Transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "' and Transdate <='" & Format(dtpto, "YYYY/MM/DD") & "'"
        MyLookupOLDB.Caption = "Client Receipts"
        MyLookupOLDB.Show 1
        If txtLocCode <> "" Then Call txtLocCode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtLocCode.SetFocus
End Sub

Private Sub txtLocCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 If KeyCode = vbKeyReturn And Len(txtLocCode.Text) > 0 Then
 txtLocCode = DoPad(txtLocCode, txtLocCode.MaxLength)
 If pr_dumy.State = 1 Then pr_dumy.Close
 ls_sql = "Select TransCode, TransDate from IC_Receiable where Compcode ='" & Gs_compcode & "'  and TransCode = '" & txtLocCode & "'"
 pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Client Receipt not found !!!", vbCritical)
            txtLocCode = ""
            Text1 = ""
            txtLocCode.SetFocus
        Else
         Text1 = pr_dumy("TransDate")
        End If
 pr_dumy.Close

ElseIf KeyCode = vbKeyF12 Then
        Command2_Click
End If

End Sub
Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then dtpto.SetFocus
End Sub
Private Sub Form_Load()
dtpfrom = Date
dtpto = Date
End Sub


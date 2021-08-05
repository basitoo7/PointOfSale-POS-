VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLeaseInfoStat 
   Caption         =   "Lease Info Status"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4830
   Icon            =   "frmLeaseInfoStat.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1605
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   4815
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3930
         TabIndex        =   18
         Top             =   1230
         Width           =   765
      End
      Begin VB.TextBox txtCustName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2670
         MaxLength       =   50
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   210
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         TabIndex        =   16
         Top             =   1230
         Width           =   765
      End
      Begin VB.CheckBox chkProvision 
         Height          =   315
         Left            =   3855
         TabIndex        =   13
         Top             =   900
         Width           =   315
      End
      Begin VB.CheckBox chkPortfolio 
         Height          =   315
         Left            =   3855
         TabIndex        =   11
         Top             =   570
         Width           =   315
      End
      Begin VB.CheckBox chkLegalStatus 
         Height          =   315
         Left            =   1575
         TabIndex        =   9
         Top             =   1200
         Width           =   315
      End
      Begin VB.CommandButton cmdLookup1 
         Height          =   315
         Left            =   2070
         Picture         =   "frmLeaseInfoStat.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   570
         Width           =   315
      End
      Begin VB.CheckBox chkReminder 
         Height          =   315
         Left            =   1575
         TabIndex        =   5
         Top             =   900
         Width           =   315
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   315
         Left            =   2310
         Picture         =   "frmLeaseInfoStat.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   210
         Width           =   315
      End
      Begin MSMask.MaskEdBox txtCustCode 
         Height          =   315
         Left            =   1575
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   210
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtLeaseNo 
         Height          =   315
         Left            =   1575
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   570
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox textx 
         Height          =   315
         Left            =   2430
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   570
         Visible         =   0   'False
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483644
         PromptInclude   =   0   'False
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Provision :"
         Height          =   195
         Left            =   3000
         TabIndex        =   14
         Top             =   930
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Portofolio Out :"
         Height          =   195
         Left            =   2685
         TabIndex        =   12
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Legal Status :"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   1230
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Reminder :"
         Height          =   195
         Left            =   690
         TabIndex        =   6
         Top             =   930
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Customer Code :"
         Height          =   195
         Left            =   285
         TabIndex        =   3
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lease No :"
         Height          =   195
         Left            =   675
         TabIndex        =   2
         Top             =   600
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmLeaseInfoStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkGls0 As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object

'Dim PR_LeaseInfo As Recordset
Dim PR_Customer As Recordset
Dim PR_LeaseInfo As Recordset
Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCustCode
    Set PO_DESC = txtCustName
    
    GoTop PR_Customer
    MyLookup.Caption = "Customer"
    MyLookup.FillGrid PR_Customer, "CustomerNo", "CustomerName", 6
    MyLookup.Show 1
    
    If Len(txtCustCode) > 0 Then txtCustCode_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub cmdLookup1_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtLeaseNo
    Set PO_DESC = textx
    
    GoTop PR_LeaseInfo
    MyLookup.Caption = "Lease Info"
    MyLookup.FillGrid PR_LeaseInfo, "LeaseNo", "LeaseRef", 3
    MyLookup.Show 1
    
    If Len(txtLeaseNo) > 0 Then txtLeaseNo_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Command1_Click()
   Call SaveValues
End Sub

Private Sub Command2_Click()
     txtCustCode = ""
     txtLeaseNo = ""
     chkProvision.Value = 0
     chkReminder.Value = 0
     chkLegalStatus.Value = 0
     chkPortfolio.Value = 0
     txtCustName = ""
End Sub

Private Sub Form_Load()
  
' Setting up Preveliges
  
  SetToolBar(1) = chkRights("CSTFACT001")
  SetToolBar(2) = chkRights("CSTFACT002")
  SetToolBar(3) = chkRights("CSTFACT003")
  SetToolBar(4) = chkRights("CSTFACT004")
  
  
  Set PR_LeaseInfo = New Recordset
  Set PR_Customer = New Recordset
   
  PR_Customer.Open "Select Customer.* from Customer Order By CustomerNo", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_LeaseInfo.Open "Select LM_LeaseInfo.* from LM_LeaseInfo where Compcode = '" & Gs_compcode & "' And BranchCode = '" & Gs_BranchCode & "' Order By LeaseNo", gc_dbcon, adOpenDynamic, adLockOptimistic, 1

  Call SetVal
  PB_BlnkGls0 = IIf(PR_LeaseInfo.EOF, True, False)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_LeaseInfo.Close
    PR_Customer.Close
End Sub

Public Sub SaveValues()
Dim cntsql As New ADODB.Command
PB_BlnkGls0 = False
Dim ln_ActiveStat As Integer

cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText

     If chkLegalStatus.Enabled = True Then
        ln_ActiveStat = 2
     Else
        ln_ActiveStat = 1
     End If
             
     cntsql.CommandText = "UPDATE LM_LeaseInfo SET ReminderStat= " & chkReminder & ", ProvisionStat = " & chkProvision & ", ActiveStatus = " & ln_ActiveStat & " WHERE  compcode = '" & Gs_compcode & "' And LeaseNo= '" & txtLeaseNo.Text & "' And CustomerNo= '" & txtCustCode.Text & "' And BranchCode= '" & Gs_BranchCode & "'"
     cntsql.Execute
           
     PR_LeaseInfo.Requery
     Call ClearVal
     
End Sub
Public Sub ClearVal()

     txtCustCode = ""
     txtLeaseNo = ""
     chkProvision.Value = 0
     chkReminder.Value = 0
     chkLegalStatus.Value = 0
     chkPortfolio.Value = 0
     
End Sub

Private Sub SetVal()
     
    If PR_LeaseInfo.RecordCount <> 0 And Not PR_LeaseInfo.EOF Then
      txtCustCode = PR_LeaseInfo("CustomerNo")
      txtLeaseNo = PR_LeaseInfo("LeaseNo")
      chkReminder.Value = PR_LeaseInfo("ReminderStat")
     
    Select Case PR_LeaseInfo("ActiveStatus")
       Case 0
         chkPortfolio.Value = 1
       Case 2
         chkLegalStatus.Value = 1
    End Select
      chkProvision.Value = PR_LeaseInfo("ProvisionStat")
    End If

End Sub

Public Function ChkInputs() As Boolean
    If Len(txtCustCode.Text) = txtRecCode.MaxLength And Len(txtLeaseNo.Text) = txtLeaseNo.MaxLength Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Private Sub txtCustCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
   
   If LastKey(KeyCode) And txtCustCode.Text <> "" Then
      txtCustCode.Text = UCase(txtCustCode.Text)
      lb_found = MySeek(txtCustCode.Text, "CustomerNo", PR_Customer)

        If Not lb_found Then
           Call SetErr(Gs_RecNFMsg, vbCritical)
        Else
           textx.Text = PR_Customer("CustomerName")
           txtLeaseNo.SetFocus
        End If
    ElseIf KeyCode = vbKeyF12 Then
       Call cmdLookup_Click
    ElseIf KeyCode = vbKeyPageUp Then
       txtLeaseNo.SetFocus
    End If
End Sub

Private Sub txtLeaseNo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
   
   If LastKey(KeyCode) And txtLeaseNo.Text <> "" Then
      txtLeaseNo.Text = UCase(txtLeaseNo.Text)
      lb_found = MySeek(txtLeaseNo.Text, "LeaseNo", PR_LeaseInfo)

        If Not lb_found Then
           Call SetErr(Gs_RecNFMsg, vbCritical)
        Else
'           textx.Text = PR_Customer("CustomerName")
           chkReminder.SetFocus
        End If
    ElseIf KeyCode = vbKeyF12 Then
       Call cmdLookup1_Click
    ElseIf KeyCode = vbKeyPageUp Then
       chkReminder.SetFocus
    End If
End Sub

Private Sub chkLegalStatus_Click()
  
  If chkLegalStatus.Value = 1 Then
     chkPortfolio.Enabled = False
     chkPortfolio.Value = 0
  Else
     chkPortfolio.Enabled = True
     chkLegalStatus.Enabled = False
     chkPortfolio.Value = 1
  End If

End Sub

Private Sub chkPortfolio_Click()
  
  If chkPortfolio.Value = 1 Then
     chkLegalStatus.Enabled = False
     chkLegalStatus.Value = 0
  Else
     chkLegalStatus.Enabled = True
     chkPortfolio.Enabled = False
     chkLegalStatus.Value = 1
  End If
End Sub

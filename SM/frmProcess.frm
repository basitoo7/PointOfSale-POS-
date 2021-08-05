VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcess 
   Caption         =   "Process"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   Icon            =   "frmProcess.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1605
      Left            =   0
      TabIndex        =   0
      Top             =   570
      Width           =   4935
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2205
         MaxLength       =   64
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   165
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Height          =   315
         Left            =   1860
         Picture         =   "frmProcess.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   165
         Width           =   315
      End
      Begin VB.CheckBox chkStatus 
         Height          =   315
         Left            =   1050
         TabIndex        =   8
         Top             =   1245
         Width           =   315
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   315
         Left            =   2265
         Picture         =   "frmProcess.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   525
         Width           =   315
      End
      Begin VB.TextBox txtDesc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   885
         Width           =   3795
      End
      Begin MSMask.MaskEdBox txtProcess 
         Height          =   315
         Left            =   1050
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   525
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         MaxLength       =   10
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
      Begin MSMask.MaskEdBox txtmodule 
         Height          =   315
         Left            =   1050
         TabIndex        =   11
         Tag             =   "SKIPN"
         Top             =   165
         Width           =   855
         _ExtentX        =   1508
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Sys. Module :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   12
         Top             =   195
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Status :"
         Height          =   195
         Left            =   465
         TabIndex        =   7
         Top             =   1275
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Process Id :"
         Height          =   195
         Left            =   165
         TabIndex        =   5
         Top             =   525
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   885
         Width           =   885
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   1005
      ButtonWidth     =   1217
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&New"
            Description     =   "Add"
            Object.ToolTipText     =   "Add new record"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Edit"
            Description     =   "Edit"
            Object.ToolTipText     =   "Edit an existing record"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Delete"
            Description     =   "Remove "
            Object.ToolTipText     =   "Remove an existing record."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Save"
            Description     =   "Save a new Record"
            Object.ToolTipText     =   "Save on disk"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Listing"
            Description     =   "Print Listing."
            Object.ToolTipText     =   "Print listing."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Description     =   "Find a Record."
            Object.ToolTipText     =   "Find a record."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancel"
            Description     =   "Cancel Operation"
            Object.ToolTipText     =   "Cancel operation mode"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   14
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4920
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProcess.frx":05EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProcess.frx":0A42
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProcess.frx":0E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProcess.frx":12EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProcess.frx":173E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProcess.frx":1B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProcess.frx":22E6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkGls0 As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_SyProc As Recordset
Dim PR_Module As New Recordset
Private Sub Command3_Click()
    Set PO_CODE = Nothing
    Set PO_DESC = Nothing
    Set PO_AnyForm = Nothing
    
    Set PO_AnyForm = Me
    Set PO_CODE = txtmodule
    Set PO_DESC = Text2
    
    GoTop PR_Module
    MyLookup.Caption = "Modules"
    MyLookup.FillGrid PR_Module, "IdCode", "IdDescrip", 5
    MyLookup.Show 1
    If Len(txtmodule.Text) > 0 Then txtmodule_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtprocess
    Set PO_DESC = txtdesc
    GoTop PR_SyProc
    PR_SyProc.Filter = "proctype = '" & txtmodule & "'"
    MyLookup.Caption = "Processes"
    MyLookup.FillGrid PR_SyProc, "ProcCode", "ProcDesc", 12
    MyLookup.Show 1
    If Len(txtprocess) > 0 Then txtprocess_KeyDown vbKeyReturn, vbKeyShift
    PR_SyProc.Filter = adFilterNone
End Sub

Private Sub Form_Load()
  
  
' Setting up Preveliges
  
  SetToolBar(1) = chkRights("SMPROCS001")
  SetToolBar(2) = chkRights("SMPROCS002")
  SetToolBar(3) = chkRights("SMPROCS003")
  SetToolBar(4) = chkRights("SMPROCS004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  Set PR_SyProc = New Recordset
   
  PR_SyProc.Open "Select * from SyProc", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_Module.Open "Select * from Fcm_Ids where Recid = 'MOD'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PB_BlnkGls0 = IIf(PR_SyProc.EOF, True, False)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_SyProc.Close
    PR_Module.Close
End Sub

Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtdesc = StrConv(txtdesc, 3)
End Sub

Private Sub txtmodule_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtmodule <> "" Then
        txtmodule = UCase(txtmodule)
      If MySeek(LTrim(RTrim(txtmodule.Text)), "Idcode", PR_Module) Then
        Text2 = PR_Module("IdDescrip")
        If txtprocess.Enabled Then txtprocess.SetFocus
      Else
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtmodule.SetFocus
      End If
End If
End Sub


Private Sub txtprocess_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn And Len(txtprocess.Text) > 0 Then
    PR_SyProc.Filter = "proctype = '" & txtmodule & "'"
         txtprocess.Text = UCase(txtprocess.Text)
         lb_found = MySeek(txtprocess.Text, "proccode", PR_SyProc)
         
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   ''Cancel = True
                   SetClear Me
                   txtprocess.SetFocus
                Else
                   txtdesc.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   ''Cancel = True
                   SetClear Me
                   txtprocess.SetFocus
                Else
                   Call SetVal
                   If Mode <> "D" Then
                      txtprocess.Enabled = False
                      txtdesc.SetFocus
                   End If
                End If
            End Select
       PR_SyProc.Filter = adFilterNone
       End If
       
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
       cmdLookup.Enabled = False
    ElseIf Range(Button.Index, 2, 3) Then
       cmdLookup.Enabled = True
    End If
    
    If PB_BlnkGls0 And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
      ' 'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_SyProc, frmProcess, txtmodule, txtdesc, "X", "CompCount", 3, "processid", "processDesc", 1, False, Toolbar1)
    End If
    
End Sub

Public Sub SaveValues()
On Error GoTo LocalErr
Dim cntsql As New ADODB.Command
PB_BlnkGls0 = False

cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText
gc_dbcon.BeginTrans
     Select Case Mode
           Case "A"
              cntsql.CommandText = "INSERT into syproc(procCode,procdesc,procStat,ProcType) VALUES ('" & txtprocess.Text & "','" & txtdesc.Text & "'," & chkStatus.Value & ",'" & txtmodule.Text & "')"
               cntsql.Execute
           Case "E"
              cntsql.CommandText = "UPDATE syproc SET procDesc= '" & txtdesc.Text & "', procStat='" & chkStatus.Value & "', proctype='" & txtmodule & "' WHERE  procCode= '" & txtprocess.Text & "'"
              txtprocess.Enabled = True
              cntsql.Execute
           Case "D"
              cntsql.CommandText = "DELETE FROM syproc WHERE procCode = '" & txtprocess.Text & "'"
              cntsql.Execute
           
     End Select
gc_dbcon.CommitTrans
     PR_SyProc.Requery

Exit Sub
LocalErr:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0

End Sub

Private Sub SetVal()
     txtdesc = PR_SyProc("procdesc")
     chkStatus.Value = PR_SyProc("procStat")
     txtmodule = PR_SyProc("proctype")
     If txtmodule <> "" Then Call txtmodule_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtprocess.Text) = txtprocess.MaxLength And txtdesc <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub FrmRefresh()
   PR_SyProc.Requery
End Sub

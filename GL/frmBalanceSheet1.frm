VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBalancesheet0 
   Caption         =   "Balance Sheet Main Heads Setup "
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4965
   Icon            =   "frmBalanceSheet1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1035
      Left            =   30
      TabIndex        =   1
      Top             =   570
      Width           =   4935
      Begin VB.CommandButton cmdLookup 
         Height          =   315
         Left            =   1860
         Picture         =   "frmBalanceSheet1.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   225
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
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   3495
      End
      Begin MSMask.MaskEdBox txtsub0 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Tag             =   "SKIP"
         Top             =   225
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         ForeColor       =   0
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code :"
         Height          =   195
         Left            =   765
         TabIndex        =   5
         Top             =   285
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         Height          =   195
         Left            =   345
         TabIndex        =   4
         Top             =   660
         Width           =   885
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   1005
      ButtonWidth     =   1217
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
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
               Picture         =   "frmBalanceSheet1.frx":047C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBalanceSheet1.frx":08D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBalanceSheet1.frx":0D24
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBalanceSheet1.frx":1178
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBalanceSheet1.frx":15CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBalanceSheet1.frx":1A20
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBalanceSheet1.frx":2174
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmBalancesheet0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkGls0 As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_GlBS01 As Recordset

Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtsub0
    Set PO_DESC = txtdesc
    GoTop PR_GlBS01
    MyLookup.Caption = "Balance Sheet Codes"
    MyLookup.FillGrid PR_GlBS01, "BCODE", "BDESC", 3
    MyLookup.Show 1
    
    If Len(txtsub0) > 0 Then
        txtsub0_Validate False
        SendKeys vbTab
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    ElseIf KeyCode = vbKeyF11 Then
        Mode = DentMode(Mode, 4, PR_GlBS01, Me, txtsub0, txtdesc, "X", "CompCount", 3, "Acct_sub0", "Acct_Desc", 1, False, Toolbar1)
    End If
End Sub

Private Sub Form_Load()
    SetToolBar(1) = chkRights("GLFRM10001")
    SetToolBar(2) = chkRights("GLFRM10002")
    SetToolBar(3) = chkRights("GLFRM10003")
    SetToolBar(4) = chkRights("GLFRM10004")
    
    Toolbar1.Buttons(1).Enabled = SetToolBar(1)
    Toolbar1.Buttons(2).Enabled = SetToolBar(2)
    Toolbar1.Buttons(3).Enabled = SetToolBar(3)
    Toolbar1.Buttons(5).Enabled = SetToolBar(4)
    
    
    Set PR_GlBS01 = New Recordset
     
    PR_GlBS01.Open "Select Gl_BSheet1.* from Gl_BSheet1 where compcode = '" & Gs_compcode & "' ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
     
    PB_BlnkGls0 = IIf(PR_GlBS01.EOF, True, False)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_GlBS01.Close
End Sub

Private Sub txtsub0_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        Call cmdLookup_Click
    ElseIf KeyCode = vbKeyReturn And Mode = "D" Then
        txtsub0_Validate True
    End If
End Sub
  
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If PB_BlnkGls0 And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_GlBS01, Me, txtsub0, txtdesc, "X", "CompCount", 3, "Acct_sub0", "Acct_Desc", 1, False, Toolbar1)
       If Mode = "A" Then
            cmdLookup.Enabled = False
        Else
            cmdLookup.Enabled = True
       End If
    End If
End Sub

Public Sub SaveValues()
'On Error GoTo LocalErr
PB_BlnkGls0 = False
gc_dbcon.BeginTrans
     Select Case Mode
           Case "A"
              gc_dbcon.Execute "INSERT INTO Gl_BSheet1(Compcode,BCODE, BDESC) VALUES ('" & Gs_compcode & "', '" & RepApp(Trim(txtsub0.Text)) & "','" & UCase(RepApp(Trim(txtdesc.Text))) & "')"
              
           Case "E"
              gc_dbcon.Execute "UPDATE Gl_BSheet1 SET BDESC = '" & UCase(RepApp(txtdesc.Text)) & "' WHERE BCODE = '" & RepApp(txtsub0.Text) & "' and compcode = '" & Gs_compcode & "'"
              
           Case "D"
              gc_dbcon.Execute "DELETE FROM Gl_BSheet1 WHERE BCODE = '" & RepApp(txtsub0.Text) & "' and compcode = '" & Gs_compcode & "'"
              
     End Select
gc_dbcon.CommitTrans
Call FrmRefresh

Exit Sub

LocalErr:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub SetVal()
     txtdesc = Trim(PR_GlBS01("BDESC"))
End Sub
Public Function ChkInputs() As Boolean
    If Len(Trim(txtsub0.Text)) > 3 Then
        MsgBox "Invalid Control Code!!!", vbInformation, App.ProductName
        txtsub0.SetFocus
        ChkInputs = False
        Exit Function
    End If
    
    If Trim(txtdesc) = "" Then
        MsgBox "Enter Description!!!", vbInformation, App.ProductName
        txtdesc.SetFocus
        ChkInputs = False
        Exit Function
    End If
    
    ChkInputs = True
End Function

Public Sub FrmRefresh()
    PR_GlBS01.Requery
End Sub

Private Sub txtsub0_Validate(Cancel As Boolean)

    Dim lb_found As Boolean
    
    If Trim(txtsub0) <> "" Then
        txtsub0.Text = DoPad(UCase(txtsub0.Text), txtsub0.MaxLength)
        lb_found = MySeek(txtsub0.Text, "BCODE", PR_GlBS01)
        
        Select Case Mode
            Case "A"
                If lb_found Then
                    Call SetErr(Gs_RecFdMsg, vbCritical)
                    'Cancel = True
                    SetClear Me
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   'Cancel = True
                   SetClear Me
                Else
                   Call SetVal
                End If
        End Select
    Else
        SetClear Me
    End If

End Sub

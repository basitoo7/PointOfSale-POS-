VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRecoverer 
   Caption         =   "Lease Management Recovery"
   ClientHeight    =   2280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4845
   Icon            =   "frmRecoverer.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1710
      Left            =   0
      TabIndex        =   0
      Top             =   570
      Width           =   4845
      Begin VB.TextBox txtDesig 
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   930
         Width           =   3225
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   570
         Width           =   3225
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   315
         Left            =   2040
         Picture         =   "frmRecoverer.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   210
         Width           =   315
      End
      Begin MSMask.MaskEdBox txtRecCode 
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   210
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
      Begin MSMask.MaskEdBox txtRecLimit 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Designation :"
         Height          =   195
         Left            =   615
         TabIndex        =   10
         Top             =   960
         Width           =   930
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Rec Limit :"
         Height          =   195
         Left            =   750
         TabIndex        =   5
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code :"
         Height          =   195
         Left            =   1035
         TabIndex        =   3
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         Height          =   195
         Left            =   615
         TabIndex        =   2
         Top             =   600
         Width           =   885
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4845
      _ExtentX        =   8546
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
         Left            =   4560
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
               Picture         =   "frmRecoverer.frx":047C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRecoverer.frx":08D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRecoverer.frx":0D24
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRecoverer.frx":1178
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRecoverer.frx":15CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRecoverer.frx":1A20
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRecoverer.frx":2174
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmRecoverer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkGls0 As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Recoverer As Recordset

Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtRecCode
    Set PO_DESC = txtDesc
    
    GoTop PR_Recoverer
    MyLookup.Caption = "Lease Management Recovery"
    MyLookup.FillGrid PR_Recoverer, "RecCode", "RecName", txtRecCode.MaxLength
    MyLookup.Show 1
    
    If Len(txtRecCode) > 0 Then txtRecCode_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Form_Load()
  
' Setting up Preveliges
  
  SetToolBar(1) = chkRights("LMRECOVER1")
  SetToolBar(2) = chkRights("LMRECOVER2")
  SetToolBar(3) = chkRights("LMRECOVER3")
  SetToolBar(4) = chkRights("LMRECOVER4")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  Set PR_Recoverer = New Recordset
   
  PR_Recoverer.Open "Select LM_Recoverer.* from LM_Recoverer Order By RecCode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1

  PB_BlnkGls0 = IIf(PR_Recoverer.EOF, True, False)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_Recoverer.Close
End Sub

Private Sub txtdesc_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
    txtDesig.SetFocus
 End If
End Sub

Private Sub txtDesig_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       txtRecLimit.SetFocus
    ElseIf KeyCode = vbKeyPageUp Then
       txtDesc.SetFocus
    End If
End Sub

Private Sub txtRecCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn And Len(txtRecCode.Text) > 0 Then
          
         txtRecCode.Text = DoPad(txtRecCode.Text, 3)
         lb_found = MySeek(txtRecCode.Text, "RecCode", PR_Recoverer)
         
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   Cancel = True
                    SetClear Me
                   txtRecCode.SetFocus
                Else
                   txtDesc.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   Cancel = True
                    SetClear Me
                   txtRecCode.SetFocus
                Else
                   Call SetVal
                   If Mode <> "D" Then
                      txtRecCode.Enabled = False
                      txtDesc.SetFocus
                   End If
                End If
            End Select
  ElseIf KeyCode = vbKeyF12 Then
     cmdLookup_Click
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
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_Recoverer, Me, txtRecCode, txtDesc, "X", "RecCode", 3, "RecCode", "RecName", 1, False, Toolbar1)
    End If
    
End Sub

Public Sub SaveValues()
Dim cntsql As New ADODB.Command
PB_BlnkGls0 = False

cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText

     Select Case Mode
           Case "A"
              cntsql.CommandText = "INSERT into LM_Recoverer(RecCode,RecName,RecLimit,RecDesig) VALUES ('" & txtRecCode.Text & "','" & Trim(txtDesc.Text) & "'," & Val(0 & txtRecLimit) & ",'" & Trim(txtDesig) & "')"
              cntsql.Execute
           Case "E"
              cntsql.CommandText = "UPDATE LM_Recoverer SET RecName= '" & Trim(txtDesc.Text) & "', RecLimit = " & txtRecLimit & ",RecDesig = '" & Trim(txtDesig) & "' WHERE  RecCode= '" & txtRecCode.Text & "'"
              txtRecCode.Enabled = True
              cntsql.Execute
           Case "D"
              cntsql.CommandText = "DELETE FROM LM_Recoverer WHERE RecCode = '" & Trim(txtRecCode.Text) & "'"
              cntsql.Execute
           
     End Select
     PR_Recoverer.Requery
End Sub

Private Sub SetVal()
     txtDesc = PR_Recoverer("RecName")
     txtRecLimit = Val(0 & PR_Recoverer("RecLimit"))
     txtDesig = Trim(PR_Recoverer("RecDesig") & "")
End Sub

Public Function ChkInputs() As Boolean
    If Len(txtRecCode.Text) = txtRecCode.MaxLength And txtDesc <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub FrmRefresh()
PR_Recoverer.Requery
End Sub

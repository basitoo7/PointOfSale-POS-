VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGlSub0 
   Caption         =   "Control Accounts"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   Icon            =   "frmGlSub0.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtsub0 
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   6
      Top             =   810
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   570
      Width           =   4935
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
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   600
         Width           =   3495
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   315
         Left            =   1815
         Picture         =   "frmGlSub0.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   225
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         Height          =   195
         Left            =   390
         TabIndex        =   4
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code :"
         Height          =   195
         Left            =   810
         TabIndex        =   2
         Top             =   240
         Width           =   465
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
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
               Picture         =   "frmGlSub0.frx":047C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlSub0.frx":08D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlSub0.frx":0D24
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlSub0.frx":1178
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlSub0.frx":15CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlSub0.frx":1A20
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlSub0.frx":2174
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmGlSub0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkGls0 As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_GlSub0 As Recordset

Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtsub0
    Set PO_DESC = txtdesc
    GoTop PR_GlSub0
    MyLookup.Caption = "Sub<0> Levels"
    MyLookup.FillGrid PR_GlSub0, "Acct_Sub0", "Acct_Desc", 5
    MyLookup.Show 1
    
    If Len(txtsub0) > 0 Then txtsub0_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then Mode = DentMode(Mode, 4, PR_GlSub0, frmGlSub0, txtsub0, txtdesc, "X", "CompCount", 3, "Acct_sub0", "Acct_Desc", 1, False, Toolbar1)
End Sub

Private Sub Form_Load()
  
  SetToolBar(1) = chkRights("GLSUB00001")
  SetToolBar(2) = chkRights("GLSUB00002")
  SetToolBar(3) = chkRights("GLSUB00003")
  SetToolBar(4) = chkRights("GLSUB00004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  
  Set PR_GlSub0 = New Recordset
   
  PR_GlSub0.Open "Select * from Gl_Sub0 where compcode ='" & Gs_compcode & "'", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
   
  PB_BlnkGls0 = IIf(PR_GlSub0.EOF, True, False)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_GlSub0.Close
End Sub
Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
   If Lastkey(KeyCode) Then txtdesc = StrConv(txtdesc, vbProperCase)
End Sub

Private Sub txtsub0_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn And Len(txtsub0.Text) > 0 And Val(txtsub0.Text) > 0 Then
          PR_GlSub0.Requery
          
         txtsub0.Text = DoPad(UCase(txtsub0.Text), gn_sublen(0))
         lb_found = MySeek(txtsub0.Text, "Acct_sub0", PR_GlSub0)
                  
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   'Cancel = True
                   SetClear Me
                   txtsub0.SetFocus
                Else
                   txtdesc.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   'Cancel = True
                   SetClear Me
                   txtsub0.SetFocus
                Else
                   Call SetVal
                   If Mode <> "D" Then
                      txtsub0.Enabled = False
                      txtdesc.SetFocus
                   End If
                End If
            End Select
    ElseIf KeyCode = vbKeyF12 Then
        cmdLookup_Click
    End If
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If PB_BlnkGls0 And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_GlSub0, frmGlSub0, txtsub0, txtdesc, "X", "CompCount", 3, "Acct_sub0", "Acct_Desc", 1, False, Toolbar1)
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
              cntsql.CommandText = "INSERT into Gl_sub0(compcode,Acct_sub0,Acct_desc,userid,adddate,addtime) VALUES ('" & Gs_compcode & "','" & txtsub0.Text & "','" & txtdesc.Text & "','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "')"
             ' MsgBox cntsql.CommandText
              cntsql.Execute
           Case "E"
              cntsql.CommandText = "UPDATE Gl_sub0 SET Acct_Desc= '" & txtdesc.Text & "' WHERE  compcode = '" & Gs_compcode & "' and Acct_sub0= '" & txtsub0.Text & "'"
              txtsub0.Enabled = True
              cntsql.Execute
           Case "D"
              cntsql.CommandText = "DELETE FROM Gl_sub0 WHERE Acct_sub0 = '" & txtsub0.Text & "' and compcode = '" & Gs_compcode & "'"
              cntsql.Execute
           
     End Select
gc_dbcon.CommitTrans
PR_GlSub0.Requery

Exit Sub

LocalErr:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub SetVal()
     txtdesc = Trim(PR_GlSub0("Acct_Desc"))
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtsub0.Text) = gn_sublen(0) And Len(RTrim(txtdesc)) > 0 Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub FrmRefresh()
   PR_GlSub0.Requery
End Sub
